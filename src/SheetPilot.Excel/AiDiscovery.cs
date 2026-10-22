using System;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private object CaptureCatalog(string scope)
        {
            dynamic app=application();object books=app.Workbooks;var catalog=new JArray();
            var root=new Snapshot{id=Guid.NewGuid().ToString("N"),scope=scope,book=app.ActiveWorkbook};
            if(root.book==null){Release(books);root.Dispose();throw new InvalidOperationException("请先打开工作簿。");}
            root.bookName=(string)((dynamic)root.book).Name;
            try
            {
                for(int i=1;i<=(int)((dynamic)books).Count;i++)
                {
                    object book=((dynamic)books).Item[i],sheets=null;
                    var child=new Snapshot{id=Guid.NewGuid().ToString("N"),scope=scope,book=book,bookName=(string)((dynamic)book).Name,parent=root.id};
                    snapshots.Add(child.id,child);
                    var entry=new JObject{["bookId"]=child.id,["snapshotId"]=child.id,["book"]=child.bookName,["active"]=ReferenceEquals(book,root.book),["readOnly"]=(bool)((dynamic)book).ReadOnly,["hasVba"]=(bool)((dynamic)book).HasVBProject};
                    try
                    {
                        sheets=((dynamic)book).Worksheets;var pages=new JArray();
                        for(int n=1;n<=(int)((dynamic)sheets).Count;n++)
                        {
                            object sheet=((dynamic)sheets).Item[n],used=null;
                            try {used=((dynamic)sheet).UsedRange;pages.Add(new JObject{["sheet"]=(string)((dynamic)sheet).Name,["visible"]=(int)((dynamic)sheet).Visible,["usedRange"]=(string)((dynamic)used).Address[false,false]});child.sheets.Add((string)((dynamic)sheet).Name);}
                            finally{Release(used);Release(sheet);}
                        }
                        entry["sheets"]=pages;
                    }
                    catch(Exception ex){entry["unavailable"]=ex.Message;}
                    finally{Release(sheets);}
                    catalog.Add(entry);
                }
                RetireOldContext();snapshots.Add(root.id,root);snapshotOrder.Enqueue(root.id);
                return new {snapshotId=root.id,book=root.bookName,scope,catalog,ranges=new object[0],scannedCells=0,metadataOnly=true,excelVersion=(string)app.Version,
                    workflow="Directory only, NO cell data scanned. All open books in this Excel process (including hidden worksheets and VBA) are the maximum scope. First review the ENTIRE catalog. auto: select relevant books/sheets based on the task and follow formula references across them. all: consider EVERY listed book and explicitly account for unavailable/irrelevant books. Call excel_context (mode=capture) with this directory snapshotId and bookId and optional sheet/address for a stable full local scan; omission captures all worksheets including hidden sheets. Use returned snapshotId for excel_read/find/statistics and all edits. Prefer full-scope local find/statistics over paging every cell; samples are not complete data. Follow nextOffset/hasMore to completion when inspecting records or VBA. Each catalog bookId is also its snapshotId for metadata/VBA tools immediately; excel_context mode=bind is optional. Use that ID for workbook.describe/vba.read, not the directory ID. Read vba.read per module with pagination when macros affect the task. Cite workbook+sheet+range and state any omissions. Newly opened or closed/replaced books require a fresh user context; never open disk files or other Excel processes."};
            }
            catch {foreach(var child in snapshots.Values.Where(s=>s.parent==root.id).ToArray()){snapshots.Remove(child.id);child.Dispose();}root.Dispose();throw;}
            finally{Release(books);}
        }
        private void RetireOldContext()
        {
            if(snapshotOrder.Count<8)return;
            string oldest=snapshotOrder.Dequeue();
            foreach(var s in snapshots.Values.Where(s=>s.id==oldest||s.parent==oldest).ToArray())
            {
                foreach(string key in proposals.Where(p=>p.Value.snapshot==s.id).Select(p=>p.Key).ToArray())proposals.Remove(key);
                if(undo?.Any(p=>p.snapshot==s.id)==true)continue;
                s.Dispose();snapshots.Remove(s.id);
            }
        }
        public async Task<object> ExtendedAsync(JObject request)
        {
            string operation=(string)request["operation"];
            if(operation!="context.capture"&&operation!="context.bind")return Extended(request);
            var root=Get((string)request["snapshotId"]);if(root.parent!=null)root=Get(root.parent);var o=(JObject)request["options"];
            if(root.scope!="auto"&&root.scope!="all"||root.parent!=null)throw new InvalidOperationException("仅自动/所有工作簿目录支持按需读取。");
            var target=Get(Required(o,"bookId"));
            if(target.parent!=root.id)throw new InvalidOperationException("目标不属于本次打开工作簿目录。");
            if(operation=="context.bind")return new{snapshotId=target.id,book=target.bookName,metadataOnly=true,note="No cells captured. Use context.capture before excel_read/find/statistics. VBA and live metadata tools available."};
            string name=(string)o["sheet"],address=(string)o["address"];
            if(address!=null&&name==null)throw new ArgumentException("指定 address 时必须指定 sheet。");
            if(address!=null)AiRange.Parse(address);
            dynamic book=target.book;object sheets=book.Worksheets;
            // Acquire a fresh reference owned by the new frozen snapshot.
            object books=((dynamic)application()).Workbooks,ownedBook=null;
            try{for(int i=1;i<=(int)((dynamic)books).Count;i++){object b=((dynamic)books).Item[i];if(ReferenceEquals(b,target.book)){ownedBook=b;break;}Release(b);}}
            finally{Release(books);}
            if(ownedBook==null){Release(sheets);throw new InvalidOperationException("工作簿已关闭。");}
            var captured=new Snapshot{id=Guid.NewGuid().ToString("N"),scope=root.scope,parent=root.id,book=ownedBook,bookName=target.bookName,dataCaptured=true};
            var changes=new ExcelEvents(application());bool changed=false;CancelCapture=false;
            try
            {
                changes.Change(()=>changed=true);
                for(int i=1;i<=(int)((dynamic)sheets).Count;i++)
                {
                    object sheet=((dynamic)sheets).Item[i],range=null;
                    try {if(name!=null&&(string)((dynamic)sheet).Name!=name)continue;if(address!=null)range=((dynamic)sheet).Range[address];await CaptureSheet(captured,(dynamic)sheet,range);}
                    finally{Release(range);Release(sheet);}
                }
                if(captured.sheets.Count==0)throw new ArgumentException("工作表不存在。");
                if(changed)throw new InvalidOperationException("读取期间发生编辑，请重新读取。");
                snapshots.Add(captured.id,captured);
                return new{snapshotId=captured.id,book=captured.bookName,ranges=captured.ranges,scannedCells=captured.scanned,nonEmptyCells=captured.count,hiddenSheetsIncluded=name==null,sample=captured.cells.Take(80).ToArray(),sampleTruncated=captured.count>80,statistics=Statistics(captured.cells),capturedAt=DateTime.UtcNow.ToString("o"),note="Complete frozen local scan of ranges; use this snapshotId for all subsequent data tools and edits. Sample alone is not the full data."};
            }
            catch{captured.Dispose();throw;}
            finally{changes.Dispose();Release(sheets);}
        }
    }
}
