using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed class SortSnapshot
    {
        internal string Book,Sheet,Address,Fingerprint,Table;
        internal int Row,Column,Rows,Columns;
        internal object[,] Values;
        internal List<int[]> Merges;
        internal string[] Formats;
        internal bool Date1904;
        public int RowCount => Rows;
        public int ColumnCount => Columns;
        public bool IsTable => Table!=null;
        public bool TableHasHeader { get; internal set; }
        public int MergeCount => Merges.Count;
        public string Description { get; internal set; }
        public string[] Headers { get; internal set; }
        public int InitialColumn { get; internal set; }
        public IReadOnlyList<string> Keys(int column,bool header)
        {
            if(column<0||column>=Columns)throw new InvalidOperationException("请选择排序依据列。");
            if(Table!=null&&header!=TableHasHeader)throw new InvalidOperationException("Excel 表格的标题行必须保留。");
            return Enumerable.Range(header?1:0,Rows-(header?1:0)).Select(i=>Convert.ToString(Values[i,column],CultureInfo.InvariantCulture)??"").ToArray();
        }
    }
    public sealed partial class ExcelGateway
    {
        private static object[,] SortMatrix(object raw,int rows,int columns)
        {
            var result=new object[rows,columns];if(raw is Array a){for(int r=0;r<rows;r++)for(int c=0;c<columns;c++)result[r,c]=SortValue(a.GetValue(r+a.GetLowerBound(0),c+a.GetLowerBound(1)));}else result[0,0]=SortValue(raw);return result;
        }
        private static object SortValue(object value) => value is int code && (code & unchecked((int)0xFFFF0000))==unchecked((int)0x800A0000)?new System.Runtime.InteropServices.ErrorWrapper(code):value;
        private static string SortFingerprint(dynamic range) => JsonConvert.SerializeObject(new object[]{(object)range.Formula,(object)range.Value2});
        public SortSnapshot CaptureSort()
        {
            CheckThread();dynamic app=application(),selected=app.Selection;
            if(selected==null||Microsoft.VisualBasic.Information.TypeName((object)selected).TrimStart('_')!="Range"||(int)selected.Areas.Count!=1)throw new InvalidOperationException("请在连续数据区域内选中一个单元格。");
            dynamic sheet=selected.Worksheet,book=sheet.Parent;
            if((bool)book.ReadOnly||(bool)sheet.ProtectContents)throw new InvalidOperationException("工作簿只读或工作表受保护，无法排序。");
            dynamic firstCell=selected.Cells[1,1];
            bool oneRecordCell=(double)selected.CountLarge==1||((bool)firstCell.MergeCells&&(string)selected.Address==(string)firstCell.MergeArea.Address);
            dynamic range=oneRecordCell?selected.CurrentRegion:selected;
            if((int)range.Rows.Count==(int)sheet.Rows.Count||(int)range.Columns.Count==(int)sheet.Columns.Count)range=app.Intersect(range,sheet.UsedRange);
            dynamic table=selected.Cells[1,1].ListObject;
            if(table!=null)range=table.Range;
            else if((bool)sheet.AutoFilterMode)
            {dynamic filter=sheet.AutoFilter;if(filter!=null&&app.Intersect(selected,filter.Range)!=null)range=filter.Range;}
            if(table!=null&&(bool)table.ShowTotals)range=range.Resize[(int)range.Rows.Count-1,(int)range.Columns.Count];
            int rows=range.Rows.Count,columns=range.Columns.Count;
            if(rows<2||rows>200001||(long)rows*columns>2000000)throw new InvalidOperationException("排序区域需有至少 2 行，最多 200,001 行 / 2,000,000 个单元格。");
            ValidateSortRange(book,sheet,range);
            var values=SortMatrix((object)range.Value2,rows,columns);
            var formats=new string[columns];for(int c=0;c<columns;c++)formats[c]=(object)range.Cells[2,c+1].Resize[rows-1,1].NumberFormat as string;
            return new SortSnapshot {Book=WorkbookKey(book),Sheet=SheetKey(sheet),Address=range.Address,Row=range.Row,Column=range.Column,Rows=rows,Columns=columns,Table=table==null?null:(string)table.Name,TableHasHeader=table!=null&&(bool)table.ShowHeaders,
                Values=values,Formats=formats,Date1904=(bool)book.Date1904,Merges=RecordMerges(range),Fingerprint=SortFingerprint(range),InitialColumn=Math.Max(0,Math.Min(columns-1,(int)app.ActiveCell.Column-(int)range.Column)),
                Headers=Enumerable.Range(0,columns).Select(c=>"第 "+(c+1)+" 列 · "+Convert.ToString(values[0,c],CultureInfo.InvariantCulture)).ToArray(),
                Description=(string)book.Name+" / "+(string)sheet.Name+" / "+(string)range.Address+" · "+rows+" 行 × "+columns+" 列。单个单元格扩展为连续区域；多格选区按所选范围；包含隐藏行。"};
        }
        private static void ValidateSortRange(dynamic book,dynamic sheet,dynamic range)
        {
            if((bool)book.ReadOnly||(bool)sheet.ProtectContents)throw new InvalidOperationException("工作簿只读或工作表受保护，无法排序。");
            if(!Equals((object)range.HasArray,false))throw new InvalidOperationException("排序区域包含数组公式，无法安全排序。");
            try{if(!Equals((object)range.HasSpill,false))throw new InvalidOperationException("排序区域包含动态数组，请先复制为值。");}
            catch(System.Runtime.InteropServices.COMException){/* Older Excel has no HasSpill. */}
            catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){ }
            if((int)sheet.PivotTables().Count>0)
                for(int i=1;i<=(int)sheet.PivotTables().Count;i++)
                    if(sheet.Application.Intersect(range,sheet.PivotTables(i).TableRange2)!=null)throw new InvalidOperationException("请在数据透视表的源数据中排序。");
        }
        public string ApplySort(SortSnapshot snapshot,int column,SortOptions options)
        {
            var multi=new MultiSortOptions {HasHeader=options.HasHeader,MatchCase=options.MatchCase,
                Levels=new List<SortLevel>{new SortLevel{Column=column,Mode=options.Natural?"natural":options.Alphabetical?"values":"custom",Descending=options.Descending,Rules=options.Rules,Sequences=options.Sequences}}};
            return ApplySort(snapshot,PlanSort(snapshot,multi));
        }
        private void SortRows(SortSnapshot snapshot,bool header,int[] ranks)
        {
            dynamic app=application(),book=FindBook(snapshot.Book),sheet=FindPage(book,snapshot.Sheet),range=sheet.Range[snapshot.Address];
            ValidateSortRange(book,sheet,range);SortBlocks.Inverse(ranks);
            dynamic records=sheet.Cells[snapshot.Row+(header?1:0),snapshot.Column].Resize[ranks.Length,snapshot.Columns];
            var merges=RecordMerges(records);
            // Every merged span must remain contiguous and in its original internal order.
            foreach(var m in merges)for(int i=1;i<m[2];i++)if(ranks[m[0]+i]!=ranks[m[0]]+i)throw new InvalidOperationException("排序计划会拆散合并记录，请重新读取并预览。");
            dynamic table=snapshot.Table==null?null:sheet.ListObjects[snapshot.Table],added=null;
            if(table!=null&&(header!=snapshot.TableHasHeader||(bool)table.ShowHeaders!=snapshot.TableHasHeader||(string)table.Range.Resize[snapshot.Rows,snapshot.Columns].Address!=snapshot.Address))throw new InvalidOperationException("Excel 表格范围已改变，请重新读取。");
            bool events=app.EnableEvents,screen=app.ScreenUpdating;object autoExpand=null;
            dynamic sort=table==null?sheet.Sort:table.Sort;
            bool moved=false,inserted=false,revealed=false;var opened=new List<int[]>();
            try
            {
                app.EnableEvents=false;app.ScreenUpdating=false;
                autoExpand=app.AutoCorrect.AutoExpandListRange;app.AutoCorrect.AutoExpandListRange=false;
                // Excel's table Sort.Apply silently ignores a hidden header row. Reveal
                // it only for this operation and restore the user's layout afterward.
                if(table!=null&&!header){table.ShowHeaders=true;revealed=true;if((int)table.DataBodyRange.Row!=snapshot.Row)throw new InvalidOperationException("显示表格标题会移动数据，请先在 Excel 中显示标题后重试。");}
                // Insert before the data, never borrow an adjacent column or erase its formatting.
                if(table!=null)added=table.ListColumns.Add(1);
                else {sheet.Columns[snapshot.Column].Insert();inserted=true;}
                records=sheet.Cells[snapshot.Row+(header?1:0),snapshot.Column+1].Resize[ranks.Length,snapshot.Columns];
                foreach(var m in merges){records.Cells[m[0]+1,m[1]+1].Resize[m[2],m[3]].UnMerge();opened.Add(m);}
                var keys=new object[ranks.Length,1];for(int i=0;i<ranks.Length;i++)keys[i,0]=ranks[i];
                dynamic key=sheet.Cells[snapshot.Row+(header?1:0),snapshot.Column].Resize[ranks.Length,1];key.Value2=keys;
                // Excel assigns a unique helper name, including when table headers are hidden.
                sort.SortFields.Clear();sort.SortFields.Add(key,0,1,Type.Missing,0);
                if(table==null)sort.SetRange(key.Resize[ranks.Length,snapshot.Columns+1]);
                // ListObject.Sort requires xlYes even with its header row hidden.
                sort.Header=table!=null?1:2;sort.MatchCase=false;sort.Orientation=1;sort.Apply();moved=true;
            }
            finally
            {
                try
                {
                    try{foreach(var m in opened){int row=moved?ranks[m[0]]:m[0];records.Cells[row+1,m[1]+1].Resize[m[2],m[3]].Merge();}}
                    finally{try{sort.SortFields.Clear();if(added!=null)added.Delete();else if(inserted)sheet.Columns[snapshot.Column].Delete();}finally{if(revealed)table.ShowHeaders=false;}}
                }
                finally{try{if(autoExpand!=null)app.AutoCorrect.AutoExpandListRange=autoExpand;}finally{try{app.EnableEvents=events;}finally{app.ScreenUpdating=screen;}}}
            }
            sheet.Activate();sheet.Range[snapshot.Address].Select();
        }
        // Zero-based relative row/column and height/width. Adaptive subdivision avoids
        // one COM call per cell on the common unmerged path.
        private static List<int[]> RecordMerges(dynamic records)
        {
            var result=new List<int[]>();if(Equals((object)records.MergeCells,false))return result;
            int rows=records.Rows.Count,columns=records.Columns.Count,firstRow=records.Row,firstCol=records.Column;
            var seen=new HashSet<string>();var watch=System.Diagnostics.Stopwatch.StartNew();
            Action<dynamic> scan=null;scan=block=>
            {
                if(watch.ElapsedMilliseconds>30000)throw new InvalidOperationException("合并布局检查超过 30 秒，请缩小选区。");
                if(Equals((object)block.MergeCells,false))return;
                dynamic first=block.Cells[1,1],area=first.MergeArea;
                if((bool)first.MergeCells&&(int)area.Row<=(int)block.Row&&(int)area.Column<=(int)block.Column&&
                    (int)area.Row+(int)area.Rows.Count>=(int)block.Row+(int)block.Rows.Count&&
                    (int)area.Column+(int)area.Columns.Count>=(int)block.Column+(int)block.Columns.Count)
                {
                    int r=(int)area.Row-firstRow,c=(int)area.Column-firstCol,h=area.Rows.Count,w=area.Columns.Count;
                    if(r<0||c<0||r+h>rows||c+w>columns)throw new InvalidOperationException("合并区域 "+(string)area.Address+" 超出排序边界。请选择完整合并记录，并检查表头选项。");
                    if(seen.Add(r+":"+c))result.Add(new[]{r,c,h,w});return;
                }
                int height=block.Rows.Count,width=block.Columns.Count;
                if(height>1){int half=height/2;scan(block.Resize[half,width]);scan(block.Offset[half,0].Resize[height-half,width]);}
                else if(width>1){int half=width/2;scan(block.Resize[1,half]);scan(block.Offset[0,half].Resize[1,width-half]);}
            };scan(records);
            return result.OrderBy(m=>m[0]).ThenBy(m=>m[1]).ToList();
        }
        public string SortSheets(IReadOnlyList<SheetItem> expected,SortOptions options)
        {
            CheckThread();if(expected.Count==0)return "没有工作表。";
            var order=CustomSort.Order(expected.Select(s=>s.Name).ToArray(),options);
            var desired=order.Select(i=>expected[i]).ToArray();ReorderSheets(expected,desired);
            undoInvalidated=false;undo=()=>{ReorderSheets(desired,expected);return "已恢复工作表原始顺序。";};return "已排序 "+expected.Count+" 个工作表标签；可通过 GridPilot 撤销。";
        }
        private void ReorderSheets(IReadOnlyList<SheetItem> expected,IReadOnlyList<SheetItem> desired)
        {
            dynamic book=FindBook(expected[0].WorkbookKey),app=application();
            if((bool)book.ReadOnly||(bool)book.ProtectStructure)throw new InvalidOperationException("工作簿只读或结构已保护。");
            var live=Capture(true).Sheets.Where(s=>s.WorkbookKey==expected[0].WorkbookKey).ToArray();
            if(live.Length!=expected.Count||live.Where((s,i)=>s.Key!=expected[i].Key||s.Name!=expected[i].Name||s.Visibility!=expected[i].Visibility).Any())throw new InvalidOperationException("工作表列表在预览后已变化，请重新读取。");
            bool screen=app.ScreenUpdating,events=app.EnableEvents;dynamic active=app.ActiveSheet;
            Action<IReadOnlyList<SheetItem>> move=items=>{for(int i=items.Count-1;i>=0;i--){dynamic page=FindPage(book,items[i].Key);if((int)page.Index!=1)page.Move(book.Sheets[1]);}};
            try{app.ScreenUpdating=false;app.EnableEvents=false;
                // Excel ignores a hidden Before anchor. Make anchors visible only while moving,
                // then restore Hidden / VeryHidden exactly before repainting.
                foreach(var item in expected)if(!item.CanNavigate)((dynamic)FindPage(book,item.Key)).Visible=-1;
                try{move(desired);}catch{move(expected);throw;}}
            finally{try{foreach(var item in expected)if(!item.CanNavigate)((dynamic)FindPage(book,item.Key)).Visible=item.Visibility==SheetVisibility.VeryHidden?2:0;
                active?.Activate();}finally{app.EnableEvents=events;app.ScreenUpdating=screen;}}
        }
    }
}
