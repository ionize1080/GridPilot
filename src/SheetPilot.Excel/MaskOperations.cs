using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class ExcelGateway
    {
        internal static int LastMaskWriterProcess;
        internal static string LastMaskPerformance;
        public MaskSnapshot CaptureMaskWorkbook(CancellationToken cancel=default(CancellationToken),Action<string,int,long,long> progress=null)
        {
            CheckThread();dynamic app=application(),book=app.ActiveWorkbook;
            if(book==null)throw new InvalidOperationException("请先打开需要脱敏的工作簿。");
            if((int)app.CalculationState!=0)throw new InvalidOperationException("请等待 Excel 完成计算后再分析。");
            var snapshot=new MaskSnapshot {WorkbookKey=WorkbookKey(book),Date1904=(bool)book.Date1904};
            GetWindowThreadProcessIdForMask(new IntPtr((int)app.Hwnd),out snapshot.HostProcessId);
            dynamic sheets=book.Worksheets;int total=sheets.Count;long count=0;
            for(int i=1;i<=total;i++)
            {
                cancel.ThrowIfCancellationRequested();dynamic sheet=sheets[i];var spec=new DataSheet {Name=(string)sheet.Name};snapshot.Sheets.Add(spec);
                var rows=new HashSet<int>();var columns=new HashSet<int>();int firstRow=int.MaxValue;
                ScanMaskSheet((object)sheet,i-1,cancel,cell=>
                {
                    snapshot.Cells.Add(cell);rows.Add(cell.Row);columns.Add(cell.Column);firstRow=Math.Min(firstRow,cell.Row);
                    spec.Rows=Math.Max(spec.Rows,cell.Row);spec.Columns=Math.Max(spec.Columns,cell.Column);spec.NonEmptyCells++;count++;
                    if((count&32767)==0)CheckMaskMemory();
                },n=>progress?.Invoke("扫描「"+spec.Name+"」",(i-1)*100/total,count,0));
                spec.NonEmptyRows=rows.Count;spec.NonEmptyColumns=columns.Count;
                if(rows.Count>0)
                {
                    var merged=new HashSet<string>();
                    foreach(var span in columns.GroupBy(c=>(c-1)/256).OrderBy(g=>g.Key))
                    {cancel.ThrowIfCancellationRequested();VisitMerges((object)sheet,new CellRect(firstRow,span.Min(),Math.Min(40,spec.Rows-firstRow+1),span.Max()-span.Min()+1),m=>
                        {if(merged.Add(m.Address))spec.HeaderMerges.Add(new HeaderMerge {Row=m.Row,Column=m.Column,Rows=m.Rows,Columns=m.Columns});},(stage,done)=>cancel.ThrowIfCancellationRequested());}
                }
                progress?.Invoke("已扫描「"+spec.Name+"」",i*100/total,count,0);
            }
            if(total==0)throw new InvalidOperationException("工作簿没有普通工作表。");
            snapshot.Cells.Sort(CompareMaskCell);return snapshot;
        }
        private static void CheckMaskMemory()
        {
            // Leave headroom for the output snapshot, formula caches, Excel and the writer.
            long budget=Environment.Is64BitProcess?4L*1024*1024*1024:450L*1024*1024;
            if(GC.GetTotalMemory(false)>budget)throw new InvalidOperationException("已接近当前进程的安全内存预算。未生成部分结果；请使用64位 Excel 或拆分工作簿。");
        }
        public static IReadOnlyList<MaskRule> SuggestMaskRules(MaskSnapshot snapshot)
        {
            return MaskInference.Suggest(snapshot);
        }
        public void VerifyMaskSource(MaskSnapshot source,CancellationToken cancel=default(CancellationToken),Action<string,int,long,long> progress=null)
        {
            CheckThread();dynamic app=application(),book=app.ActiveWorkbook;
            if(book==null || WorkbookKey(book)!=source.WorkbookKey || (bool)book.Date1904!=source.Date1904 || (int)book.Worksheets.Count!=source.Sheets.Count || (int)app.CalculationState!=0)
                throw new InvalidOperationException("源工作簿已切换或改变，请重新分析。");
            long count=0;
            for(int i=0;i<source.Sheets.Count;i++)
            {
                dynamic sheet=book.Worksheets[i+1];
                if((string)sheet.Name!=source.Sheets[i].Name)throw new InvalidOperationException("源工作表已改变，请重新分析。");
                ScanMaskSheet((object)sheet,i,cancel,cell=>
                {var expected=FindMaskCell(source.Cells,cell);if(expected==null || !Equals(expected.Value,cell.Value) || expected.Formula!=cell.Formula)throw new InvalidOperationException("源工作簿在预览后改变，请重新分析。");count++;},n=>progress?.Invoke("复核源工作簿",(int)Math.Min(100,100*count/Math.Max(1,source.Cells.Count)),count,source.Cells.Count));
            }
            if(count!=source.Cells.Count)throw new InvalidOperationException("源工作簿有数据被删除，请重新分析。");
        }
        private string PrepareMaskExport(MaskSnapshot source,string target,bool preserveFormatting,CancellationToken cancel,Action<string,int,long,long> progress)
        {
            CheckThread();VerifyMaskSource(source,cancel,progress);cancel.ThrowIfCancellationRequested();
            if(!string.Equals(Path.GetExtension(target),".xlsx",StringComparison.OrdinalIgnoreCase) || File.Exists(target))throw new InvalidOperationException("请选择尚不存在的 .xlsx 文件；不覆盖已有文件。");
            string directory=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"GridPilot","Staging",Guid.NewGuid().ToString("N"));Directory.CreateDirectory(directory);
            try
            {
                if(preserveFormatting)
                {dynamic book=((dynamic)application()).ActiveWorkbook;string extension=Path.GetExtension((string)book.Name);if(string.IsNullOrEmpty(extension))extension=".xlsx";book.SaveCopyAs(Path.Combine(directory,"appearance-source"+extension));}
                return directory;
            }
            catch{CleanupMaskDirectory(directory);throw;}
        }
        private string CommitMaskExport(MaskSnapshot source,string stage,string target,CancellationToken cancel,Action<string,int,long,long> progress)
        {
            CheckThread();VerifyMaskSource(source,cancel,progress);cancel.ThrowIfCancellationRequested();
            string destination=Path.Combine(Path.GetDirectoryName(target),".GridPilot-"+Guid.NewGuid().ToString("N")+".tmp");
            try{File.Copy(stage,destination,false);cancel.ThrowIfCancellationRequested();File.Move(destination,target);progress?.Invoke("导出完成",100,source.Cells.Count,source.Cells.Count);return target;}
            finally{if(File.Exists(destination))File.Delete(destination);}
        }
        private static Task<string> StartMaskWriter(MaskSnapshot source,MaskResult result,string directory,CancellationToken cancel,bool preserveFormatting,Action<string,int,long,long> progress)
        {
            var completion=new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            var thread=new Thread(()=>{try{completion.SetResult(ExportMaskWorker(source,result,directory,cancel,preserveFormatting,progress));}catch(OperationCanceledException){completion.SetCanceled();}catch(Exception ex){completion.SetException(ex);}}){IsBackground=true,Name="GridPilot export STA"};
            thread.SetApartmentState(ApartmentState.STA);thread.Start();return completion.Task;
        }
        public async Task<string> ExportMaskAsync(MaskSnapshot source,MaskResult result,string target,CancellationToken cancel,bool preserveFormatting=false,Action<string,int,long,long> progress=null)
        {
            target=Path.GetFullPath(target);string directory=PrepareMaskExport(source,target,preserveFormatting,cancel,progress);
            try{string stage=await StartMaskWriter(source,result,directory,cancel,preserveFormatting,progress);return CommitMaskExport(source,stage,target,cancel,progress);}
            finally{CleanupMaskDirectory(directory);}
        }
        public string ExportMask(MaskSnapshot source,MaskResult result,string target,CancellationToken cancel,bool preserveFormatting=false)
        {
            target=Path.GetFullPath(target);string directory=PrepareMaskExport(source,target,preserveFormatting,cancel,null);
            try
            {
                var task=StartMaskWriter(source,result,directory,cancel,preserveFormatting,null);
                // Synchronous API for native callers/tests still pumps the host STA for OLE.
                while(!task.IsCompleted){System.Windows.Forms.Application.DoEvents();Thread.Sleep(10);}
                return CommitMaskExport(source,task.GetAwaiter().GetResult(),target,cancel,null);
            }
            finally{CleanupMaskDirectory(directory);}
        }
        private static void CleanupMaskDirectory(string directory)
        {try{foreach(string path in Directory.GetFiles(directory))File.Delete(path);Directory.Delete(directory);}catch(IOException){}catch(UnauthorizedAccessException){}}
        private static string ExportMaskWorker(MaskSnapshot source,MaskResult result,string stageDirectory,CancellationToken cancel,bool preserveFormatting,Action<string,int,long,long> progress)
        {
            cancel.ThrowIfCancellationRequested();
            string stage=Path.Combine(stageDirectory,"result.xlsx");
            var lifetime=new MaskComLifetime();dynamic writer=null,book=null; int process=0;
            var clock=System.Diagnostics.Stopwatch.StartNew();var stages=new List<string>();Action<string> timing=stage=>stages.Add(stage+"="+clock.ElapsedMilliseconds+"ms");
            try
            {
                // A separate owned Excel instance avoids recalculating or changing the source workbook.
                writer=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                GetWindowThreadProcessIdForMask(new IntPtr((int)writer.Hwnd),out process);
                if(process==source.HostProcessId){Marshal.ReleaseComObject((object)writer);writer=null;throw new InvalidOperationException("无法建立独立导出进程。");}
                lifetime.Hold((object)writer);
                writer.Visible=false;writer.DisplayAlerts=false;writer.EnableEvents=false;writer.AutomationSecurity=3;
                LastMaskWriterProcess=process;
                timing("writer");
                dynamic books=lifetime.Hold(writer.Workbooks);book=lifetime.Hold(books.Add(-4167));book.Date1904=source.Date1904;
                writer.Calculation=-4135;writer.CalculateBeforeSave=false;try{writer.PrintCommunication=false;}catch(COMException){}
                progress?.Invoke("写入脱敏副本",5,0,source.Cells.Count);long written=0;
                timing("blank");
                dynamic sheets=lifetime.Hold(book.Worksheets);
                for(int s=0;s<source.Sheets.Count;s++)
                { dynamic sheet=lifetime.Hold(s==0?sheets[1]:sheets.Add(Type.Missing,lifetime.Hold(sheets[sheets.Count])));sheet.Name="Sheet"+(s+1); }
                for(int s=0;s<source.Sheets.Count;s++)
                {
                    cancel.ThrowIfCancellationRequested();dynamic sheet=lifetime.Hold(sheets[s+1]),all=lifetime.Hold(sheet.Cells);var spec=source.Sheets[s];
                    foreach(var group in result.Output.Cells.Where(c=>c.Sheet==s).GroupBy(c=>(c.Row-1)/128))
                    {
                        using(var scope=lifetime.Scope())
                        {
                        var cells=group.ToArray();int top=cells.Min(c=>c.Row),bottom=cells.Max(c=>c.Row),left=cells.Min(c=>c.Column),right=cells.Max(c=>c.Column);
                        // One row block is at most 128 x 16384; split columns too for sparse workbooks.
                        for(int start=left;start<=right;start+=128)
                        {
                            int end=Math.Min(right,start+127);var subset=cells.Where(c=>c.Column>=start && c.Column<=end).ToArray();if(subset.Length==0)continue;
                            var data=new object[bottom-top+1,end-start+1];
                            foreach(var cell in subset)if(cell.Formula==null)data[cell.Row-top,cell.Column-start]=cell.Value is string text?"'"+text:cell.Value;
                            lifetime.Hold(sheet.Range[lifetime.Hold(all[top,start]),lifetime.Hold(all[bottom,end])]).Value2=data;
                        }
                        foreach(var cell in cells.Where(c=>c.Formula!=null)){cancel.ThrowIfCancellationRequested();lifetime.Hold(all[cell.Row,cell.Column]).Formula=cell.Formula;}
                        foreach(var cell in cells.Where(c=>c.NumberFormat!=null))lifetime.Hold(all[cell.Row,cell.Column]).NumberFormat=cell.NumberFormat;
                        written+=cells.Length;progress?.Invoke("写入脱敏副本",5+(int)(45*written/Math.Max(1,source.Cells.Count)),written,source.Cells.Count);
                        }
                    }
                    lifetime.Hold(sheet.Columns).ColumnWidth=20;
                }
                timing("data");progress?.Invoke("重算并验证副本",55,source.Cells.Count,source.Cells.Count);
                if(result.Formulas>0)writer.CalculateFullRebuild();
                timing("calculate");
                ValidateMaskBook(book,result,cancel,lifetime);
                timing("write+validate");
                dynamic properties=lifetime.Hold(book.BuiltinDocumentProperties);
                foreach(string property in new[]{"Author","Last Author","Title","Subject","Comments","Keywords","Category","Company","Manager"}) try { lifetime.Hold(properties[property]).Value=""; }catch(COMException) { }
                book.SaveAs(stage,51);book.Close(false);lifetime.Release((object)book);book=null;
                timing("save");
                if(preserveFormatting) PreserveMaskAppearance(Directory.GetFiles(stageDirectory,"appearance-source.*").Single(),(object)writer,stage,stageDirectory,result,cancel,lifetime);
                timing("appearance");progress?.Invoke("重开文件并核对全表",80,source.Cells.Count,source.Cells.Count);
                CleanMaskMetadata(stage);
                try {book=lifetime.Hold(books.Open(stage,0,true));}
                catch(COMException ex){throw new InvalidOperationException("Excel 无法重新打开生成文件，未交付结果。",ex);}
                if(result.Formulas>0)writer.CalculateFullRebuild();ValidateMaskBook(book,result,cancel,lifetime);book.Close(false);lifetime.Release((object)book);book=null;
                timing("reopen");
                cancel.ThrowIfCancellationRequested();return stage;
            }
            finally
            {
                try {if(book!=null)book.Close(false);}catch { }
                try {if(writer!=null)writer.Quit();}catch { }
                timing("quit");
                lifetime.Dispose();writer=null;book=null;
                timing("release");LastMaskPerformance=string.Join("; ",stages);
                // Some Office configurations keep an owned invisible instance alive after Quit.
                // Bound teardown on this STA, and terminate only the private process we created.
                if(process>0 && process!=source.HostProcessId)
                {
                    try
                    {
                        using(var owned=System.Diagnostics.Process.GetProcessById(process))
                        {
                            var cleanup=System.Diagnostics.Stopwatch.StartNew();
                            while(!owned.HasExited && cleanup.ElapsedMilliseconds<2000){System.Windows.Forms.Application.DoEvents();Thread.Sleep(10);}
                            if(!owned.HasExited){owned.Kill();owned.WaitForExit(2000);stages.Add("owned-process cleanup");}
                        }
                    }
                    catch(ArgumentException){}catch(InvalidOperationException){}
                    LastMaskPerformance=string.Join("; ",stages);
                }

            }
        }
        private static void ValidateMaskBook(dynamic book,MaskResult result,CancellationToken cancel,MaskComLifetime lifetime)
        {
            dynamic sheets=lifetime.Hold(book.Worksheets),names=lifetime.Hold(book.Names);
            if((int)sheets.Count!=result.Output.Sheets.Count || (int)lifetime.Hold(book.Connections).Count!=0 || (bool)book.HasVBProject)throw new InvalidOperationException("结果包含未允许的工作簿对象。");
            for(int n=1;n<=(int)names.Count;n++)
            {
                dynamic name=lifetime.Hold(names.Item(n));string id=name.Name;
                if((id!="_xlfn.SUMIFS" && id!="_xlfn.COUNTIFS") || (string)name.RefersTo!="=#NAME?")throw new InvalidOperationException("结果包含未允许的定义名称。");
            }
            var neutral=result.Output.Sheets.Select((s,n)=>new DataSheet {Name="Sheet"+(n+1)}).ToArray();
            long seen=0;
            for(int i=1;i<=(int)sheets.Count;i++)
            {
                dynamic sheet=lifetime.Hold(sheets[i]);
                if((int)lifetime.Hold(sheet.Shapes).Count!=0 || (int)lifetime.Hold(sheet.Hyperlinks).Count!=0 || (int)lifetime.Hold(sheet.Comments).Count!=0)throw new InvalidOperationException("结果包含未允许的工作表对象。");
                ScanMaskSheet((object)sheet,i-1,cancel,actual=>
                {
                    var expected=FindMaskCell(result.Output.Cells,actual);
                    if(expected==null)throw new InvalidOperationException("输出出现非预期数据。");
                    object wanted=expected.Formula==null?expected.Value:result.Calculated[expected.Key];
                    if(!MaskPlanner.Equivalent(wanted,actual.Value))throw new InvalidOperationException(expected.Location+"：Excel 重算／保存后验证不一致，未导出。");
                    if((expected.Formula==null)!=(actual.Formula==null) || (actual.Formula!=null && SafeFormula.Rewrite(actual.Formula,neutral,new Dictionary<string,string>())!=expected.Formula))throw new InvalidOperationException(expected.Location+"：公式保存后改变。");
                    seen++;
                },null,lifetime);
            }
            if(seen!=result.Output.Cells.Count)throw new InvalidOperationException("输出单元格数量不一致，未交付不完整结果。");
        }
        private static void CleanMaskMetadata(string path)
        {
            // Excel repopulates lastModifiedBy on SaveAs, even after clearing built-in properties.
            // Scrub the saved container before the final read-only Excel reopen and validation.
            using(var zip=ZipFile.Open(path,ZipArchiveMode.Update))
            {
                // Manual calculation is private to the writer; the delivered workbook opens in auto mode.
                var workbook=ReadPart(zip,"xl/workbook.xml");var calculation=workbook.Root.Element(SpreadsheetNs+"calcPr");
                if(calculation!=null){calculation.SetAttributeValue("calcMode","auto");calculation.SetAttributeValue("fullCalcOnLoad","1");WritePart(zip,"xl/workbook.xml",workbook);}
                if(zip.GetEntry("docProps/custom.xml")!=null)throw new InvalidOperationException("结果出现未允许的自定义文档属性。");
                foreach(string part in new[]{"docProps/core.xml","docProps/app.xml"})
                {
                    var entry=zip.GetEntry(part);if(entry==null)continue;
                    XDocument document;
                    using(var stream=entry.Open())using(var reader=XmlReader.Create(stream,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))document=XDocument.Load(reader);
                    foreach(var element in document.Root.Elements())
                    {
                        string name=element.Name.LocalName;
                        if(part.EndsWith("core.xml",StringComparison.Ordinal)?name!="created" && name!="modified":name=="Company" || name=="Manager" || name=="HyperlinkBase")element.Value="";
                    }
                    entry.Delete();using(var stream=zip.CreateEntry(part,CompressionLevel.Optimal).Open())document.Save(stream);
                }
            }
        }
        [DllImport("user32.dll",EntryPoint="GetWindowThreadProcessId")] private static extern uint GetWindowThreadProcessIdForMask(IntPtr hwnd,out int process);
    }
}
