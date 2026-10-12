using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class ExcelGateway
    {
        internal static int LastMaskWriterProcess;
        public MaskSnapshot CaptureMaskWorkbook()
        {
            CheckThread(); dynamic app=application(), book=app.ActiveWorkbook;
            if(book==null) throw new InvalidOperationException("请先打开需要脱敏的工作簿。");
            if((int)app.CalculationState!=0) throw new InvalidOperationException("请等待 Excel 完成计算后再分析。");
            var snapshot=new MaskSnapshot { WorkbookKey=WorkbookKey(book),Date1904=(bool)book.Date1904 };
            dynamic sheets=book.Worksheets; long count=0;
            for(int i=1;i<=(int)sheets.Count;i++)
            {
                dynamic sheet=sheets[i], used=sheet.UsedRange;
                int rows=used.Rows.Count,columns=used.Columns.Count,firstRow=used.Row,firstColumn=used.Column;
                count+=(long)rows*columns;if(count>200000 || columns>512) throw new InvalidOperationException("本版脱敏支持总计20万已使用单元格、每表512列；请缩小工作簿。");
                snapshot.Sheets.Add(new DataSheet { Name=(string)sheet.Name,Rows=firstRow+rows-1,Columns=firstColumn+columns-1 });
                object raw=used.Value2,rawFormula=used.Formula;var values=raw as object[,];var formulas=rawFormula as object[,];
                for(int r=1;r<=rows;r++) for(int c=1;c<=columns;c++)
                {
                    object value=values==null?raw:values[r,c], f=formulas==null?rawFormula:formulas[r,c];
                    if(value==null && f==null)continue;
                    if(value is ErrorWrapper || value is int) throw new InvalidOperationException("S"+i+"!"+DataCell.Address(r+firstRow-1,c+firstColumn-1)+"：源数据含 Excel 错误值，请先处理。");
                    string formula=null;
                    if(f is string text && text.StartsWith("=",StringComparison.Ordinal))
                    {
                        dynamic cell=used.Cells[r,c];
                        if((bool)cell.HasFormula)
                        { if((bool)cell.HasArray) throw new InvalidOperationException("S"+i+"：数组公式不在本版支持范围。");formula=text; }
                    }
                    snapshot.Cells.Add(new DataCell { Sheet=i-1,Row=r+firstRow-1,Column=c+firstColumn-1,Value=value,Formula=formula });
                }
            }
            if(snapshot.Sheets.Count==0)throw new InvalidOperationException("工作簿没有普通工作表。");
            return snapshot;
        }
        public static IReadOnlyList<MaskRule> SuggestMaskRules(MaskSnapshot snapshot)
        {
            return snapshot.Cells.GroupBy(c=>c.Sheet+":"+c.Column).Select(group=>
            {
                var first=group.OrderBy(c=>c.Row).First();
                return new MaskRule { Sheet=first.Sheet,SheetName=snapshot.Sheets[first.Sheet].Name,Column=first.Column,StartRow=first.Row+1,Header=Convert.ToString(first.Value),Domain="" };
            }).OrderBy(r=>r.Sheet).ThenBy(r=>r.Column).ToArray();
        }
        public void VerifyMaskSource(MaskSnapshot source)
        {
            var current=CaptureMaskWorkbook();
            if(source.WorkbookKey!=current.WorkbookKey || source.Date1904!=current.Date1904 || source.Sheets.Count!=current.Sheets.Count || source.Cells.Count!=current.Cells.Count ||
                !source.Sheets.Select(s=>s.Name).SequenceEqual(current.Sheets.Select(s=>s.Name)) ||
                source.Cells.Where((c,i)=>c.Key!=current.Cells[i].Key || !Equals(c.Value,current.Cells[i].Value) || c.Formula!=current.Cells[i].Formula).Any())
                throw new InvalidOperationException("源工作簿在预览后改变，请关闭窗口并重新分析。");
        }
        public string ExportMask(MaskSnapshot source,MaskResult result,string target,CancellationToken cancel)
        {
            CheckThread(); VerifyMaskSource(source); cancel.ThrowIfCancellationRequested();
            target=Path.GetFullPath(target);
            if(!string.Equals(Path.GetExtension(target),".xlsx",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("结果必须保存为新的 .xlsx 文件。");
            if(File.Exists(target))throw new InvalidOperationException("目标文件已存在，请选择新名称；不覆盖已有文件。");
            string stageDirectory=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"GridPilot","Staging",Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stageDirectory);string stage=Path.Combine(stageDirectory,"result.xlsx");
            string destinationStage=Path.Combine(Path.GetDirectoryName(target),".GridPilot-"+Guid.NewGuid().ToString("N")+".tmp");
            object appObject=null;dynamic writer=null,book=null; int process=0;
            try
            {
                // A separate owned Excel instance avoids recalculating or changing the source workbook.
                appObject=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));writer=appObject;
                writer.Visible=false;writer.DisplayAlerts=false;writer.EnableEvents=false;writer.AutomationSecurity=3;
                GetWindowThreadProcessIdForMask(new IntPtr((int)writer.Hwnd),out process);
                LastMaskWriterProcess=process;
                book=writer.Workbooks.Add(-4167);book.Date1904=source.Date1904;
                for(int s=0;s<source.Sheets.Count;s++)
                { dynamic sheet=s==0?book.Worksheets[1]:book.Worksheets.Add(Type.Missing,book.Worksheets[book.Worksheets.Count]);sheet.Name="Sheet"+(s+1); }
                for(int s=0;s<source.Sheets.Count;s++)
                {
                    cancel.ThrowIfCancellationRequested();dynamic sheet=book.Worksheets[s+1];var spec=source.Sheets[s];
                    if((long)spec.Rows*spec.Columns>2000000)throw new InvalidOperationException("工作表数据位置过于稀疏，清理已使用范围后重试。");
                    var data=new object[Math.Max(1,spec.Rows),Math.Max(1,spec.Columns)];
                    foreach(var cell in result.Output.Cells.Where(c=>c.Sheet==s && c.Formula==null)) data[cell.Row-1,cell.Column-1]=cell.Value is string text?"'"+text:cell.Value;
                    dynamic area=sheet.Range[sheet.Cells[1,1],sheet.Cells[data.GetLength(0),data.GetLength(1)]];
                    area.Value2=data;
                    foreach(var cell in result.Output.Cells.Where(c=>c.Sheet==s && c.Formula!=null)) {cancel.ThrowIfCancellationRequested();sheet.Cells[cell.Row,cell.Column].Formula=cell.Formula;}
                    foreach(var cell in result.Output.Cells.Where(c=>c.Sheet==s && c.NumberFormat!=null))sheet.Cells[cell.Row,cell.Column].NumberFormat=cell.NumberFormat;
                    sheet.Columns.ColumnWidth=20;
                }
                writer.CalculateFullRebuild();
                ValidateMaskBook(book,result,cancel);
                foreach(string property in new[]{"Author","Last Author","Title","Subject","Comments","Keywords","Category","Company","Manager"}) try { book.BuiltinDocumentProperties[property].Value=""; }catch(COMException) { }
                book.SaveAs(stage,51);book.Close(false);book=null;
                CleanMaskMetadata(stage);
                book=writer.Workbooks.Open(stage,0,true);writer.CalculateFullRebuild();ValidateMaskBook(book,result,cancel);book.Close(false);book=null;
                cancel.ThrowIfCancellationRequested(); VerifyMaskSource(source);
                // Commit only a complete, verified file; a failed copy never exposes a partial .xlsx.
                File.Copy(stage,destinationStage,false);
                cancel.ThrowIfCancellationRequested();
                File.Move(destinationStage,target);
                return target;
            }
            finally
            {
                try {if(book!=null)book.Close(false);}catch { }
                try {if(writer!=null)writer.Quit();}catch { }
                if(appObject!=null && Marshal.IsComObject(appObject))Marshal.ReleaseComObject(appObject);
                appObject=null;writer=null;book=null;
                try {if(File.Exists(destinationStage))File.Delete(destinationStage);}catch(IOException) { }
                try { if(File.Exists(stage))File.Delete(stage);Directory.Delete(stageDirectory); }catch(IOException) { }
            }
        }
        private static void ValidateMaskBook(dynamic book,MaskResult result,CancellationToken cancel)
        {
            if((int)book.Worksheets.Count!=result.Output.Sheets.Count || (int)book.Connections.Count!=0 || (bool)book.HasVBProject)throw new InvalidOperationException("结果包含未允许的工作簿对象。");
            for(int n=1;n<=(int)book.Names.Count;n++)
            {
                dynamic name=book.Names.Item(n);string id=name.Name;
                if((id!="_xlfn.SUMIFS" && id!="_xlfn.COUNTIFS") || (string)name.RefersTo!="=#NAME?")throw new InvalidOperationException("结果包含未允许的定义名称。");
            }
            var index=result.Output.Cells.ToDictionary(c=>c.Key);
            for(int i=1;i<=(int)book.Worksheets.Count;i++)
            {
                dynamic sheet=book.Worksheets[i],used=sheet.UsedRange;
                if((int)sheet.Shapes.Count!=0 || (int)sheet.Hyperlinks.Count!=0 || (int)sheet.Comments.Count!=0)throw new InvalidOperationException("结果包含未允许的工作表对象。");
                object raw=used.Value2;var values=raw as object[,];int r0=used.Row,c0=used.Column,rowCount=used.Rows.Count,columnCount=used.Columns.Count;
                for(int r=1;r<=rowCount;r++)for(int c=1;c<=columnCount;c++)
                {
                    cancel.ThrowIfCancellationRequested();string key=(i-1)+":"+(r+r0-1)+":"+(c+c0-1);object actual=values==null?raw:values[r,c];
                    if(!index.TryGetValue(key,out var expected)) {if(actual!=null)throw new InvalidOperationException("输出出现非预期数据。");continue;}
                    object wanted=expected.Formula==null?expected.Value:result.Calculated[key];
                    if(!MaskPlanner.Equivalent(wanted,actual))throw new InvalidOperationException(expected.Location+"：Excel 重算／保存后验证不一致，未导出。");
                    if(expected.Formula!=null)
                    {
                        var neutral=result.Output.Sheets.Select((s,n)=>new DataSheet {Name="Sheet"+(n+1)}).ToArray();
                        if(SafeFormula.Rewrite((string)sheet.Cells[r+r0-1,c+c0-1].Formula,neutral,new Dictionary<string,string>())!=expected.Formula)
                            throw new InvalidOperationException(expected.Location+"：公式保存后改变。");
                    }
                }
            }
        }
        private static void CleanMaskMetadata(string path)
        {
            // Excel repopulates lastModifiedBy on SaveAs, even after clearing built-in properties.
            // Scrub the saved container before the final read-only Excel reopen and validation.
            using(var zip=ZipFile.Open(path,ZipArchiveMode.Update))
            {
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
