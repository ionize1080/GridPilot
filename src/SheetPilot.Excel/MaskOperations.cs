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
        internal static string LastMaskPerformance;
        public MaskSnapshot CaptureMaskWorkbook()
        {
            CheckThread(); dynamic app=application(), book=app.ActiveWorkbook;
            if(book==null) throw new InvalidOperationException("请先打开需要脱敏的工作簿。");
            if((int)app.CalculationState!=0) throw new InvalidOperationException("请等待 Excel 完成计算后再分析。");
            var snapshot=new MaskSnapshot { WorkbookKey=WorkbookKey(book),Date1904=(bool)book.Date1904 };
            dynamic sheets=book.Worksheets; long count=0;
            for(int i=1;i<=(int)sheets.Count;i++)
            {
                dynamic sheet=sheets[i];
                var spec=new DataSheet { Name=(string)sheet.Name };snapshot.Sheets.Add(spec);
                dynamic all=sheet.Cells;
                dynamic lastRow=all.Find("*",Type.Missing,-4123,2,1,2,false,Type.Missing,false);
                if(lastRow==null)continue;
                dynamic lastColumn=all.Find("*",Type.Missing,-4123,2,2,2,false,Type.Missing,false);
                dynamic first=all.Find("*",Type.Missing,-4123,2,1,1,false,Type.Missing,false);
                spec.Rows=(int)lastRow.Row;spec.Columns=(int)lastColumn.Column;
                int firstRow=(int)first.Row, limit=Environment.Is64BitProcess?2000000:500000;
                dynamic bounded=sheet.Range[sheet.Cells[firstRow,1],sheet.Cells[spec.Rows,spec.Columns]];
                var ranges=new List<object>();
                // SpecialCells dirties some Excel workbooks even for read-only inspection.
                // Dense data uses bounded blocks; widely separated cells use Find traversal.
                if((long)(spec.Rows-firstRow+1)*spec.Columns<=8000000)ranges.Add((object)bounded);
                else
                {
                    dynamic cursor=first;string initial=cursor.Address;
                    do
                    {
                        ranges.Add((object)cursor);
                        if(ranges.Count>limit)throw new InvalidOperationException("有内容单元格超过本进程上限 "+limit.ToString("N0")+"，请分批处理。");
                        cursor=all.Find("*",cursor,-4123,2,1,1,false,Type.Missing,false);
                    }while(cursor!=null && (string)cursor.Address!=initial);
                }
                var seen=new HashSet<string>();
                foreach(dynamic range in ranges)
                {
                    int r0=range.Row,c0=range.Column,rows=range.Rows.Count,cols=range.Columns.Count;
                    int step=Math.Max(1,16384/cols);
                    for(int offset=0;offset<rows;offset+=step)
                    {
                        int nr=Math.Min(step,rows-offset);dynamic block=sheet.Range[sheet.Cells[r0+offset,c0],sheet.Cells[r0+offset+nr-1,c0+cols-1]];
                        object raw=block.Value2,rawFormula=block.Formula;var values=raw as object[,];var formulas=rawFormula as object[,];
                        for(int r=1;r<=nr;r++)for(int c=1;c<=cols;c++)
                        {
                            object value=values==null?raw:values[r,c],f=formulas==null?rawFormula:formulas[r,c];
                            if(value==null && (f==null || Equals(f,"")))continue;
                            int row=r0+offset+r-1,col=c0+c-1;string key=row+":"+col;if(!seen.Add(key))continue;
                            if(value is ErrorWrapper || value is int)throw new InvalidOperationException("S"+i+"!"+DataCell.Address(row,col)+"：源数据含 Excel 错误值，请先处理。");
                            string formula=null;
                            if(f is string text && text.StartsWith("=",StringComparison.Ordinal))
                            { dynamic cell=sheet.Cells[row,col];if((bool)cell.HasFormula){if((bool)cell.HasArray)throw new InvalidOperationException("S"+i+"：数组公式不在本版支持范围。");formula=text;} }
                            if(++count>limit)throw new InvalidOperationException("有内容单元格超过本进程上限 "+limit.ToString("N0")+"（格式空白不计数），请分批处理。");
                            snapshot.Cells.Add(new DataCell {Sheet=i-1,Row=row,Column=col,Value=value,Formula=formula});
                        }
                    }
                }
                var merged=new HashSet<string>();
                VisitMerges((object)sheet,new CellRect(firstRow,1,Math.Min(40,spec.Rows-firstRow+1),spec.Columns),m=>
                {if(merged.Add(m.Address))spec.HeaderMerges.Add(new HeaderMerge {Row=m.Row,Column=m.Column,Rows=m.Rows,Columns=m.Columns});},(stage,done)=>{});

            }
            if(snapshot.Sheets.Count==0)throw new InvalidOperationException("工作簿没有普通工作表。");
            snapshot.Cells=snapshot.Cells.OrderBy(c=>c.Sheet).ThenBy(c=>c.Row).ThenBy(c=>c.Column).ToList();
            return snapshot;
        }
        public static IReadOnlyList<MaskRule> SuggestMaskRules(MaskSnapshot snapshot)
        {
            return MaskInference.Suggest(snapshot);
        }
        public void VerifyMaskSource(MaskSnapshot source)
        {
            var current=CaptureMaskWorkbook();
            if(source.WorkbookKey!=current.WorkbookKey || source.Date1904!=current.Date1904 || source.Sheets.Count!=current.Sheets.Count || source.Cells.Count!=current.Cells.Count ||
                !source.Sheets.Select(s=>s.Name).SequenceEqual(current.Sheets.Select(s=>s.Name)) ||
                source.Cells.Where((c,i)=>c.Key!=current.Cells[i].Key || !Equals(c.Value,current.Cells[i].Value) || c.Formula!=current.Cells[i].Formula).Any())
                throw new InvalidOperationException("源工作簿在预览后改变，请关闭窗口并重新分析。");
        }
        public string ExportMask(MaskSnapshot source,MaskResult result,string target,CancellationToken cancel,bool preserveFormatting=false)
        {
            CheckThread(); VerifyMaskSource(source); cancel.ThrowIfCancellationRequested();
            target=Path.GetFullPath(target);
            if(!string.Equals(Path.GetExtension(target),".xlsx",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("结果必须保存为新的 .xlsx 文件。");
            if(File.Exists(target))throw new InvalidOperationException("目标文件已存在，请选择新名称；不覆盖已有文件。");
            string stageDirectory=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"GridPilot","Staging",Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stageDirectory);string stage=Path.Combine(stageDirectory,"result.xlsx");
            string destinationStage=Path.Combine(Path.GetDirectoryName(target),".GridPilot-"+Guid.NewGuid().ToString("N")+".tmp");
            var lifetime=new MaskComLifetime();dynamic writer=null,book=null; int process=0;
            var clock=System.Diagnostics.Stopwatch.StartNew();var stages=new List<string>();Action<string> timing=stage=>stages.Add(stage+"="+clock.ElapsedMilliseconds+"ms");
            try
            {
                // A separate owned Excel instance avoids recalculating or changing the source workbook.
                writer=lifetime.Hold(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")));
                writer.Visible=false;writer.DisplayAlerts=false;writer.EnableEvents=false;writer.AutomationSecurity=3;
                GetWindowThreadProcessIdForMask(new IntPtr((int)writer.Hwnd),out process);
                LastMaskWriterProcess=process;
                timing("writer");
                dynamic books=lifetime.Hold(writer.Workbooks);book=lifetime.Hold(books.Add(-4167));book.Date1904=source.Date1904;
                writer.Calculation=-4135;writer.CalculateBeforeSave=false;
                timing("blank");
                dynamic sheets=lifetime.Hold(book.Worksheets);
                for(int s=0;s<source.Sheets.Count;s++)
                { dynamic sheet=lifetime.Hold(s==0?sheets[1]:sheets.Add(Type.Missing,lifetime.Hold(sheets[sheets.Count])));sheet.Name="Sheet"+(s+1); }
                for(int s=0;s<source.Sheets.Count;s++)
                {
                    cancel.ThrowIfCancellationRequested();dynamic sheet=lifetime.Hold(sheets[s+1]),all=lifetime.Hold(sheet.Cells);var spec=source.Sheets[s];
                    foreach(var group in result.Output.Cells.Where(c=>c.Sheet==s).GroupBy(c=>(c.Row-1)/128))
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
                    }
                    lifetime.Hold(sheet.Columns).ColumnWidth=20;
                }
                timing("data");
                if(result.Formulas>0)writer.CalculateFullRebuild();
                timing("calculate");
                ValidateMaskBook(book,result,cancel,lifetime);
                timing("write+validate");
                dynamic properties=lifetime.Hold(book.BuiltinDocumentProperties);
                foreach(string property in new[]{"Author","Last Author","Title","Subject","Comments","Keywords","Category","Company","Manager"}) try { lifetime.Hold(properties[property]).Value=""; }catch(COMException) { }
                book.SaveAs(stage,51);book.Close(false);book=null;
                timing("save");
                if(preserveFormatting) PreserveMaskAppearance((object)((dynamic)application()).ActiveWorkbook,(object)writer,stage,stageDirectory,result,cancel,lifetime);
                timing("appearance");
                CleanMaskMetadata(stage);
                try {book=lifetime.Hold(books.Open(stage,0,true));}
                catch(COMException ex){throw new InvalidOperationException("Excel 无法重新打开生成文件，未交付结果。",ex);}
                if(result.Formulas>0)writer.CalculateFullRebuild();ValidateMaskBook(book,result,cancel,lifetime);book.Close(false);book=null;
                timing("reopen");
                cancel.ThrowIfCancellationRequested(); VerifyMaskSource(source);
                // Commit only a complete, verified file; a failed copy never exposes a partial .xlsx.
                File.Copy(stage,destinationStage,false);
                cancel.ThrowIfCancellationRequested();
                File.Move(destinationStage,target);
                timing("commit");
                return target;
            }
            finally
            {
                try {if(book!=null)book.Close(false);}catch { }
                try {if(writer!=null)writer.Quit();}catch { }
                timing("quit");
                lifetime.Dispose();writer=null;book=null;
                timing("release");LastMaskPerformance=string.Join("; ",stages);
                try {if(File.Exists(destinationStage))File.Delete(destinationStage);}catch(IOException) { }
                try { if(File.Exists(stage))File.Delete(stage);Directory.Delete(stageDirectory); }catch(IOException) { }
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
            var index=result.Output.Cells.ToDictionary(c=>c.Key);
            for(int i=1;i<=(int)sheets.Count;i++)
            {
                dynamic sheet=lifetime.Hold(sheets[i]),all=lifetime.Hold(sheet.Cells);
                if((int)lifetime.Hold(sheet.Shapes).Count!=0 || (int)lifetime.Hold(sheet.Hyperlinks).Count!=0 || (int)lifetime.Hold(sheet.Comments).Count!=0)throw new InvalidOperationException("结果包含未允许的工作表对象。");
                // Find is read-only and avoids expected SpecialCells errors on sheets without formulas.
                var actualCells=new HashSet<string>();
                var ranges=new List<object>();dynamic first=lifetime.Hold(all.Find("*",Type.Missing,-4123,2,1,1,false,Type.Missing,false));
                if(first!=null)
                {
                    dynamic lastRow=lifetime.Hold(all.Find("*",Type.Missing,-4123,2,1,2,false,Type.Missing,false)),lastColumn=lifetime.Hold(all.Find("*",Type.Missing,-4123,2,2,2,false,Type.Missing,false));
                    int top=first.Row,bottom=lastRow.Row,right=lastColumn.Column;
                    if((long)(bottom-top+1)*right<=8000000)ranges.Add((object)lifetime.Hold(sheet.Range[lifetime.Hold(all[top,1]),lifetime.Hold(all[bottom,right])]));
                    else
                    {
                        dynamic cursor=first;string initial=cursor.Address;
                        do {cancel.ThrowIfCancellationRequested();ranges.Add((object)cursor);if(ranges.Count>result.Output.Cells.Count)throw new InvalidOperationException("输出出现非预期数据。");cursor=lifetime.Hold(all.Find("*",cursor,-4123,2,1,1,false,Type.Missing,false));}while(cursor!=null && (string)cursor.Address!=initial);
                    }
                }
                foreach(dynamic region in ranges)
                {
                        int rows=lifetime.Hold(region.Rows).Count,cols=lifetime.Hold(region.Columns).Count,r0=region.Row,c0=region.Column;
                        int step=Math.Max(1,16384/cols);
                        for(int offset=0;offset<rows;offset+=step)
                        {
                            int nr=Math.Min(step,rows-offset);dynamic block=lifetime.Hold(sheet.Range[lifetime.Hold(all[r0+offset,c0]),lifetime.Hold(all[r0+offset+nr-1,c0+cols-1])]);
                            object raw=block.Value2;var values=raw as object[,];
                            for(int r=1;r<=nr;r++)for(int c=1;c<=cols;c++)
                            {
                                cancel.ThrowIfCancellationRequested();string key=(i-1)+":"+(r+r0+offset-1)+":"+(c+c0-1);object actual=values==null?raw:values[r,c];
                                if(actual==null)continue;actualCells.Add(key);
                                if(!index.TryGetValue(key,out var expected))throw new InvalidOperationException("输出出现非预期数据。");
                                object wanted=expected.Formula==null?expected.Value:result.Calculated[key];
                                if(!MaskPlanner.Equivalent(wanted,actual))throw new InvalidOperationException(expected.Location+"：Excel 重算／保存后验证不一致，未导出。");
                                if(expected.Formula!=null)
                                {
                                    var neutral=result.Output.Sheets.Select((s,n)=>new DataSheet {Name="Sheet"+(n+1)}).ToArray();
                                    if(SafeFormula.Rewrite((string)lifetime.Hold(all[r+r0+offset-1,c+c0-1]).Formula,neutral,new Dictionary<string,string>())!=expected.Formula)throw new InvalidOperationException(expected.Location+"：公式保存后改变。");
                                }
                            }
                        }
                }
                foreach(var expected in result.Output.Cells.Where(c=>c.Sheet==i-1 && (c.Value!=null || c.Formula!=null)))
                    if(!actualCells.Contains(expected.Key) && !MaskPlanner.Equivalent(expected.Formula==null?expected.Value:result.Calculated[expected.Key],null))throw new InvalidOperationException(expected.Location+"：输出缺少数据。");

            }
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
