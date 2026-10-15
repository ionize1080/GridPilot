using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using SheetPilot.Core;
using SheetPilot.Excel;

internal static partial class Program
{
    private static SimpleMaskSnapshot HeaderFixture(object[,] values,int top=1)
    {
        var s=new SimpleMaskSnapshot();s.Blocks.Add(new SimpleMaskBlock{Row=top,Column=1,Rows=values.GetLength(0),Columns=values.GetLength(1)});
        for(int r=0;r<values.GetLength(0);r++)for(int c=0;c<values.GetLength(1);c++)if(values[r,c]!=null)
        {s.HeaderContext.Add(new HeaderProbeCell{Row=r+top,Column=c+1,Value=values[r,c]});s.Cells.Add(new SimpleMaskCell{Row=r+top,Column=c+1,Address=DataCell.Address(r+top,c+1),Before=values[r,c]});}return s;
    }
    private static void SimpleHeaderTests()
    {
        var s=HeaderFixture(new object[,]{{"余额报表",null,null},{"导出时间：2026-10-15",null,null},{"导出者：王某",null,null},{"客户名","贷款号","余额"},{"星海有限公司","P123",123.0},{"长河有限公司","P456",234.0}});
        s.HeaderMerges.Add(new HeaderMerge{Row=1,Column=1,Rows=1,Columns=3});
        var r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);
        Check(r.Headers.Bands.Count==1&&r.Headers.Bands[0].StartRow==4&&r.Retained==3,"metadata and merged title excluded; default fields retained");
        Check(r.Cells[0].Role=="preamble"&&r.Cells[2].After!="导出者：王某","preamble classified and sensitive metadata masked by default");
        Check(SimpleMask.Output(s,r).Output.Cells.Any(c=>Equals(c.Value,"客户名")),"retained headers included in output");
        r=SimpleMask.Build(s,new SimpleMaskOptions{MaskHeaders=true},CancellationToken.None);Check(r.Retained==0&&r.Cells[3].After!="客户名","optional header masking");
        r=SimpleMask.Build(s,new SimpleMaskOptions{PreservePreamble=true},CancellationToken.None);Check(r.Retained==6,"explicit preamble retention");
        r=SimpleMask.Build(s,new SimpleMaskOptions{HeaderMode="off"},CancellationToken.None);Check(r.Retained==0,"recognition off");
        r=SimpleMask.Build(s,new SimpleMaskOptions{HeaderMode="manual",HeaderRows="1,4-4"},CancellationToken.None);Check(r.Retained==4,"manual absolute row override");
        foreach(string bad in new[]{"","0","4-2","1-101","1048577","a"})Reject(()=>SimpleHeaderDetector.ParseRows(bad),"invalid manual rows "+bad);
        s=HeaderFixture(new object[,]{{"星海有限公司"},{"长河有限公司"},{"客户名称有限公司"}},2);
        Check(SimpleHeaderDetector.Analyze(s,new SimpleMaskOptions(),CancellationToken.None).Bands.Count==0,"company data without header not retained");
        s=HeaderFixture(new object[,]{{"客户名","金额",null},{null,"本期","上期"},{"星海有限公司",12.0,13.0},{"长河有限公司",20.0,21.0}},5);
        s.HeaderMerges.Add(new HeaderMerge{Row=5,Column=1,Rows=2,Columns=1});s.HeaderMerges.Add(new HeaderMerge{Row=5,Column=2,Rows=1,Columns=2});
        var h=SimpleHeaderDetector.Analyze(s,new SimpleMaskOptions(),CancellationToken.None);
        Check(h.Bands.Count==1&&h.Bands[0].StartRow==5&&h.Bands[0].EndRow==6&&h.Bands[0].Fields[1].Label=="金额 / 本期","multilevel merged header and offset rows");
        s.Cells=s.Cells.Where(c=>c.Column==2).ToList();r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);Check(r.Retained==2,"single column uses full context");
        s.Cells=s.Cells.Where(c=>c.Row>=7).ToList();r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);Check(r.Retained==0&&r.Cells.All(c=>c.Role=="data"),"data only selection does not promote first data row");
        s=HeaderFixture(new object[,]{{"Foo","Bar"},{"a","b"}});s.NativeHeaders.Add(new SimpleHeaderBand{StartRow=1,EndRow=1,FirstColumn=1,LastColumn=2});
        r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);Check(r.Retained==2,"native Excel table overrides unknown labels");
        s.Cells[0].Skip="公式";r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);Check(r.Cells[0].Skip=="公式"&&!r.Cells[0].Retained,"header does not bypass formula skip");
        s.Cells[0].Skip="合并单元格";r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);Check(r.Cells[0].Retained,"merged header can be retained");
        r=SimpleMask.Build(s,new SimpleMaskOptions{MaskHeaders=true},CancellationToken.None);Check(r.Cells[0].Skip!=null,"partial merged header never written");
        s.Cells[0].MergeSelectedFull=true;r=SimpleMask.Build(s,new SimpleMaskOptions{MaskHeaders=true},CancellationToken.None);Check(r.Cells[0].Skip==null&&r.Cells[0].After!="Foo","fully selected merged header can be masked");
        s=HeaderFixture(new object[,]{{"Name","Amount"},{"Alice",12.0},{"Bob",15.0},{"Name","Amount"},{"Claire",17.0},{"Dave",19.0}});
        Check(SimpleHeaderDetector.Analyze(s,new SimpleMaskOptions(),CancellationToken.None).Bands.Count==2,"repeated print header identified");
        s=HeaderFixture(new object[,]{{"余额",null},{"客户名","金额"},{"星海有限公司",10.0}});s.HeaderMerges.Add(new HeaderMerge{Row=1,Column=1,Rows=1,Columns=2});
        Check(SimpleHeaderDetector.Analyze(s,new SimpleMaskOptions(),CancellationToken.None).Bands.Single().StartRow==2,"field word in wide merged title not a header tier");
    }
    private static void SimpleHeaderExcelTests(string sample=null)
    {
        dynamic app=null;string output=Path.Combine(Path.GetTempPath(),"header-"+Guid.NewGuid().ToString("N")+".xlsx");
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            dynamic book=app.Workbooks.Add(-4167),sheet=book.Worksheets[1];var gateway=new ExcelGateway(()=>(object)app);
            sheet.Range["A1:C1"].Merge();sheet.Range["A1"].Value2="贷款余额报表";sheet.Range["A2:C2"].Value2=new object[,]{{"客户名","贷款号","余额"}};sheet.Range["A3:C4"].Value2=new object[,]{{"星海有限公司","P123",12.0},{"长河有限公司","P456",15.0}};
            sheet.Range["A:C"].Select();var s=gateway.CaptureSimpleMask(CancellationToken.None);var r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);
            Check(r.Retained==3&&r.Headers.Bands[0].StartRow==2&&r.Cells[0].Skip==null,"actual Excel merged title and fields");
            gateway.ExportSimpleMask(s,r,output,CancellationToken.None);dynamic exported=app.Workbooks.Open(output,0,true);Check((string)exported.Worksheets[1].Range["A2"].Value2=="客户名"&&(string)exported.Worksheets[1].Range["A3"].Value2!="星海有限公司","XLSX keeps headers and masks records");exported.Close(false);
            gateway.ApplySimpleMask(s,r,true,CancellationToken.None);Check((string)sheet.Range["A2"].Value2=="客户名"&&(bool)sheet.Range["A1"].MergeCells&&(string)sheet.Range["A1"].Value2!="贷款余额报表","in place retains fields and masks merged title");gateway.UndoLast();Check((string)sheet.Range["A1"].Value2=="贷款余额报表","undo merged title");
            sheet.Range["B:B"].Select();s=gateway.CaptureSimpleMask(CancellationToken.None);r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);Console.WriteLine("Column selection: "+(string)app.Selection.Address+" retained="+r.Retained+" bands="+r.Headers.Message);Check(r.Retained==s.Cells.Count(c=>c.Row==2),"actual whole column header recognition");
            sheet.Range["A3:C4"].Select();s=gateway.CaptureSimpleMask(CancellationToken.None);r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);Check(r.Retained==0,"actual data only selection");
            sheet.Range["A2:B2"].Merge();sheet.Range["A2"].Value2="客户名";sheet.Range["A2:C4"].Select();s=gateway.CaptureSimpleMask(CancellationToken.None);r=SimpleMask.Build(s,new SimpleMaskOptions{HeaderMode="manual",HeaderRows="2",MaskHeaders=true},CancellationToken.None);
            gateway.ApplySimpleMask(s,r,true,CancellationToken.None);Check((string)sheet.Range["A2"].Value2=="KHM"&&(bool)sheet.Range["A2"].MergeCells,"merged header opt in masks anchor without unmerging");gateway.UndoLast();Check((string)sheet.Range["A2"].Value2=="客户名","merged header undo");
            sheet.Range["A2:B2"].UnMerge();gateway.ApplySimpleMask(s,r,true,CancellationToken.None);Check((string)sheet.Range["A2"].Value2=="客户名","changed merge shape skipped");if(gateway.CanUndo)gateway.UndoLast();
            sheet.Range["F90:G92"].Value2=new object[,]{{"CustomX","CustomY"},{"alpha",1.0},{"beta",2.0}};sheet.ListObjects.Add(1,sheet.Range["F90:G92"],Type.Missing,1);sheet.Range["F90:G92"].Select();s=gateway.CaptureSimpleMask(CancellationToken.None);r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);Check(r.Retained==2&&r.Headers.Bands.Any(b=>b.Source=="excel-table"&&b.StartRow==90),"formal Excel table outside scan window");
            gateway.ReleaseUndoEvents();book.Close(false);
            if(sample!=null)
            {
                var before=File.ReadAllBytes(sample);book=app.Workbooks.Open(sample,0,true);sheet=book.Worksheets["Page1_1"];sheet.Activate();sheet.Range["A:T"].Select();
                s=gateway.CaptureSimpleMask(CancellationToken.None);r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);
                Check(r.Headers.Bands.Count==1&&r.Headers.Bands[0].StartRow==2&&r.Retained==20,"provided balance workbook twenty fields row two");
                sheet=book.Worksheets["Sheet2"];sheet.Activate();sheet.Range["B:B"].Select();s=gateway.CaptureSimpleMask(CancellationToken.None);r=SimpleMask.Build(s,new SimpleMaskOptions(),CancellationToken.None);Check(r.Headers.Bands.Count==0&&r.Retained==0,"provided company sheet no header");book.Close(false);Check(before.SequenceEqual(File.ReadAllBytes(sample)),"reference workbook byte identical");
            }
        }
        finally{if(app!=null){try{app.Quit();}catch{}Marshal.ReleaseComObject(app);}if(File.Exists(output))File.Delete(output);}
    }
}
