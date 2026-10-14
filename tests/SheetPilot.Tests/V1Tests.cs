using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using SheetPilot.Core;
using SheetPilot.Excel;

internal static partial class Program
{
    private static MaskSnapshot V1Snapshot(params DataCell[] cells)
    {var s=new MaskSnapshot();s.Sheets.Add(new DataSheet{Name="Data"});s.Cells.AddRange(cells);return s;}
    private static MaskRule V1Rule(MaskKind kind=MaskKind.ChineseName,int col=1)=>new MaskRule{Sheet=0,Column=col,StartRow=1,Kind=kind,Domain="identity"};
    private static void V1CoreTests()
    {
        var options=new ToolPreferences();
        foreach(string invalid in new[]{"1.2,3","1,23","1,,234","1,234.5,6"})Check(!NumericConversion.TryNumber(invalid,options,CultureInfo.InvariantCulture,out _,out _),"F05 malformed number "+invalid);
        Check(NumericConversion.TryNumber("1,234.56",options,CultureInfo.InvariantCulture,out double n,out _) && n==1234.56,"F05 valid grouping remains supported");
        var literal=V1Snapshot(new DataCell{Row=1,Column=1,Value="Alice",Formula="=\"Alice\""});
        var result=MaskPlanner.Build(literal,new[]{V1Rule()},CancellationToken.None,true);
        Check((string)result.Calculated["0:1:1"]!="Alice" && !result.Output.Cells[0].Formula.Contains("Alice"),"F01 formula-only identity literal is masked");
        var reference=V1Snapshot(new DataCell{Row=1,Column=1,Value="Alice",Formula="=B1"},new DataCell{Row=1,Column=2,Value="Alice"});
        Reject(()=>MaskPlanner.Build(reference,new[]{V1Rule()},CancellationToken.None,false),"F01 kept source identity reference cannot leak");
        result=MaskPlanner.Build(reference,new[]{V1Rule()},CancellationToken.None,true);
        Check(Equals(result.Calculated["0:1:1"],result.Output.Cells[1].Value),"F01 reference linked by explicit mention replacement");
        foreach(string op in new[]{"<>","<=",">=","=","<",">"})
            Check(SafeFormula.Rewrite("=COUNTIF(A1:A3,\""+op+"Alice\")",literal.Sheets,new Dictionary<string,string>{{"Alice","Cedar"}})=="=COUNTIF(A1:A3,\""+op+"Cedar\")","S01 criterion operator "+op);
        var criteria=V1Snapshot(new DataCell{Row=1,Column=1,Value="Alice"},new DataCell{Row=2,Column=1,Value="Bob"},new DataCell{Row=1,Column=2,Value=1.0,Formula="=COUNTIF(A1:A2,\"<>Alice\")"});
        result=MaskPlanner.Build(criteria,new[]{V1Rule()},CancellationToken.None,true);
        Check(Equals(result.Calculated["0:1:2"],1.0),"S01 COUNTIF inequality preserves selected records");
        var fullColumn=V1Snapshot(new DataCell{Row=1,Column=1,Value="Alice"},new DataCell{Row=2,Column=1,Value="Bob"},new DataCell{Row=1,Column=2,Value=1048575.0,Formula="=COUNTIF(A:A,\"<>Alice\")"});
        result=MaskPlanner.Build(fullColumn,new[]{V1Rule()},CancellationToken.None,true);
        Check(Equals(result.Calculated["0:1:2"],1048575.0),"S01 full-column COUNTIF includes blank rows correctly");
        var shorts=V1Snapshot(new DataCell{Row=1,Column=1,Value="云帆"},new DataCell{Row=2,Column=1,Value="12"},new DataCell{Row=3,Column=1,Value="FF"},new DataCell{Row=1,Column=2,Value="备注：云帆 / 12 / FF"});
        for(int i=0;i<20;i++)
        {result=MaskPlanner.Build(shorts,new[]{V1Rule()},CancellationToken.None,true);Check(result.Output.Cells.Take(3).All(c=>!new[]{"云帆","12","FF"}.Any(v=>((string)c.Value).Contains(v))),"S02 no short-identity collisions or double replacement "+i);}
        foreach(double serial in new[]{59.0,2958465.0})
        {
            var date=V1Snapshot(new DataCell{Row=1,Column=1,Value=serial},new DataCell{Row=2,Column=1,Value=serial,Formula="=A1"});
            Reject(()=>MaskPlanner.Build(date,new[]{new MaskRule{Column=1,StartRow=2,Kind=MaskKind.Date,Domain="date",DayShift=1},new MaskRule{Column=2,StartRow=1,Kind=MaskKind.ChineseName,Domain="unused"}},CancellationToken.None),"F06 invalid date shift blocked "+serial);
        }
        // A valid changed identity ensures formula-only date checks actually run.
        var formulaDate=V1Snapshot(new DataCell{Row=1,Column=1,Value="Alice"},new DataCell{Row=1,Column=2,Value=59.0,Formula="=59"});
        Reject(()=>MaskPlanner.Build(formulaDate,new[]{V1Rule(),new MaskRule{Column=2,StartRow=1,Kind=MaskKind.Date,Domain="date",DayShift=1}},CancellationToken.None),"F06 formula date fictitious leap day");
        var affix=V1Rule();AffixAnalyzer.Apply(affix,new[]{"甲有限责任公司","乙有限责任公司","丙有限责任公司","丁有限责任公司","戊项目"});Check(affix.Suffix=="有限责任公司","T01 modal company suffix");
        AffixAnalyzer.Apply(affix,new[]{"甲改造项目","乙改造项目","丙改造项目","丁建设项目"});Check(affix.Suffix=="改造项目","T01 modal project suffix");
        AffixAnalyzer.Apply(affix,new[]{"李明","王芳","张志强"});Check(affix.ShortName && affix.Digits==0 && affix.Suffix=="","T01 natural person mode");
        affix=V1Rule(MaskKind.Identifier);AffixAnalyzer.Apply(affix,new[]{"HT-2026-000185","HT-2026-002910","HT-2026-004921"});Check(affix.Prefix=="HT-2026-" && affix.Digits==6,"T02 structured identifier inference");
        affix=V1Rule();affix.Prefix="【昆明】";affix.Suffix="重点工程";affix.Digits=3;
        result=MaskPlanner.Build(V1Snapshot(new DataCell{Row=1,Column=1,Value="原项目甲"}),new[]{affix},CancellationToken.None);
        string generated=(string)result.Output.Cells[0].Value;Check(generated.StartsWith("【昆明】") && generated.EndsWith("重点工程") && System.Text.RegularExpressions.Regex.IsMatch(generated,@"\d{3}重点工程$"),"T03 custom prefix suffix and digits");
        var capacity=V1Rule(MaskKind.Identifier);capacity.Digits=1;
        Reject(()=>MaskPlanner.Build(V1Snapshot(Enumerable.Range(1,11).Select(i=>new DataCell{Row=i,Column=1,Value="CODE"+i}).ToArray()),new[]{capacity},CancellationToken.None),"identifier finite capacity fails without duplicate identities");
        var budget=new TextOutputBudget(new SegmentationOptions{OutputMode=0,Separator=" / "},1);budget.BeginRow();budget.Accept(new string('a',32767));Reject(()=>budget.Accept("b"),"F10 joined character budget checked during token generation");
        budget=new TextOutputBudget(new SegmentationOptions{OutputMode=1},20000);budget.BeginRow();Reject(()=>{for(int i=0;i<100;i++)budget.Accept("word");},"F10 full batch output budget checked incrementally");
        var frame=new FocusFrame{Viewport=new Rectangle(0,0,1000,700),Panes=new[]{new Rectangle(40,25,960,675)},RipplePanes=new[]{new Rectangle(0,0,1000,700)},Cells=new[]{new Rectangle(400,300,100,30)}};
        bool rowHead=false,colHead=false,trough=false;
        var watch=Stopwatch.StartNew();
        for(int i=0;i<120;i++){var f=FocusOverlay.RippleFrame(frame,i/120.0,new ToolPreferences());rowHead|=f.Ripple.Any(r=>r.Bounds.Left<40);colHead|=f.Ripple.Any(r=>r.Bounds.Top<25);trough|=f.Ripple.Any(r=>r.Opacity<0);Check(f.Ripple.All(r=>frame.Viewport.Contains(r.Bounds)&&!frame.Cells[0].IntersectsWith(r.Bounds)),"ripple bounds/active-cell safety "+i);}
        Check(rowHead&&colHead&&trough,"ripple crests/troughs reach both headings");Console.WriteLine("RIPPLE geometry 120 frames ms="+watch.ElapsedMilliseconds);
        var canceled=new CancellationTokenSource();canceled.Cancel();try{MaskPlanner.Build(shorts,new[]{V1Rule()},canceled.Token);Check(false,"cancel");}catch(OperationCanceledException){Check(true,"mask cancellation");}
    }
    private static void V1ExcelTests(bool large)
    {
        string directory=Path.Combine(Path.GetTempPath(),"GridPilot-v1-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(directory);
        int pid=V1ExcelBody(directory,large);CopyExcelExitCheck(pid);Console.WriteLine("V1 artifacts: "+directory);
    }
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static int V1ExcelBody(string directory,bool large)
    {
        dynamic app=null;int pid=0;
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;app.EnableEvents=true;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out pid);Console.WriteLine("V1 Excel version="+app.Version+" build="+app.Build+" pid="+pid);
            var gateway=new ExcelGateway(()=>(object)app);dynamic book=app.Workbooks.Add(-4167),sheet=book.Worksheets[1];
            sheet.Range["A1:A2"].Value2=new object[,]{{"'123.50"},{123.5}};sheet.Range["A1"].Select();gateway.ConvertSelection(false,new ToolPreferences());Check(gateway.CanUndo,"F02 actual text conversion creates undo");sheet.Rows[1].Delete();
            System.Windows.Forms.Application.DoEvents();Check(!gateway.CanUndo,"F02 external row deletion invalidates undo");Reject(()=>gateway.UndoLast(),"F02 invalidated undo explains the refusal");Check(sheet.Range["A1"].Value2 is double,"F02 shifted equal-valued cell remains numeric");
            sheet.Range["A1:A3"].ClearContents();sheet.Range["A1"].Value2=10.0;dynamic other=book.Worksheets.Add(Type.Missing,sheet);other.Name="Other";other.Range["A1"].Value2=20.0;sheet.Range["A2"].Formula="=SUBTOTAL(9,Other!A1)";sheet.Range["A3"].Formula="=SUBTOTAL(9,A1:A2)";sheet.Activate();sheet.Range["A1:A3"].Select();
            var copied=gateway.CopySelection(true,new CopyOptions{CellFormats=false,FormulaFallback=0});dynamic copiedBook=copied.Workbook;Check((double)copiedBook.Worksheets[1].Range["A3"].Value2==10 && copied.ValueFormulas==2,"F03 nested subtotal drift freezes to source cache");copiedBook.Close(false);book.Activate();sheet.Activate();sheet.Range["A1:A3"].Select();Reject(()=>gateway.CopySelection(true,new CopyOptions{FormulaFallback=1}),"F03 strict copy rejects unsafe source formula");book.Close(false);
            book=app.Workbooks.Add(-4167);sheet=book.Worksheets[1];sheet.Range["A1:B3"].Value2=new object[,]{{"姓名","类型"},{"FF","类别"},{"'1","类别"}};sheet.Range["A2:B3"].Font.Color=255;sheet.Range["B2"].NumberFormat="\"FF\"0";
            sheet.Range["XFD1048576"].Value2="boundary-marker";sheet.Range["AZ500000"].Value2="sparse-marker";book.Saved=true;
            var timer=Stopwatch.StartNew();var snapshot=gateway.CaptureMaskWorkbook();Console.WriteLine("SPARSE scan ms="+timer.ElapsedMilliseconds);Check(snapshot.Cells.Count==8 && snapshot.Sheets[0].NonEmptyRows==5 && snapshot.Sheets[0].NonEmptyColumns==4,"T04 sparse full-grid boundary and nonempty counts");Check((bool)book.Saved,"scan leaves source saved state intact");
            var result=MaskPlanner.Build(snapshot,new[]{new MaskRule{Column=1,StartRow=2,Kind=MaskKind.ChineseName,Domain="person"}},CancellationToken.None,true);
            string path=Path.Combine(directory,"styles-boundary.xlsx");gateway.ExportMask(snapshot,result,path,CancellationToken.None,true);
            dynamic output=app.Workbooks.Open(path,0,true);Check((int)output.Worksheets[1].Range["A2"].Font.Color==255,"F04 short identities do not corrupt ARGB styles");Check((string)output.Worksheets[1].Range["XFD1048576"].Value2=="boundary-marker","full-grid final cell survives formatted export");output.Close(false);book.Activate();
            for(int i=0;i<20;i++){timer.Restart();gateway.ExportMask(snapshot,result,Path.Combine(directory,"repeat-"+i+".xlsx"),CancellationToken.None,false);Console.WriteLine("REPEAT "+i+" ms="+timer.ElapsedMilliseconds+" "+ExcelGateway.LastMaskPerformance);Check(timer.ElapsedMilliseconds<45000,"P00 repeated export avoids 60s stall "+i);CopyExcelExitCheck(ExcelGateway.LastMaskWriterProcess);}
            string canceledPath=Path.Combine(directory,"canceled.xlsx");
            using(var cancellation=new CancellationTokenSource())
            {
                cancellation.CancelAfter(600);
                try{gateway.ExportMask(snapshot,result,canceledPath,cancellation.Token);Check(false,"canceled export");}
                catch(OperationCanceledException){Check(!File.Exists(canceledPath),"canceled export never commits a partial file");}
                CopyExcelExitCheck(ExcelGateway.LastMaskWriterProcess);
            }
            string stalePath=Path.Combine(directory,"stale.xlsx");
            using(var mutate=new System.Windows.Forms.Timer{Interval=600})
            {
                mutate.Tick+=(sender,args)=>{mutate.Stop();sheet.Range["B2"].Value2="changed while exporting";};
                mutate.Start();Reject(()=>gateway.ExportMask(snapshot,result,stalePath,CancellationToken.None),"source edits during background export block commit");
                Check(!File.Exists(stalePath),"changed source leaves no delivered file");
            }
            book.Close(false);
            if(large)
            {
                book=app.Workbooks.Add(-4167);sheet=book.Worksheets[1];
                for(int start=1;start<=50000;start+=1000)
                {var values=new object[1000,50];for(int r=0;r<1000;r++)for(int c=0;c<50;c++)values[r,c]=c==0?(object)("客户原文"+((start+r)%2000)):((double)(start+r)*100+c);sheet.Range["A"+start+":AX"+(start+999)].Value2=values;}
                sheet.Range["XFD1048576"].Value2="edge";book.Saved=true;
                timer.Restart();snapshot=gateway.CaptureMaskWorkbook();Console.WriteLine("LARGE scan count="+snapshot.Cells.Count+" ms="+timer.ElapsedMilliseconds+" managed="+GC.GetTotalMemory(false));Check(snapshot.Cells.Count==2500001,"T05 scans all 2.5 million cells plus XFD boundary");
                timer.Restart();result=MaskPlanner.Build(snapshot,new[]{new MaskRule{Column=1,StartRow=1,Kind=MaskKind.ChineseName,Domain="customer"}},CancellationToken.None,true);Console.WriteLine("LARGE plan ms="+timer.ElapsedMilliseconds+" managed="+GC.GetTotalMemory(false));Check(result.Changed==50000,"T05 masks every selected row, no preview truncation");
                timer.Restart();gateway.ExportMask(snapshot,result,Path.Combine(directory,"large-complete.xlsx"),CancellationToken.None);Console.WriteLine("LARGE export ms="+timer.ElapsedMilliseconds+" "+ExcelGateway.LastMaskPerformance);Check((bool)book.Saved,"large export leaves source unchanged");CopyExcelExitCheck(ExcelGateway.LastMaskWriterProcess);
                timer.Restart();gateway.ExportMask(snapshot,result,Path.Combine(directory,"large-formatted.xlsx"),CancellationToken.None,true);Console.WriteLine("LARGE formatted export ms="+timer.ElapsedMilliseconds+" managed="+GC.GetTotalMemory(false)+" "+ExcelGateway.LastMaskPerformance);Check(File.Exists(Path.Combine(directory,"large-formatted.xlsx")) && (bool)book.Saved,"2.5 million cells roundtrip with streamed formatting");book.Close(false);CopyExcelExitCheck(ExcelGateway.LastMaskWriterProcess);
            }
            gateway.ReleaseUndoEvents();
        }
        finally{if(app!=null){try{app.Quit();}catch{}System.Runtime.InteropServices.Marshal.ReleaseComObject(app);}}
        return pid;
    }
}
