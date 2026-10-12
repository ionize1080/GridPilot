using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using SheetPilot.Core;
using SheetPilot.Excel;
using SheetPilot.UI;

internal static partial class Program
{
    private static void Reject(Action action,string label)
    { try {action();}catch(InvalidOperationException){Check(true,label);return;}throw new Exception("FAILED: "+label); }
    private static MaskSnapshot SampleMask()
    {
        var s=new MaskSnapshot();s.Sheets.Add(new DataSheet {Name="客户表",Rows=4,Columns=5});s.Sheets.Add(new DataSheet {Name="贷款表",Rows=4,Columns=5});
        Action<int,int,int,object,string> add=(sheet,row,col,value,formula)=>s.Cells.Add(new DataCell {Sheet=sheet,Row=row,Column=col,Value=value,Formula=formula});
        add(0,1,1,"客户",null);add(0,1,2,"信用代码",null);add(0,1,3,"金额",null);add(0,1,4,"绿色贷款",null);
        add(0,2,1,"云川贸易有限公司",null);add(0,2,2,"00123456789012345678",null);add(0,2,3,100.0,null);add(0,2,4,"是",null);
        add(0,3,1,"海原科技有限公司",null);add(0,3,2,"00223456789012345678",null);add(0,3,3,200.0,null);add(0,3,4,"否",null);
        add(0,4,3,300.0,"=SUM(C2:C3)");
        add(1,1,1,"客户",null);add(1,2,1,"云川贸易有限公司",null);add(1,2,2,"00123456789012345678","='客户表'!B2");
        add(1,2,3,100.0,"=VLOOKUP(A2,'客户表'!A2:C3,3,FALSE)");add(1,2,4,"是","=IF(A2=\"云川贸易有限公司\",\"是\",\"否\")");
        add(1,3,3,100.0,"=SUMIFS('客户表'!C2:C3,'客户表'!A2:A3,A2)");
        return s;
    }
    private static MaskRule[] SampleRules()=>new[] {
        new MaskRule {Sheet=0,Column=1,Kind=MaskKind.ChineseName,Domain="客户"},new MaskRule {Sheet=1,Column=1,Kind=MaskKind.ChineseName,Domain="客户"},
        new MaskRule {Sheet=0,Column=2,Kind=MaskKind.Identifier,Domain="信用代码"},
        new MaskRule {Sheet=0,Column=3,Kind=MaskKind.Amount,Domain="金额"},new MaskRule {Sheet=1,Column=3,Kind=MaskKind.Amount,Domain="金额"}};
    private static void ReleaseFeatureTests()
    {
        var segmenter=new TextSegmenter(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Resources"));var o=new SegmentationOptions();
        string text="绿色贷款 IFRS 9 MAS 610 email@example.com 2026-10-12 12.5% loan-to-value don't ABC001 绿色贷款";
        var words=segmenter.Split(text,o,CancellationToken.None);
        foreach(string token in new[]{"绿色贷款","IFRS 9","MAS 610","email@example.com","2026-10-12","12.5%","loan-to-value","don't","ABC001"})Check(words.Any(w=>w.Text==token),"segmentation protects "+token);
        Check(words.Count(w=>w.Text=="绿色贷款")==2,"segmentation preserves repeats");
        Check(words.All(w=>text.Substring(w.Offset,w.Text.Length)==w.Text),"segmentation offsets refer to original text");
        foreach(string token in new[]{"001ABC","A-01","123ABC-02","https://example.com/a?q=1","1,234.50%"})
            Check(segmenter.Split(token,o,CancellationToken.None).Single().Text==token,"segmentation preserves business token "+token);
        Check(segmenter.Split("中华人民共和国银行贷款",o,CancellationToken.None).Count<12,"Chinese dictionary groups words");
        o.Distinct=true;Check(segmenter.Split(text,o,CancellationToken.None).Count(w=>w.Text=="绿色贷款")==1,"optional stable deduplication");
        o.KeepPunctuation=true;Check(segmenter.Split("贷款，🙂",o,CancellationToken.None).Any(w=>w.Text=="🙂"),"emoji is not split into surrogates");
        var rows=new[]{new TextOutput {Source=new TextInput {Sheet="Test",Address="A1",Text=text},Words=words}};
        for(int mode=0;mode<3;mode++){o.OutputMode=mode;var layout=TextBatch.Layout(rows,o);Check(layout.GetLength(0)==(mode==2?words.Count+1:2),"output mode "+mode+" capacity");}
        var cancel=new CancellationTokenSource();cancel.Cancel();try{segmenter.Split(text,o,cancel.Token);Check(false,"segmentation cancellation");}catch(OperationCanceledException){Check(true,"segmentation cancellation");}
        var source=SampleMask();var rules=SampleRules();var result=MaskPlanner.Build(source,rules,CancellationToken.None);var cells=result.Output.Cells.ToDictionary(c=>c.Key);
        Check((string)cells["0:2:1"].Value==(string)cells["1:2:1"].Value,"mask cross-sheet name identity");
        Check(!Equals(cells["0:2:1"].Value,cells["0:3:1"].Value),"mask names remain distinct");
        Check((string)source.Cells.First(c=>c.Key=="0:2:1").Value=="云川贸易有限公司","mask planner does not mutate source");
        Check((double)cells["0:2:3"].Value==170 && (double)cells["0:3:3"].Value==340,"amount shared factor");
        Check((double)result.Calculated["0:4:3"]==510 && (double)result.Calculated["1:2:3"]==170 && (double)result.Calculated["1:3:3"]==170,"sum lookup and SUMIFS relations");
        Check((string)result.Calculated["1:2:4"]=="是" && (string)cells["0:2:4"].Value=="是","green-loan constants and formula classification preserved");
        Check((string)result.Calculated["1:2:2"]==(string)cells["0:2:2"].Value,"formula identifiers use source mapping");
        Check(cells["1:2:4"].Formula.Contains((string)cells["0:2:1"].Value),"formula sensitive literal rewritten");
        var changed=SampleMask();changed.Cells.Add(new DataCell {Sheet=1,Row=4,Column=4,Value="小额",Formula="=IF('客户表'!C2<150,\"小额\",\"大额\")"});
        Reject(()=>MaskPlanner.Build(changed,rules,CancellationToken.None),"fixed threshold branch changes block export");
        changed=SampleMask();changed.Cells.Add(new DataCell {Sheet=1,Row=4,Column=4,Value="云川贸易有限公司",Formula=null});
        Reject(()=>MaskPlanner.Build(changed,rules,CancellationToken.None),"unselected sensitive occurrences block export");
        changed=SampleMask();changed.Cells.Add(new DataCell {Sheet=1,Row=4,Column=4,Value=100.0,Formula="=INDIRECT(\"C2\")"});
        Reject(()=>MaskPlanner.Build(changed,rules,CancellationToken.None),"dynamic references block export");
        changed=SampleMask();changed.Cells.First(c=>c.Key=="0:2:2").Value=123456789012345678.0;
        Reject(()=>MaskPlanner.Build(changed,rules,CancellationToken.None),"numeric long identifiers rejected");
        changed=SampleMask();changed.Cells.First(c=>c.Key=="1:3:3").Formula="=SUMIFS('客户表'!C2:C3,'客户表'!A2:A3,\"云川*\")";
        Reject(()=>MaskPlanner.Build(changed,rules,CancellationToken.None),"wildcard criteria blocked");
        changed=SampleMask();changed.Cells.First(c=>c.Key=="0:4:3").Value=301.0;
        Reject(()=>MaskPlanner.Build(changed,rules,CancellationToken.None),"stale source formula cache blocked");
        changed=SampleMask();changed.Cells.Add(new DataCell {Sheet=1,Row=4,Column=4,Value="ok",Formula="=IF(TRUE,\"ok\",INDIRECT(\"A1\"))"});
        Reject(()=>MaskPlanner.Build(changed,rules,CancellationToken.None),"unsupported function in inactive branch blocked");
        var lookup=new MaskSnapshot();lookup.Sheets.Add(new DataSheet {Name="Codes",Rows=3,Columns=4});
        lookup.Cells.AddRange(new[]{new DataCell {Row=1,Column=1,Value="00123"},new DataCell {Row=2,Column=1,Value="00456"},new DataCell {Row=1,Column=2,Value=10.0},new DataCell {Row=2,Column=2,Value=20.0},new DataCell {Row=1,Column=3,Value="00123"},new DataCell {Row=1,Column=4,Value=10.0,Formula="=VLOOKUP(C1,A1:B2,2,FALSE)"},new DataCell {Row=2,Column=4,Value=1.0,Formula="=MATCH(C1,A1:A2,0)"}});
        var codeRules=new[]{new MaskRule {Column=1,StartRow=1,Kind=MaskKind.Identifier,Domain="code"},new MaskRule {Column=3,StartRow=1,Kind=MaskKind.Identifier,Domain="code"}};
        var looked=MaskPlanner.Build(lookup,codeRules,CancellationToken.None);
        Check((double)looked.Calculated["0:1:4"]==10 && (double)looked.Calculated["0:2:4"]==1,"exact lookup preserves numeric text identifiers");
        lookup.Cells[0].Value="Acme";lookup.Cells[4].Value="ACME";
        looked=MaskPlanner.Build(lookup,codeRules,CancellationToken.None);
        Check(Equals(looked.Output.Cells[0].Value,looked.Output.Cells[4].Value),"identity mapping preserves Excel case-insensitive equality");
        var dates=new MaskSnapshot();dates.Sheets.Add(new DataSheet {Name="Dates",Rows=2,Columns=2});
        dates.Cells.AddRange(new[]{new DataCell {Row=1,Column=1,Value=45000.0},new DataCell {Row=2,Column=1,Value=45030.0},new DataCell {Row=1,Column=2,Value=30.0,Formula="=A2-A1"}});
        var dateRules=new[]{new MaskRule {Column=1,StartRow=1,Kind=MaskKind.Date,Domain="dates",DayShift=37}};
        var dated=MaskPlanner.Build(dates,dateRules,CancellationToken.None);
        Check((double)dated.Output.Cells[0].Value==45037 && (double)dated.Calculated["0:1:2"]==30,"date shift preserves duration");
        dates.Cells[0].Value=60.0;Reject(()=>MaskPlanner.Build(dates,dateRules,CancellationToken.None),"fictional 1900 leap day is rejected");
        dates.Date1904=true;dates.Cells[0].Value=2957003.0;Reject(()=>MaskPlanner.Build(dates,dateRules,CancellationToken.None),"1904 date system upper bound rejected");
        changed=SampleMask();changed.Cells.First(c=>c.Key=="0:2:3").Value=1e-30;
        Reject(()=>MaskPlanner.Build(changed,rules,CancellationToken.None),"amount underflow cannot silently turn a nonzero amount into zero");
        var deep=SampleMask();deep.Cells.Add(new DataCell {Sheet=1,Row=4,Column=4,Value=300.0,Formula="="+string.Join("+",Enumerable.Repeat("1",300))});
        Reject(()=>new FormulaWorkbook(deep).Evaluate("1:4:4"),"long arithmetic chain rejects before exhausting call stack");
        var cancelled=new CancellationTokenSource();cancelled.Cancel();
        try {new FormulaWorkbook(source,false,cancelled.Token).Evaluate("0:4:3");Check(false,"formula cancellation");}catch(OperationCanceledException){Check(true,"formula validation accepts cancellation");}
        var prefs=new ToolPreferences {FocusAnimation=true};var a=new FocusFrame {Viewport=new System.Drawing.Rectangle(0,0,800,600),Panes=new[]{new System.Drawing.Rectangle(0,0,800,600)},Cells=new[]{new System.Drawing.Rectangle(30,30,50,20)}};
        var b=new FocusFrame {Viewport=a.Viewport,Panes=a.Panes,Cells=new[]{new System.Drawing.Rectangle(60,50,50,20)}};
        Check(FocusOverlay.CanInterpolate(a,b),"focus nearby transition allowed");
        foreach(double t in new[]{0.0,.25,.5,.75})Check(FocusOverlay.Interpolate(a,b,t,prefs).Bands.All(r=>!r.IntersectsWith(b.Cells[0])),"focus actual cell stays clear at "+t);
        Check(ReferenceEquals(FocusOverlay.Interpolate(a,b,1,prefs),b),"focus finite animation endpoint");
    }
    private static void ReleaseFeatureRenderTests(string directory)
    {
        var colors=new NavigationSettings();var sample=SampleMask();
        var mask=new MaskWindow(colors,sample,ExcelGateway.SuggestMaskRules(sample),(result,path,cancel)=>path);
        RenderWindowContent(mask,Path.Combine(directory,"mask-fields.png"),1100,790,96);mask.Close();
        var segment=new SegmentationWindow(colors,new[]{new TextInput {Sheet="示例",Address="A2",Text="绿色贷款 IFRS 9 MAS 610"}},Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"Resources"),values=>{});
        RenderWindowContent(segment,Path.Combine(directory,"segmentation.png"),930,730,96);segment.Close();
    }
    private static void ReleaseFeatureExcelTests()
    {
        int pid=ReleaseFeatureExcelBody();
        CopyExcelExitCheck(pid);
        if(ExcelGateway.LastMaskWriterProcess>0)CopyExcelExitCheck(ExcelGateway.LastMaskWriterProcess);
    }
    private static void ReleaseUiSmoke(string addin,string directory)
    {
        Directory.CreateDirectory(directory);dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));
        try
        {
            app.Visible=true;app.DisplayAlerts=false;app.AutomationSecurity=3;app.WindowState=-4137;
            int smokePid;GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out smokePid);Console.WriteLine("UI test owned Excel pid="+smokePid);
            dynamic book=app.Workbooks.Add(-4167),sheet=book.Worksheets[1];sheet.Name="发布验收";
            sheet.Range["A1:C3"].Value2=new object[,]{{"客户","代码","描述"},{"云川贸易有限公司","00123","绿色贷款 IFRS 9 MAS 610"},{"海原科技有限公司","00456","贷款余额 12.5% ABC001"}};
            sheet.Columns.ColumnWidth=28;sheet.Range["C2:C3"].Select();
            Check((bool)app.RegisterXLL(addin),"release UI packed XLL registration");
            Console.WriteLine("Ready for UI smoke test");
            var timer=new System.Windows.Threading.DispatcherTimer {Interval=TimeSpan.FromMilliseconds(200)};
            var frame=new System.Windows.Threading.DispatcherFrame();var elapsed=System.Diagnostics.Stopwatch.StartNew();
            timer.Tick+=(s,e)=>
            {
                if(elapsed.Elapsed.TotalMinutes>10){frame.Continue=false;return;}
                string path=Path.Combine(directory,"command.txt");if(!File.Exists(path))return;
                string command=File.ReadAllText(path).Trim();File.Delete(path);
                if(command=="close") {frame.Continue=false;return;}
                if(command!="mask" && command!="segment")return;
                book.Activate();sheet.Activate();sheet.Range["C2:C3"].Select();
                app.Run(command=="mask"?"GridPilot_Mask":"GridPilot_Segment");
                Console.WriteLine("Opened "+command);
            };
            timer.Start();System.Windows.Threading.Dispatcher.PushFrame(frame);timer.Stop();
        }
        finally {try{app.Quit();}catch{} }
    }
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static int ReleaseFeatureExcelBody()
    {
        dynamic app=null;int pid=0;string directory=Path.Combine(Path.GetTempPath(),"GridPilot-release-tests-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(directory);
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out pid);Console.WriteLine("Feature test Excel "+app.Version+" build "+app.Build+"; pid="+pid);
            var source=SampleMask();dynamic book=app.Workbooks.Add(-4167);book.Worksheets[1].Name=source.Sheets[0].Name;book.Worksheets.Add(Type.Missing,book.Worksheets[1]).Name=source.Sheets[1].Name;
            foreach(var cell in source.Cells) {dynamic target=book.Worksheets[cell.Sheet+1].Cells[cell.Row,cell.Column];if(cell.Formula!=null)target.Formula=cell.Formula;else target.Value2=cell.Value is string value?"'"+value:cell.Value;}
            book.Worksheets[2].Visible=2;app.CalculateFullRebuild();book.Worksheets[1].Activate();book.Saved=true;
            var gateway=new ExcelGateway(()=>(object)app);var captured=gateway.CaptureMaskWorkbook();Check(captured.Sheets.Count==2,"mask reads VeryHidden sheet");
            var result=MaskPlanner.Build(captured,SampleRules(),CancellationToken.None);string targetPath=Path.Combine(directory,"masked.xlsx");
            gateway.ExportMask(captured,result,targetPath,CancellationToken.None);Check(File.Exists(targetPath),"mask verifies saves reopens and commits new XLSX");
            using(var zip=System.IO.Compression.ZipFile.OpenRead(targetPath))using(var stream=zip.GetEntry("docProps/core.xml").Open())
            {
                var metadata=System.Xml.Linq.XDocument.Load(stream);
                Check(metadata.Root.Elements().Where(e=>e.Name.LocalName!="created" && e.Name.LocalName!="modified").All(e=>string.IsNullOrEmpty(e.Value)),"saved output removes author metadata that Excel recreates during SaveAs");
            }
            Check((bool)book.Saved && (string)book.Worksheets[1].Range["A2"].Value2=="云川贸易有限公司","mask original values and Saved flag unchanged");
            Reject(()=>gateway.ExportMask(captured,result,targetPath,CancellationToken.None),"mask refuses existing target");
            book.Worksheets[1].Range["D2"].Value2="Changed";
            Reject(()=>gateway.ExportMask(captured,result,Path.Combine(directory,"changed.xlsx"),CancellationToken.None),"source changes after preview block output");
            Check(!File.Exists(Path.Combine(directory,"changed.xlsx")),"failed validation leaves no output");
            book.Worksheets[1].Range["F1"].Value2="绿色贷款 IFRS 9";book.Worksheets[1].Range["F1"].Select();
            Check(gateway.CaptureTextSelection().Count==1,"text snapshot from actual Excel");
            gateway.WriteTextOutput(new object[,] {{"原文","结果"},{"=1+1","00123"}});dynamic textBook=app.ActiveWorkbook;
            Check(!(bool)textBook.Worksheets[1].Range["A2"].HasFormula && (string)textBook.Worksheets[1].Range["A2"].Value2=="=1+1","segmentation output does not execute formula-looking text");
            textBook.Close(false);book.Close(false);
        }
        finally {if(app!=null){try{app.Quit();}catch{}try{System.Runtime.InteropServices.Marshal.ReleaseComObject(app);}catch{}}}
        return pid;
    }
}
