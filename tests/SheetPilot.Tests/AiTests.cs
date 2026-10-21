using System;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using SheetPilot.Core;
using SheetPilot.Excel;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;

internal static partial class Program
{
    private delegate bool AiEnumWindow(IntPtr window, IntPtr data);
    [DllImport("user32.dll", EntryPoint="EnumChildWindows")] private static extern bool AiEnumChildren(IntPtr parent, AiEnumWindow callback, IntPtr data);
    [DllImport("user32.dll", EntryPoint="GetWindowThreadProcessId")] private static extern uint AiWindowProcess(IntPtr window, out uint process);
    [DllImport("user32.dll", EntryPoint="GetClassName", CharSet=CharSet.Unicode)] private static extern int AiClass(IntPtr window, StringBuilder text, int size);
    [DllImport("user32.dll", EntryPoint="GetWindowLong")] private static extern int AiStyle(IntPtr window, int index);
    [DllImport("user32.dll", EntryPoint="GetWindow")] private static extern IntPtr AiOwner(IntPtr window,uint command);
    [DllImport("user32.dll", EntryPoint="PostMessage")] private static extern bool AiPost(IntPtr window, int message, IntPtr w, IntPtr l);
    [DllImport("user32.dll", EntryPoint="EnumWindows")] private static extern bool AiEnumTop(AiEnumWindow callback,IntPtr data);
    private static void AiHostTest(string addin, int mode, bool installed = false)
    {
        dynamic app = null; Process browser = null;
        string previous = Environment.GetEnvironmentVariable("GRIDPILOT_CHATBOX_PROFILE");
        try
        {
            string directory = System.IO.Path.GetDirectoryName(addin);
            Environment.SetEnvironmentVariable("GRIDPILOT_CHATBOX_PROFILE", installed ? System.IO.Path.Combine(System.IO.Path.GetTempPath(), "GridPilot-installed-test-"+Guid.NewGuid().ToString("N")) : System.IO.Path.Combine(directory, "test-profile"));
            if(!installed)new ToolPreferencesStore(System.IO.Path.Combine(directory, "GridPilot.tools.json")).Save(new ToolPreferences { AiDisplayMode = mode });
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            app.Visible = true; app.DisplayAlerts = false; app.AutomationSecurity = 3;
            dynamic book = app.Workbooks.Add(-4167), sheet = book.Worksheets[1]; sheet.Name = "GridPilot AI 合成验证";
            sheet.Range["A1:C4"].Value2 = new object[,] { { "项目", "收入", "成本" }, { "甲", 100.0, 40.0 }, { "乙", 200.0, 90.0 }, { "丙", 150.0, 70.0 } };
            sheet.Range["D2"].Formula = "=B2-C2"; sheet.Range["D2"].Select();
            Check((bool)app.RegisterXLL(addin), "AI host XLL registration");
            var before = new HashSet<int>(Process.GetProcessesByName("electron").Select(p => p.Id));
            app.Run("GridPilot_AI");
            IntPtr chrome = IntPtr.Zero, excelWindow = new IntPtr((int)app.Hwnd);
            PumpUntil(() =>
            {
                chrome=IntPtr.Zero;browser=null;
                browser=Process.GetProcessesByName("electron").FirstOrDefault(p=>!before.Contains(p.Id)&&p.MainWindowHandle!=IntPtr.Zero&&p.MainWindowTitle=="Chatbox");
                if(browser!=null)chrome=browser.MainWindowHandle;
                // Owned windows are absent from Process.MainWindowHandle on Windows.
                if(chrome==IntPtr.Zero)AiEnumTop((window,unused)=>
                {
                    var text=new StringBuilder(128);AiClass(window,text,text.Capacity);AiWindowProcess(window,out uint pid);
                    if(text.ToString()=="Chrome_WidgetWin_1"&&!before.Contains((int)pid)&&AiOwner(window,4)==excelWindow)
                    {chrome=window;browser=Process.GetProcessById((int)pid);return false;}return true;
                },IntPtr.Zero);
                return chrome != IntPtr.Zero && (mode!=0||AiOwner(chrome,4)==excelWindow);
            }, 45000);
            Check(chrome != IntPtr.Zero, mode == 0 ? "native Chatbox is docked to actual Excel task pane" : "native Chatbox opens as independent window");
            Check((AiStyle(chrome,-16)&0x40000000)==0,"Chromium remains a keyboard-capable top-level window");
            Check((AiOwner(chrome,4)==excelWindow)==(mode==0),"Excel owns docked window only in sidebar mode");
            Console.WriteLine("AI host window=" + chrome + " process=" + browser.Id + " title=" + browser.MainWindowTitle);
            app.Run("GridPilot_AI"); // Reopen activates the same process/pane.
            AiPost(chrome, 0x0010, IntPtr.Zero, IntPtr.Zero);
            PumpUntil(() => browser.HasExited, 20000);
            Check(browser.HasExited, "Chatbox close disconnects bridge and releases pane");
            book.Close(false);
        }
        finally
        {
            if (browser != null) { try { if (!browser.HasExited) browser.CloseMainWindow(); } catch { } browser.Dispose(); }
            if (app != null) { try { app.Quit(); } catch { } Marshal.ReleaseComObject(app); }
            Environment.SetEnvironmentVariable("GRIDPILOT_CHATBOX_PROFILE", previous);
        }
    }
    private static void AiBatchTests(AiOperations ai,dynamic app,dynamic sheet,string id)
    {
        string name=sheet.Name;
        Func<string,string,JObject,JObject> propose=(kind,address,extra)=>{extra["snapshotId"]=id;extra["sheet"]=name;extra["address"]=address;extra["kind"]=kind;return JObject.FromObject(ai.Propose(extra));};
        var values=propose("values","J2:K3",new JObject{["data"]=new JArray(new JArray("=literal",12),new JArray("safe",true))});
        var formulas=propose("formulas","L2:L3",new JObject{["data"]=new JArray(new JArray("=K2*2"),new JArray("=SUM(K2:K3)"))});
        Check(sheet.Range["J2"].Value2==null,"batch proposals do not write");
        ai.ApplyBatch(new[]{(string)values["proposalId"],(string)formulas["proposalId"]});
        Check((string)sheet.Range["J2"].Value2=="=literal"&&!(bool)sheet.Range["J2"].HasFormula&&(double)sheet.Range["L2"].Value2==24,"one action writes literal values and formulas");
        ai.Apply(null,true);Check(sheet.Range["J2"].Value2==null&&sheet.Range["L2"].Value2==null,"undo entire batch");
        values=propose("values","J2:K2",new JObject{["data"]=new JArray{new JArray{8,9}}});
        formulas=propose("formula","L2",new JObject{["formula"]="=1+1"});sheet.Range["L2"].Value2="manual";
        bool conflict=false;try{ai.ApplyBatch(new[]{(string)values["proposalId"],(string)formulas["proposalId"]});}catch(InvalidOperationException){conflict=true;}
        Check(conflict&&sheet.Range["J2"].Value2==null&&(string)sheet.Range["L2"].Value2=="manual","preflight rejects entire batch before first write");
        var rollbackValues=propose("values","J5:K5",new JObject{["data"]=new JArray{new JArray{7,8}}});
        var invalid=propose("formula","L5",new JObject{["formula"]="=SUM("});
        bool rolledBack=false;try{ai.ApplyBatch(new[]{(string)rollbackValues["proposalId"],(string)invalid["proposalId"]});}catch(InvalidOperationException){rolledBack=true;}
        Check(rolledBack&&sheet.Range["J5"].Value2==null&&sheet.Range["K5"].Value2==null,"failed later write rolls back earlier batch values");
        var style=propose("format","M2:N3",new JObject{["style"]=new JObject{["bold"]=true,["fill_color"]="#D9EFE5",["number_format"]="0.00"}});
        ai.Apply((string)style["proposalId"]);Check((bool)sheet.Range["M2"].Font.Bold&&(string)sheet.Range["N3"].NumberFormat=="0.00","native range formatting");
        ai.Apply(null,true);Check(!(bool)sheet.Range["M2"].Font.Bold&&(int)sheet.Range["M2"].Interior.Pattern==-4142,"format undo restores no-fill pattern");
        var chart=propose("chart","A1:C4",new JObject{["chartType"]="column",["title"]="Synthetic chart"});
        ai.Apply((string)chart["proposalId"]);Check((int)sheet.ChartObjects().Count==1,"native chart creation");
        sheet.ChartObjects(1).Chart.ChartTitle.Text="Manual title";
        bool chartConflict=false;try{ai.Apply(null,true);}catch(InvalidOperationException){chartConflict=true;}
        Check(chartConflict&&(int)sheet.ChartObjects().Count==1,"undo refuses a manually edited chart");
        sheet.ChartObjects(1).Chart.ChartTitle.Text="Synthetic chart";
        ai.Apply(null,true);Check((int)sheet.ChartObjects().Count==0,"chart batch undo");
        var meta=JObject.FromObject(ai.Metadata(id));Check(meta["sheets"].Any(),"workbook structure inspection");
    }
    private static void AiUiLaunch(string addin)
    {
        string directory=System.IO.Path.GetDirectoryName(addin);
        Environment.SetEnvironmentVariable("GRIDPILOT_CHATBOX_PROFILE",System.IO.Path.Combine(directory,"ui-profile"));
        new ToolPreferencesStore(System.IO.Path.Combine(directory,"GridPilot.tools.json")).Save(new ToolPreferences{AiDisplayMode=1});
        dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));
        try
        {
            app.Visible=true;app.DisplayAlerts=false;dynamic book=app.Workbooks.Add(-4167),sheet=book.Worksheets[1];
            sheet.Name="AI验证_合成数据";sheet.Range["A1:C4"].Value2=new object[,]{{"项目","收入","成本"},{"甲",100.0,40.0},{"乙",200.0,90.0},{"丙",150.0,70.0}};
            sheet.Range["E8"].Value2="焦点保护哨兵";sheet.Range["E8"].Select();
            app.RegisterXLL(addin);app.Run("GridPilot_AI");
            Console.WriteLine("READY synthetic Excel for computer-use");
            string stop=System.IO.Path.Combine(directory,"stop-ui-test");
            while(!System.IO.File.Exists(stop)){System.Windows.Forms.Application.DoEvents();System.Threading.Thread.Sleep(50);}
            Check((string)sheet.Range["E8"].Value2=="焦点保护哨兵","computer-use typing never modified active Excel cell");
            book.Close(false);
        }finally{app.Quit();Marshal.ReleaseComObject(app);}
    }
    private static void AiTests(bool excel)
    {
        System.IO.Directory.CreateDirectory("docs/v1.8");
        using(var icon=RibbonIcons.Draw("GridPilot_AI",96))icon.Save("docs/v1.8/ai-icon.png");
        Check(!new ToolPreferences{AiEnabled=false}.IsEnabled("GridPilot_AI"),"AI can be disabled");
        Check(AiRange.Parse("XFD1048576").Size==1,"maximum Excel address valid");
        foreach(string address in new[]{"XFE1","A1048577","A0","A:A","[other]A1"}){bool bad=false;try{AiRange.Parse(address);}catch(InvalidOperationException){bad=true;}Check(bad,"reject invalid range "+address);}
        Check(new ToolPreferences().IsEnabled("GridPilot_AI"), "AI visible on fresh and upgraded preferences");
        var preferences = new ToolPreferences { AiDisplayMode = 98 }; preferences.Normalize();
        Check(preferences.AiDisplayMode == 0, "invalid mode normalizes to sidebar");
        AiOperations.ValidateFormula("=SUM(A1:A9)");
        foreach (var formula in new[] { "=WEBSERVICE(\"https://example.com\")", "=cmd|'/C calc'!A0", "=SUM([other.xlsx]Sheet1!A1)", "plain", "=HYPERLINK(\"url\")", "=RTD(\"x\",,\"a\")", "=CALL(\"x\")" })
        {
            bool rejected = false; try { AiOperations.ValidateFormula(formula); } catch (InvalidOperationException) { rejected = true; }
            Check(rejected, "reject active/external formula: " + formula.Split('(')[0]);
        }
        if (!excel) return;
        dynamic app = null;
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            app.Visible = false; app.DisplayAlerts = false; app.AutomationSecurity = 3;
            dynamic book = app.Workbooks.Add(-4167), sheet = book.Worksheets[1]; sheet.Name = "合成数据";
            sheet.Range["A1:C4"].Value2 = new object[,] { { "项目", "金额", "说明" }, { "甲", 10.0, "第一笔" }, { "乙", 20.0, "第二笔" }, { "丙", 30.0, null } };
            sheet.Range["D2"].Formula = "=SUM(B2:B4)";
            sheet.Range["F2"].NumberFormat = "@"; sheet.Range["F2"].Value2 = "=literal";
            using (var ai = new AiOperations(() => (object)app))
            {
                sheet.Range["A1:D4"].Select();
                var captured = JObject.FromObject(ai.Capture("selection"));
                string id = (string)captured["snapshotId"];
                Check((long)captured["scannedCells"] == 16 && (int)captured["nonEmptyCells"] == 12, "selection capture excludes blank cells");
                Check(captured["sample"].Any(c => (string)c["formula"] == "=SUM(B2:B4)"), "capture actual formula with address");
                var stats = JObject.FromObject(ai.Read(new JObject { ["snapshotId"] = id, ["operation"] = "statistics", ["column"] = 2 }));
                Check((double)stats["statistics"]["sum"] == 60 && (int)stats["statistics"]["numericCount"] == 3, "exact numeric statistics scoped to column");
                var find = JObject.FromObject(ai.Read(new JObject { ["snapshotId"] = id, ["operation"] = "find", ["comparison"] = "greater", ["number"] = 15, ["column"] = 2 }));
                Check((int)find["total"] == 2 && (string)find["cells"][0]["address"] == "B3", "numeric natural language search adapter with citations");
                var proposed = JObject.FromObject(ai.Propose(new JObject { ["snapshotId"] = id, ["sheet"] = "合成数据", ["address"] = "E2", ["formula"] = "=SUM(B2:B4)" }));
                Check(sheet.Range["E2"].Value2 == null, "proposal never writes");
                dynamic other = app.Workbooks.Add(-4167);
                ai.Apply((string)proposed["proposalId"]);
                Check((double)sheet.Range["E2"].Value2 == 60 && other.Worksheets[1].Range["E2"].Value2 == null, "apply remains bound to original workbook after switching");
                ai.Apply(null, true); Check(sheet.Range["E2"].Value2 == null, "undo restores blank target");
                proposed = JObject.FromObject(ai.Propose(new JObject { ["snapshotId"] = id, ["sheet"] = "合成数据", ["address"] = "E2", ["formula"] = "=1+1" }));
                sheet.Range["E2"].Value2 = "later edit";
                bool stale = false; try { ai.Apply((string)proposed["proposalId"]); } catch (InvalidOperationException) { stale = true; }
                Check(stale && (string)sheet.Range["E2"].Value2 == "later edit", "stale preview refuses overwrite");
                book.Activate(); sheet.Activate();
                dynamic hidden = book.Worksheets.Add(); hidden.Name = "隐藏"; sheet.Activate(); hidden.Visible = 0;
                var whole = JObject.FromObject(ai.Capture("workbook"));
                Check(whole["ranges"].All(r => (string)r["sheet"] != "隐藏"), "workbook excludes hidden sheets");
                Check(whole["sample"].Any(c => (string)c["address"] == "F2" && c["formula"].Type == JTokenType.Null), "literal equals text is not presented as formula");
                sheet.Range["G2"].Formula = "=1/0";
                var errors = JObject.FromObject(ai.Capture("sheet"));
                Check((int)errors["statistics"]["numericCount"] == 4, "Excel errors excluded from numeric statistics");
                sheet.Range["A1:A200001"].Value2 = 1;
                sheet.Range["A:A"].Select();
                var large=JObject.FromObject(ai.Capture("selection"));
                Check((long)large["nonEmptyCells"]==200001 && (double)large["statistics"]["sum"]==200001,"capture and statistics exceed 200k without truncation");
                sheet.Range["XFD1048576"].Value2=42;
                var entire=JObject.FromObject(ai.Capture("sheet"));
                Check((long)entire["scannedCells"]==17179869184L,"full worksheet bounds supported with 64-bit counts");
                var end=JObject.FromObject(ai.Read(new JObject{["snapshotId"]=(string)entire["snapshotId"],["operation"]="find",["query"]="42",["comparison"]="equals",["column"]=16384}));
                Check((string)end["cells"][0]["address"]=="XFD1048576","sparse full-sheet capture includes final cell");
                AiBatchTests(ai,app,sheet,(string)entire["snapshotId"]);
                book.Close(false);
                bool closed = false; try { ai.Read(new JObject { ["snapshotId"] = id, ["operation"] = "read" }); } catch (InvalidOperationException) { closed = true; }
                Check(closed, "closed workbook invalidates snapshots");
                var recent = new List<string>();
                for (int i = 0; i < 11; i++) recent.Add((string)JObject.FromObject(ai.Capture("sheet"))["snapshotId"]);
                bool expired = false; try { ai.Read(new JObject { ["snapshotId"] = recent[2], ["operation"] = "read" }); } catch (InvalidOperationException) { expired = true; }
                Check(expired, "snapshots expire in insertion order");
                Check(recent.Skip(3).All(key => JObject.FromObject(ai.Read(new JObject { ["snapshotId"] = key, ["operation"] = "read" }))["snapshotId"].ToString() == key), "latest eight snapshots remain usable after repeated eviction");
                other.Close(false);
            }
        }
        finally { if (app != null) { try { app.Quit(); } catch { } Marshal.ReleaseComObject(app); } }
    }
}
