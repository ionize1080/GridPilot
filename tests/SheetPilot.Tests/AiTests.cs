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
    [DllImport("user32.dll", EntryPoint="PostMessage")] private static extern bool AiPost(IntPtr window, int message, IntPtr w, IntPtr l);
    private static void AiHostTest(string addin, int mode)
    {
        dynamic app = null; Process browser = null;
        string previous = Environment.GetEnvironmentVariable("GRIDPILOT_CHATBOX_PROFILE");
        try
        {
            string directory = System.IO.Path.GetDirectoryName(addin);
            Environment.SetEnvironmentVariable("GRIDPILOT_CHATBOX_PROFILE", System.IO.Path.Combine(directory, "test-profile"));
            new ToolPreferencesStore(System.IO.Path.Combine(directory, "GridPilot.tools.json")).Save(new ToolPreferences { AiDisplayMode = mode });
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
                if (mode == 0)
                {
                    AiEnumChildren(excelWindow, (window, unused) =>
                    {
                        var text = new StringBuilder(128); AiClass(window, text, text.Capacity); AiWindowProcess(window, out uint pid);
                        if (text.ToString() == "Chrome_WidgetWin_1" && !before.Contains((int)pid))
                        { chrome = window; browser = Process.GetProcessById((int)pid); return false; }
                        return true;
                    }, IntPtr.Zero);
                }
                else
                {
                    // Electron briefly owns an untitled startup HWND. Wait for the
                    // actual Chatbox window before sending a close message.
                    browser = Process.GetProcessesByName("electron").FirstOrDefault(p => !before.Contains(p.Id) && p.MainWindowHandle != IntPtr.Zero && p.MainWindowTitle == "Chatbox");
                    if (browser != null) chrome = browser.MainWindowHandle;
                }
                return chrome != IntPtr.Zero;
            }, 45000);
            Check(chrome != IntPtr.Zero, mode == 0 ? "native Chatbox is embedded in actual Excel task pane" : "native Chatbox opens as independent window");
            Check(((AiStyle(chrome, -16) & 0x40000000) != 0) == (mode == 0), "window style matches configured mode");
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
    private static void AiTests(bool excel)
    {
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
                bool limited = false; try { ai.Capture("selection"); } catch (InvalidOperationException) { limited = true; }
                Check(limited, "over-limit capture fails instead of silently sampling");
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
