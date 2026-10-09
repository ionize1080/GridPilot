using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using SheetPilot.Core;
using SheetPilot.UI;
using SheetPilot.Excel;

internal static class Program
{
    private static int checks;
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            SearchTests();
            UiTests(args.Length > 1 && args[0] == "--render" ? args[1] : null);
            if (args.Length > 1 && args[0] == "--excel")
            {
                int processId = ExcelTests(Path.GetFullPath(args[1]));
                GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
                try { using (var process = Process.GetProcessById(processId)) Check(process.WaitForExit(15000), "XLL: owned Excel instance exits after Quit"); }
                catch (ArgumentException) { Check(true, "XLL: owned Excel instance already exited"); }
            }
            Console.WriteLine("PASS " + checks + " assertions");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    private static void Check(bool condition, string label)
    { if (!condition) throw new Exception("FAILED: " + label); checks++; Console.WriteLine("PASS " + label); }
    private static SheetItem Item(string name, SheetVisibility visibility = SheetVisibility.Visible) => new SheetItem(name, name, 1, visibility, false, "工作表");
    private static void SearchTests()
    {
        var search = new SheetSearch();
        var sheets = new[] { Item("销售汇总"), Item("Sales.2026"), Item("sales_2025"), Item("隐藏", SheetVisibility.Hidden), Item("配置", SheetVisibility.VeryHidden), Item("趋势图") };
        Func<SearchOptions, SearchResult> run = o => search.Filter(sheets, o, CancellationToken.None);
        Check(run(new SearchOptions()).Items.Count == 6, "empty query includes all sheet types");
        Check(run(new SearchOptions { Query = "销售" }).Items.Single().Name == "销售汇总", "Chinese substring");
        Check(run(new SearchOptions { Query = "SALES" }).Items.Count == 2, "case insensitive");
        Check(run(new SearchOptions { Query = "Sales", MatchCase = true }).Items.Count == 1, "case sensitive");
        Check(run(new SearchOptions { Query = "." }).Items.Count == 1, "literal regex metacharacter");
        Check(run(new SearchOptions { Query = "^sales[._]20(25|26)$", UseRegex = true }).Items.Count == 2, "regex anchors groups and alternation");
        Check(run(new SearchOptions { Query = "[", UseRegex = true }).Error != null, "invalid regex is a user message");
        Check(run(new SearchOptions { IncludeHidden = false }).Items.Count == 4, "hidden and very hidden filter");
        Check(run(new SearchOptions { Query = new string('a', 513) }).Error != null, "pattern length bound");
        Check(run(new SearchOptions { Query = "zzzz" }).Items.Count == 0, "no matches");
        var previousCulture = Thread.CurrentThread.CurrentCulture;
        Thread.CurrentThread.CurrentCulture = new CultureInfo("tr-TR");
        Check(search.Filter(new[] { Item("INVOICE") }, new SearchOptions { Query = "invoice" }, CancellationToken.None).Items.Count == 1, "ordinal casing independent of current culture");
        Thread.CurrentThread.CurrentCulture = previousCulture;
        var clock = Stopwatch.StartNew();
        Check(search.Filter(new[] { Item(new string('a', 30) + "!") }, new SearchOptions { Query = "^(a+)+$", UseRegex = true }, CancellationToken.None).Error != null, "catastrophic regex times out");
        Check(clock.ElapsedMilliseconds < 2000, "timeout is bounded");
        var cancel = new CancellationTokenSource(); cancel.Cancel();
        try { search.Filter(sheets, new SearchOptions(), cancel.Token); Check(false, "cancel"); } catch (OperationCanceledException) { Check(true, "search cancellation"); }
        Check(search.Filter(Enumerable.Range(0, 5000).Select(i => Item("表" + i)).ToArray(), new SearchOptions { Query = "4999" }, CancellationToken.None).Items.Count == 1, "5000 sheet detached search");
        var a = new WorkbookSnapshot("a", "A", sheets);
        Check(a.SameContent(new WorkbookSnapshot("a", "A", sheets)), "unchanged snapshots avoid UI churn");
        Check(!a.SameContent(new WorkbookSnapshot("b", "B", sheets)), "workbook switch changes snapshot");
        Check(!a.SameContent(new WorkbookSnapshot("a", "A", sheets.Reverse())), "sheet reordering changes snapshot");
        using (var registry = new FeatureRegistry())
        {
            var module = new TestModule(); registry.Add(module); registry.Execute("test_go"); Check(module.Executed, "module command dispatch");
            try { registry.Add(new TestModule()); Check(false, "duplicate module"); } catch (ArgumentException) { Check(true, "duplicate module rejected"); }
            registry.Dispose(); Check(module.Disposed, "module cleanup");
        }
    }
    private sealed class TestModule : IFeatureModule
    {
        public string Id => "test"; public string Label => "Test";
        public bool Executed, Disposed;
        public System.Collections.Generic.IEnumerable<FeatureCommand> Commands => new[] { new FeatureCommand("test_go", "Go", "Go", "GoTo", () => Executed = true) };
        public void Dispose() => Disposed = true;
    }
    private static void PumpUntil(Func<bool> done, int milliseconds = 4000)
    {
        var time = Stopwatch.StartNew();
        while (!done())
        {
            if (time.ElapsedMilliseconds > milliseconds) throw new TimeoutException("UI operation did not finish.");
            var frame = new DispatcherFrame();
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() => frame.Continue = false));
            Dispatcher.PushFrame(frame);
            Thread.Sleep(10);
        }
    }
    private static void UiTests(string directory)
    {
        using (var model = new NavigationModel())
        {
            var rows = new[] {
                new SheetItem("1", "项目总览", 1, SheetVisibility.Visible, true, "工作表", "demo", "2026 年度经营分析.xlsx"),
                new SheetItem("2", "销售 · 华东区域", 2, SheetVisibility.Visible, false, "工作表", "demo", "2026 年度经营分析.xlsx"),
                new SheetItem("3", "销售 · 华南区域", 3, SheetVisibility.Visible, false, "工作表", "demo", "2026 年度经营分析.xlsx"),
                new SheetItem("4", "销售 · 华北区域", 4, SheetVisibility.Visible, false, "工作表", "demo", "2026 年度经营分析.xlsx"),
                new SheetItem("5", "月度趋势", 5, SheetVisibility.Visible, false, "图表页", "demo", "2026 年度经营分析.xlsx"),
                new SheetItem("6", "数据字典", 6, SheetVisibility.Hidden, false, "工作表", "demo", "2026 年度经营分析.xlsx"),
                new SheetItem("7", "系统配置", 7, SheetVisibility.VeryHidden, false, "工作表", "demo", "2026 年度经营分析.xlsx") };
            model.Update(new WorkbookSnapshot("demo", "2026 年度经营分析.xlsx", rows));
            PumpUntil(() => model.Results.Count == 7);
            Check(model.Selected.Name == "项目总览", "active sheet selected initially");
            int navigated = 0; model.NavigateRequested += s => navigated++;
            model.Navigate(rows[5]); Check(navigated == 0 && model.HasMessage, "hidden navigation blocked");
            model.Navigate(rows[1]); Check(navigated == 1, "visible navigation requested");
            var view = new NavigationView(model);
            if (directory != null) { Directory.CreateDirectory(directory); Render(view, Path.Combine(directory, "navigation-preview.png"), 340, 820, 96); Render(view, Path.Combine(directory, "navigation-150dpi.png"), 340, 820, 144); Render(view, Path.Combine(directory, "navigation-narrow.png"), 280, 700, 96); }
            model.Query = "销售";
            model.Navigate(rows[0]); Check(navigated == 1, "pending search cannot navigate stale results");
            model.Query = "趋势";
            PumpUntil(() => model.Results.Count == 1 && model.Results[0].Name == "月度趋势");
            Check(model.Results[0].Name == "月度趋势", "latest search wins");
            model.UseRegex = true; model.Query = "[";
            PumpUntil(() => model.HasMessage && model.Results.Count == 0);
            Check(model.IsEmpty, "regex error clears results");
            if (directory != null) Render(view, Path.Combine(directory, "navigation-regex-error.png"), 340, 720, 96);
            model.Query = "^销售";
            PumpUntil(() => model.Results.Count == 3);
            Check(!model.HasMessage, "regex recovery clears error");
            if (directory != null) Render(view, Path.Combine(directory, "navigation-regex.png"), 340, 720, 96);
            bool scope = false; model.ScopeChanged += all => scope = all; model.ScopeIndex = 1;
            Check(scope && model.Results.Count == 0, "scope change clears stale rows");
            model.Update(WorkbookSnapshot.Empty);
            PumpUntil(() => model.Status == "0 / 0 个页面");
            Check(model.IsEmpty, "empty workbook state");
        }
    }
    private static void Render(FrameworkElement element, string path, int width, int height, int dpi)
    {
        element.Measure(new Size(width, height)); element.Arrange(new Rect(0, 0, width, height)); element.UpdateLayout();
        var image = new RenderTargetBitmap(width * dpi / 96, height * dpi / 96, dpi, dpi, PixelFormats.Pbgra32); image.Render(element);
        var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(image));
        using (var stream = File.Create(path)) encoder.Save(stream);
        Check(element.ActualWidth == width && element.ActualHeight == height, "WPF render " + Path.GetFileName(path));
    }
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out int processId);
    private static int ExcelTests(string addIn)
    {
        dynamic app = null;
        int processId = 0;
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            app.Visible = false; app.DisplayAlerts = false;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out processId);
            Console.WriteLine("Owned test Excel process: " + processId);
            Console.WriteLine("Excel version: " + app.Version + "; OS: " + Environment.OSVersion);
            var gateway = new ExcelGateway(() => (object)app);
            dynamic first = app.Workbooks.Add(-4167);
            dynamic overview = first.Sheets.Item[1]; overview.Name = "总览";
            dynamic hidden = first.Worksheets.Add(); hidden.Name = "隐藏数据"; hidden.Visible = 0;
            dynamic secret = first.Worksheets.Add(); secret.Name = "配置"; secret.Visible = 2;
            dynamic chart = first.Charts.Add(); chart.Name = "趋势图";
            overview.Activate();
            var snapshot = gateway.Capture();
            Check(snapshot.Sheets.Count == 4, "COM: includes all Sheets, not only Worksheets");
            Check(snapshot.Sheets.Any(s => s.Kind == "图表页"), "COM: chart sheet recognized");
            Check(snapshot.Sheets.Count(s => !s.CanNavigate) == 2, "COM: hidden and very hidden");
            var overviewItem = snapshot.Sheets.Single(s => s.Name == "总览");
            var chartItem = snapshot.Sheets.Single(s => s.Name == "趋势图");
            gateway.Navigate(chartItem.WorkbookKey, chartItem);
            Check((string)app.ActiveSheet.Name == "趋势图", "COM: chart navigation");
            overview.Name = "新总览";
            gateway.Navigate(overviewItem.WorkbookKey, overviewItem);
            Check((string)app.ActiveSheet.Name == "新总览", "COM: stable identity after sheet rename");
            try { var h = snapshot.Sheets.Single(s => s.Name == "隐藏数据"); gateway.Navigate(h.WorkbookKey, h); Check(false, "hidden"); }
            catch (InvalidOperationException) { Check(true, "COM: hidden visibility rechecked"); }
            dynamic second = app.Workbooks.Add(-4167); second.Sheets.Item[1].Name = "新总览";
            Check(gateway.Capture().Sheets.Count == 1, "COM: current workbook scope");
            Check(gateway.Capture(true).Sheets.Count == 5, "COM: all workbooks scope");
            gateway.Navigate(overviewItem.WorkbookKey, overviewItem);
            Check((string)app.ActiveWorkbook.Name == (string)first.Name, "COM: duplicate names navigate to correct workbook");
            chart.Delete();
            try { gateway.Navigate(chartItem.WorkbookKey, chartItem); Check(false, "deleted"); } catch (InvalidOperationException) { Check(true, "COM: deleted sheet rejected"); }
            first.Close(false);
            try { gateway.Navigate(overviewItem.WorkbookKey, overviewItem); Check(false, "closed"); } catch (InvalidOperationException) { Check(true, "COM: closed workbook rejected"); }
            var offThread = Task.Run(() => { try { gateway.Capture(); return false; } catch (InvalidOperationException) { return true; } });
            Check(offThread.Result, "COM: background access rejected");
            Check((bool)app.RegisterXLL(addIn), "XLL: packed add-in loaded from isolated directory");
            app.Run("SheetPilot_ShowNavigation");
            Check(true, "XLL: navigation command returned");
            // Pump Excel and WPF messages before closing the owned test instance.
            var wait = Stopwatch.StartNew(); PumpUntil(() => wait.ElapsedMilliseconds > 1800);
            second.Close(false);
        }
        finally
        {
            if (app != null) { app.Quit(); Marshal.FinalReleaseComObject((object)app); app = null; }
            GC.Collect(); GC.WaitForPendingFinalizers();
        }
        return processId;
    }
}
