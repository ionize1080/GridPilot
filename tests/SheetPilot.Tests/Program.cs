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

internal static partial class Program
{
    private static int checks;
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            if(args.Length==0 && Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetExecutingAssembly().Location).Contains("Compatibility")) return CompatibilityRunner();
            if(args.Length>0&&args[0]=="--report-json"){Console.OutputEncoding=new System.Text.UTF8Encoding(false);Console.WriteLine(Newtonsoft.Json.JsonConvert.SerializeObject(ReportReader.Read(args[1],args.Length>2?args[2]:null,0,CancellationToken.None)));return 0;}
            Console.WriteLine("Test process: " + (Environment.Is64BitProcess ? "x64" : "x86"));
            if(args.Length>1&&args[0]=="--simple-launch"){V11Launch(Path.GetFullPath(args[1]),"GridPilot_SimpleMask");Console.WriteLine("PASS "+checks+" assertions");return 0;}
            if(args.Length>1&&args[0]=="--v11-launch"){V11Launch(Path.GetFullPath(args[1]));Console.WriteLine("PASS "+checks+" assertions");return 0;}
            if(args.Length>0 && args[0]=="--simple-mask"){SimpleMaskCoreTests();if(args.Contains("--excel"))SimpleMaskExcelTests();Console.WriteLine("PASS "+checks+" assertions");return 0;}
            if(args.Length>0 && args[0]=="--v11"){V11Tests(args.Contains("--excel"));Console.WriteLine("PASS "+checks+" assertions");return 0;}
            if(args.Length>0 && args[0]=="--report-schema"){var data=Newtonsoft.Json.Linq.JObject.FromObject(ReportReader.Read(args[1],"Page1_1",0,CancellationToken.None));Console.WriteLine("Report rows="+data["rows"].Count()+" sheets="+data["sheets"].Count());return 0;}
            if(args.Length>0 && args[0]=="--v1"){V1CoreTests();if(args.Contains("--excel"))V1ExcelTests(args.Contains("--large"));Console.WriteLine("PASS "+checks+" assertions");return 0;}
            if(args.Length>2 && args[0]=="--release-ui") {ReleaseUiSmoke(Path.GetFullPath(args[1]),Path.GetFullPath(args[2]));return 0;}
            if(args.Length>0 && args[0]=="--release-features") { ReleaseFeatureTests(); if(args.Length>1 && args[1]=="--excel") ReleaseFeatureExcelTests(); Console.WriteLine("PASS "+checks+" assertions");return 0; }
            if(args.Length>2 && args[0]=="--interactive") { InteractiveTests(Path.GetFullPath(args[1]),Path.GetFullPath(args[2])); return 0; }
            if (args.Length > 0 && args[0] == "--color-picker-host") { ColorPickerHostTests(); return 0; }
            if (args.Length > 0 && args[0] == "--copy-excel") { int pid = CopyExcelTests(args.Length > 1 ? Path.GetFullPath(args[1]) : null); CopyExcelExitCheck(pid); return 0; }
            if(args.Length>0 && args[0]=="--upgrade") { UpgradeCoreTests(); if(args.Length>2)UpgradeExcelTests(args[1],args[2]);Console.WriteLine("PASS "+checks+" assertions");return 0;}
            if(args.Length>0 && args[0]=="--upgrade-excel") {UpgradeCoreTests();string temporary=Path.Combine(Path.GetTempPath(),"GridPilot-upgrade-"+Guid.NewGuid().ToString("N"));string samples=Path.Combine(temporary,"samples");Directory.CreateDirectory(samples);UpgradeExcelTests(samples,Path.Combine(temporary,"output"));Console.WriteLine("PASS "+checks+" assertions");return 0;}
            SimpleMaskCoreTests();
            SearchTests();
            SettingsTests();
            PreferenceTests();
            FeatureCoreTests();
            CopyCoreTests();
            RegressionTests();
            FeatureRenderTests(args.Length > 1 && args[0] == "--render" ? args[1] : null);
            UiTests(args.Length > 1 && args[0] == "--render" ? args[1] : null);
            if (args.Length > 1 && args[0] == "--excel")
            {
                int processId = ExcelTests(Path.GetFullPath(args[1]));
                // COM finalizers may marshal to this STA; keep it pumping instead of blocking it.
                for (int i = 0; i < 2; i++) { GC.Collect(); var finalizers = Task.Run(() => GC.WaitForPendingFinalizers()); PumpUntil(() => finalizers.IsCompleted, 10000); }
                try { using (var process = Process.GetProcessById(processId)) { PumpUntil(() => process.HasExited, 15000); Check(true, "XLL: owned Excel instance exits after Quit"); } }
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
    private static void SettingsTests()
    {
        var directory = Path.Combine(Path.GetTempPath(), "SheetPilot-settings-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "SheetPilot.settings.json");
            var store = new SettingsStore(path);
            var settings = store.Load();
            Check(settings.DockSide == 0 && settings.WorkbookNameMode == 0 && !settings.ShowHeader, "settings: compact right dock defaults");
            settings.DockSide = 1; settings.RowHeight = 52; settings.RowGap = 4; settings.FontSize = 16;
            settings.WorkbookNameMode = 2; settings.PaneWidth = 380; settings.ShowHeader = true;
            settings.ShowSheetNumbers = false; settings.ShowFooter = false; settings.ShowKeyboardHints = true;
            settings.AutoRefresh = false; settings.RefreshSeconds = 7; settings.ScopeIndex = 1; settings.UseRegex = true; settings.IncludeHidden = false;
            Check(store.Save(settings), "settings: portable JSON save");
            var loaded = new SettingsStore(path).Load();
            Check(loaded.DockSide == 1 && loaded.RowHeight == 52 && loaded.RowGap == 4 && loaded.FontSize == 16 && loaded.PaneWidth == 380, "settings: layout restored after new store");
            Check(loaded.WorkbookNameMode == 2 && loaded.ScopeIndex == 1 && loaded.UseRegex && !loaded.IncludeHidden, "settings: scope and search preferences restored");
            Check(loaded.ShowHeader && !loaded.ShowFooter && !loaded.ShowSheetNumbers && loaded.ShowKeyboardHints && !loaded.AutoRefresh && loaded.RefreshSeconds == 7, "settings: display and refresh preferences restored");
            loaded.RowHeight = 999; loaded.RowGap = -1; loaded.PaneWidth = 1; loaded.FontSize = 100;
            Check(loaded.RowHeight == 80 && loaded.RowGap == 0 && loaded.PaneWidth == 260 && loaded.FontSize == 18, "settings: unsafe numeric values clamped");
            Check(store.Save(loaded) && File.Exists(path + ".bak"), "settings: atomic replacement keeps backup");
            File.SetAttributes(path, FileAttributes.ReadOnly);
            Check(!store.Save(settings) && store.HasError, "settings: read-only config does not report a successful save");
            File.SetAttributes(path, FileAttributes.Normal);
            File.WriteAllText(path, "{broken");
            var corrupt = new SettingsStore(path); Check(corrupt.Load().DockSide == 0 && corrupt.HasError, "settings: corrupt file falls back without startup failure");
            Check(corrupt.Save(new NavigationSettings()) && File.ReadAllText(path + ".bak") == "{broken", "settings: corrupt original backed up on explicit change");
            File.WriteAllText(path, "{\"SchemaVersion\":999}");
            var future = new SettingsStore(path); future.Load();
            Check(!future.Save(new NavigationSettings()) && File.ReadAllText(path).Contains("999"), "settings: newer schema is never overwritten");
            File.WriteAllText(path, "{\"SchemaVersion\":1,\"DockSide\":1}");
            var partial = new SettingsStore(path).Load();
            Check(partial.DockSide == 1 && partial.RowHeight == 36 && partial.IncludeHidden, "settings: missing fields get defaults");
            var blocker = Path.Combine(directory, "not-a-folder"); File.WriteAllText(blocker, "");
            var unavailable = new SettingsStore(Path.Combine(blocker, "settings.json"));
            Check(!unavailable.Save(settings) && unavailable.HasError, "settings: unwritable destination reports failure");
            settings.Reset();
            Check(settings.DockSide == 0 && settings.RowHeight == 36 && settings.WorkbookNameMode == 0 && !settings.ShowHeader, "settings: restore defaults");
            using (var restoredModel = new NavigationModel(partial))
                Check(restoredModel.Settings.DockSide == 1 && !restoredModel.ShowWorkbookNames, "settings: new UI uses persisted preferences");
        }
        finally { Directory.Delete(directory, true); }
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
    private static void PreferenceTests()
    {
        Check(!PreferenceInput.TryInteger("", 28, 80, out _) && !PreferenceInput.TryInteger("3.5", 28, 80, out _), "input: incomplete and fractional numbers rejected");
        Check(PreferenceInput.TryInteger("999", 28, 80, out int clamped) && clamped == 80, "input: typed number constrained to supported range");
        Check(PreferenceInput.TryColor(" #abc ", out string color) && color == "#AABBCC", "input: shorthand hex normalized");
        Check(!PreferenceInput.TryColor("#GG1111", out _) && !PreferenceInput.TryColor("", out _), "input: invalid colors rejected");
        var settings = new NavigationSettings();
        var guard = new NativeLayoutGuard(); int writes = 0;
        settings.PropertyChanged += (s,e) => { if (e.PropertyName == nameof(settings.PaneWidth) && !guard.IsObserving) writes++; };
        guard.Observe(() => { settings.PaneWidth = 490; guard.Observe(() => settings.PaneWidth = 500); });
        Check(settings.PaneWidth == 500 && writes == 0, "resize: native observations persist without geometry feedback");
        settings.PaneWidth = 510; Check(writes == 1, "resize: explicit setting still requests geometry change");
        try { guard.Observe(() => throw new InvalidOperationException()); } catch (InvalidOperationException) { }
        Check(!guard.IsObserving, "resize: origin guard resets after failure");
        using (var model = new NavigationModel(settings))
        {
            int layoutNotifications = 0;
            model.PropertyChanged += (s,e) => { if (e.PropertyName == "EffectiveRowHeight" || e.PropertyName == "RowMargin" || e.PropertyName == "ShowWorkbookNames") layoutNotifications++; };
            settings.PaneWidth = 520;
            Check(layoutNotifications == 0, "resize: width does not invalidate result row layout");
            settings.RowHeight = 44; Check(layoutNotifications == 1, "resize: genuine row height change still invalidates layout");
            using (var view = new NavigationView(model))
            {
                var number = (NumberEditor)view.FindName("RowHeightEditor");
                // These editors are live bound controls, not parsing-only mocks.
                view.Measure(new Size(340, 900)); view.Arrange(new Rect(0, 0, 340, 900)); view.UpdateLayout();
                var box = ((System.Windows.Controls.DockPanel)number.Children[0]).Children.OfType<System.Windows.Controls.TextBox>().Single();
                box.Text = ""; Check(settings.RowHeight == 44, "input: clearing draft does not mutate setting");
                box.Text = "57"; Check(number.CommitInput() && settings.RowHeight == 57, "input: commit reaches two-way setting binding");
                box.Text = "x"; Check(!number.CommitInput() && settings.RowHeight == 57, "input: invalid draft preserves previous setting");
                ((System.Windows.Controls.Slider)number.Children[1]).Value = 60;
                Check(settings.RowHeight == 60 && box.Text == "60", "input: slider and typed value stay synchronized");
                var editor = new ColorEditor(); editor.SetBinding(ColorEditor.ValueProperty, new System.Windows.Data.Binding("AccentColor") { Source = settings, Mode = System.Windows.Data.BindingMode.TwoWay });
                var colorBox = ((System.Windows.Controls.DockPanel)editor.Children[0]).Children.OfType<System.Windows.Controls.TextBox>().Single();
                colorBox.Text = "a3c"; Check(editor.CommitInput() && settings.AccentColor == "#AA33CC", "color: editor normalizes and saves custom color");
                colorBox.Text = "bad-value"; Check(!editor.CommitInput() && settings.AccentColor == "#AA33CC", "color: invalid draft preserves palette");
                settings.ApplyPalette("dark");
                Check(((SolidColorBrush)view.Resources["CanvasBrush"]).Color == (Color)ColorConverter.ConvertFromString("#20262C"), "color: live palette updates view resources");
                settings.ShowTooltips = false;
                Check(!System.Windows.Controls.ToolTipService.GetIsEnabled((DependencyObject)view.FindName("ClearButton")), "tooltip: preference disables button hints");
                settings.ShowTooltips = true;
                Check(System.Windows.Controls.ToolTipService.GetIsEnabled((DependencyObject)view.FindName("ClearButton")), "tooltip: preference restores button hints");
            }
        }
        var path = Path.Combine(Path.GetTempPath(), "SheetPilot-colors-" + Guid.NewGuid().ToString("N") + ".json");
        try
        {
            settings.ShowTooltips = false; settings.BackgroundColor = "#123456"; settings.TextColor = "#ABCDEF"; settings.PaneWidth = 1200;
            Check(new SettingsStore(path).Save(settings), "color: save custom palette");
            var restored = new SettingsStore(path).Load();
            Check(!restored.ShowTooltips && restored.BackgroundColor == "#123456" && restored.TextColor == "#ABCDEF" && restored.PaneWidth == 1200, "color: hints and palette survive restart");
            File.WriteAllText(path, "{\"SchemaVersion\":1,\"BackgroundColor\":\"invalid\"}");
            restored = new SettingsStore(path).Load();
            Check(restored.BackgroundColor == "#F7F9F7" && restored.ShowTooltips, "color: old or malformed fields receive safe defaults");
        }
        finally { File.Delete(path); }
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
            Check(!model.ShowWorkbookNames && model.EffectiveRowHeight == 36, "UI: current scope omits redundant workbook subtitle");
            model.Settings.WorkbookNameMode = 1;
            Check(model.ShowWorkbookNames && model.EffectiveRowHeight >= 40, "UI: always-show reserves two-line height");
            model.Settings.FontSize = 18; model.Settings.RowHeight = 28;
            Check(model.EffectiveRowHeight >= 50, "UI: large font and small row do not overlap");
            model.Settings.Reset();
            if (directory != null) { Directory.CreateDirectory(directory); Render(view, Path.Combine(directory, "navigation-preview.png"), 340, 820, 96); Render(view, Path.Combine(directory, "navigation-150dpi.png"), 340, 820, 144); Render(view, Path.Combine(directory, "navigation-narrow.png"), 280, 700, 96); }
            if (directory != null)
            {
                var list = (System.Windows.Controls.ListBox)view.FindName("SheetList");
                Check(((System.Windows.Controls.ListBoxItem)list.ItemContainerGenerator.ContainerFromIndex(0)).ActualHeight == model.EffectiveRowHeight, "UI: configured row height reaches actual WPF item");
                var many = rows.Concat(Enumerable.Range(8, 16).Select(i => new SheetItem(i.ToString(), new[] { "费用明细", "预算执行", "应收账款", "现金流", "库存清单", "采购计划", "客户分析", "资产台账" }[(i-8)%8] + " " + i, i, SheetVisibility.Visible, false, "工作表", "demo", "2026 年度经营分析.xlsx"))).ToArray();
                model.Update(new WorkbookSnapshot("demo", "2026 年度经营分析.xlsx", many)); PumpUntil(() => model.Results.Count == 23);
                Render(view, Path.Combine(directory, "navigation-preview.png"), 340, 820, 96);
                model.Settings.RowHeight = 28; model.Settings.RowGap = 0; model.Settings.ShowSheetNumbers = false; model.Settings.ShowFooter = false;
                Render(view, Path.Combine(directory, "navigation-compact.png"), 280, 700, 96);
                model.Settings.Reset();
                model.IsSettingsOpen = true;
                Render(view, Path.Combine(directory, "navigation-settings.png"), 340, 1000, 96);
                model.Settings.ApplyPalette("dark");
                Render(view, Path.Combine(directory, "navigation-settings-dark.png"), 340, 1100, 96);
                var settingsScroll = (System.Windows.Controls.ScrollViewer)view.FindName("SettingsScroll");
                settingsScroll.ScrollToEnd(); view.UpdateLayout();
                Render(view, Path.Combine(directory, "navigation-colors.png"), 340, 900, 96);
                model.IsSettingsOpen = false;
                Render(view, Path.Combine(directory, "navigation-dark.png"), 340, 820, 96);
                model.Settings.ApplyPalette("neutral");
                Render(view, Path.Combine(directory, "navigation-neutral.png"), 340, 820, 96);
                model.Settings.Reset();
                model.Update(new WorkbookSnapshot("demo", "2026 年度经营分析.xlsx", rows)); PumpUntil(() => model.Results.Count == 7);
            }
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
            Check(model.ShowWorkbookNames, "UI: all-workbook scope automatically shows workbook names");
            model.Update(new WorkbookSnapshot("demo", "全部 1 个工作簿", rows)); PumpUntil(() => model.Results.Count == 3);
            if (directory != null) Render(view, Path.Combine(directory, "navigation-all-workbooks.png"), 340, 720, 96);
            model.Settings.WorkbookNameMode = 2;
            Check(!model.ShowWorkbookNames, "UI: never-show preference overrides all-workbook scope");
            model.Settings.WorkbookNameMode = 0;
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
        var owned = new System.Collections.Generic.List<object>();
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
            owned.Add((object)first);
            dynamic overview = first.Sheets.Item[1]; overview.Name = "总览";
            owned.Add((object)overview);
            dynamic hidden = first.Worksheets.Add(); hidden.Name = "隐藏数据"; hidden.Visible = 0;
            owned.Add((object)hidden);
            dynamic secret = first.Worksheets.Add(); secret.Name = "配置"; secret.Visible = 2;
            owned.Add((object)secret);
            dynamic chart = first.Charts.Add(); chart.Name = "趋势图";
            owned.Add((object)chart);
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
            owned.Add((object)second);
            Check(gateway.Capture().Sheets.Count == 1, "COM: current workbook scope");
            var allBooks = gateway.Capture(true);
            Check(allBooks.Sheets.Count == 5, "COM: all workbooks scope");
            // Excel can initialize a formerly empty CodeName when creating another workbook.
            // Cross-book navigation uses the refreshed all-workbooks directory, as the UI does.
            overviewItem = allBooks.Sheets.Single(s => s.WorkbookKey == overviewItem.WorkbookKey && s.Name == "新总览");
            gateway.Navigate(overviewItem.WorkbookKey, overviewItem);
            Check((string)app.ActiveWorkbook.Name == (string)first.Name, "COM: duplicate names navigate to correct workbook");
            chart.Delete();
            try { gateway.Navigate(chartItem.WorkbookKey, chartItem); Check(false, "deleted"); } catch (InvalidOperationException) { Check(true, "COM: deleted sheet rejected"); }
            first.Close(false);
            try { gateway.Navigate(overviewItem.WorkbookKey, overviewItem); Check(false, "closed"); } catch (InvalidOperationException) { Check(true, "COM: closed workbook rejected"); }
            var offThread = Task.Run(() => { try { gateway.Capture(); return false; } catch (InvalidOperationException) { return true; } });
            Check(offThread.Result, "COM: background access rejected");
            FeatureExcelTests((object)app, Path.GetDirectoryName(addIn));
            second.Activate();
            Check((bool)app.RegisterXLL(addIn), "XLL: packed add-in loaded from isolated directory");
            app.Run("SheetPilot_ShowNavigation");
            Check(true, "XLL: navigation command returned");
            // Pump Excel and WPF messages before closing the owned test instance.
            var wait = Stopwatch.StartNew(); PumpUntil(() => wait.ElapsedMilliseconds > 1800);
            var expectedSettings = new SettingsStore(Path.Combine(Path.GetDirectoryName(addIn), "GridPilot.navigation.json")).Load();
            string paneLog = File.ReadAllText(Path.Combine(Path.GetTempPath(), "GridPilot-" + processId + ".log"));
            Check(paneLog.Contains("dock=" + (expectedSettings.DockSide == 1 ? "msoCTPDockPositionLeft" : "msoCTPDockPositionRight")), "XLL: actual task pane uses saved/default dock side");
            Check(paneLog.Contains("width=" + expectedSettings.PaneWidth), "XLL: actual task pane uses saved/default width");
            dynamic commandCell = second.Sheets.Item[1].Range["A1"]; commandCell.Value2 = 246.5; commandCell.Select();
            app.Run("GridPilot_Text"); PumpUntil(() => (object)commandCell.Value2 is string);
            Check((string)commandCell.Value2 == "246.5", "XLL: registered conversion command executes queued operation");
            app.Run("GridPilot_Undo"); PumpUntil(() => (object)commandCell.Value2 is double);
            Check((double)commandCell.Value2 == 246.5, "XLL: registered undo command restores content");
            second.Close(false);
        }
        finally
        {
            // Only this isolated test owns these RCWs. Never do this in the production gateway.
            foreach (var item in owned.AsEnumerable().Reverse()) if (Marshal.IsComObject(item)) Marshal.FinalReleaseComObject(item);
            owned.Clear();
            if (app != null) { app.Quit(); Marshal.FinalReleaseComObject((object)app); app = null; }
        }
        return processId;
    }
}
