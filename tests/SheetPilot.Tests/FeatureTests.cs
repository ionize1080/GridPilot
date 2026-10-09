using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using SheetPilot.Core;
using SheetPilot.Excel;
using SheetPilot.UI;

internal static partial class Program
{
    private static void FeatureCoreTests()
    {
        var source = new[] { new SheetItem("a", "Sales_2025", 1, SheetVisibility.Visible, true, "工作表", "b", "Book"), new SheetItem("b", "Budget_2025", 2, SheetVisibility.Hidden, false, "工作表", "b", "Book"), new SheetItem("c", "Sales_2025", 1, SheetVisibility.Visible, false, "工作表", "c", "Other") };
        var options = new RenameOptions { Regex = true, Find = "(.*)_2025$", Replacement = "$1_2026" };
        var plan = RenamePlanner.Build(source, options);
        Check(plan.All(p => p.Changed && p.Error == null) && plan[0].NewName == "Sales_2026", "rename: regex capture replacement and cross-book duplicate allowed");
        options.Find = ".*"; options.Replacement = "Same"; options.FirstOnly = true;
        Check(RenamePlanner.Build(source, options).Count(p => p.Error != null) == 2, "rename: collisions checked within each workbook");
        options = new RenameOptions { Find = "Sales", Replacement = "$1" };
        Check(RenamePlanner.Build(source, options)[0].NewName == "$1_2025", "rename: literal replacement does not expand dollar groups");
        options = new RenameOptions { Numbering = true, Start = 7, Step = 2, Digits = 3, IncludeHidden = false, Find = "", Prefix = "Q_" };
        plan = RenamePlanner.Build(source, options);
        Check(plan[0].NewName == "007 Q_Sales_2025" && !plan[1].Selected && plan[2].NewName.StartsWith("007"), "rename: numbering filters and per-book reset");
        options.RestartPerBook = false; Check(RenamePlanner.Build(source, options)[2].NewName.StartsWith("009"), "rename: numbering across workbooks");
        Check(RenamePlanner.ValidateName("bad/name") != null && RenamePlanner.ValidateName("'bad") != null && RenamePlanner.ValidateName(new string('a',32)) != null, "rename: Excel name restrictions");
        var exclude = new HashSet<string> { RenamePlanner.Key(source[0]) };
        Check(!RenamePlanner.Build(source, options, exclude)[0].Selected, "rename: excluded items remain unchanged");
        try { RenamePlanner.Build(source, new RenameOptions { Regex=true, Find="[" }); Check(false,"rename invalid regex"); } catch (ArgumentException) { Check(true,"rename: invalid regex rejected before mutation"); }
        var settings = new ToolPreferences(); double value; string reason;
        Check(!NumericConversion.TryNumber("00123", settings, CultureInfo.InvariantCulture, out value, out reason) && reason.Contains("前导零"), "conversion: leading zero protection");
        Check(!NumericConversion.TryNumber("1234567890123456", settings, CultureInfo.InvariantCulture, out value, out reason), "conversion: precision protection");
        Check(NumericConversion.TryNumber(" 1,234.50 ", settings, CultureInfo.InvariantCulture, out value, out reason) && value == 1234.5, "conversion: whitespace and grouping");
        Check(!NumericConversion.TryNumber("1,23", settings, CultureInfo.InvariantCulture, out value, out reason), "conversion: malformed grouping rejected");
        Check(!NumericConversion.TryNumber("2026-10-09", settings, CultureInfo.InvariantCulture, out value, out reason), "conversion: dates are not implicitly parsed");
        settings.AllowPercent = true;
        Check(NumericConversion.TryNumber("12%", settings, CultureInfo.InvariantCulture, out value, out reason) && value == .12, "conversion: optional percent parser");
        Check(NumericConversion.TryNumber("1.234,5", settings, CultureInfo.GetCultureInfo("de-DE"), out value, out reason) && value == 1234.5, "conversion: decimal culture");
        var path = Path.Combine(Path.GetTempPath(), "GridPilot-options-"+Guid.NewGuid().ToString("N")+".json");
        try { var store = new ToolPreferencesStore(path); settings.FocusColor="#ABCDEF"; settings.Rename.Prefix="Test"; Check(store.Save(settings),"tools: portable settings save"); var loaded = new ToolPreferencesStore(path).Load(); Check(loaded.FocusColor=="#ABCDEF" && loaded.AllowPercent && loaded.Rename.Prefix=="Test", "tools: preferences restored"); File.WriteAllText(path,"{\"SchemaVersion\":999}"); store=new ToolPreferencesStore(path);store.Load();Check(!store.Save(settings),"tools: newer schema not overwritten"); }
        finally { File.Delete(path); if(File.Exists(path+".bak"))File.Delete(path+".bak"); }
    }
    private static void FeatureRenderTests(string directory)
    {
        if (directory == null) return; Directory.CreateDirectory(directory);
        var colors = new NavigationSettings(); var preferences = new ToolPreferences();
        preferences.Rename.Find = "2025"; preferences.Rename.Replacement = "2026";
        var sources = new[] { "经营总览2025", "销售_2025", "预算2025", "说明", "归档2024" }.Select((name, i) => new SheetItem(i.ToString(), name, i+1, SheetVisibility.Visible, i==0, "工作表", "book", "年度分析.xlsx")).ToArray();
        var rename = new RenameWindow(preferences, colors, all => new WorkbookSnapshot("book", "年度分析.xlsx", sources), p => "测试预览", () => "已保存");
        rename.RefreshSources(); RenderWindowContent(rename, Path.Combine(directory,"rename-preview.png"), 1060, 720, 96); rename.Close();
        var settings = new PreferencesWindow(preferences, colors, () => "已保存", () => {});
        RenderWindowContent(settings, Path.Combine(directory,"tools-settings.png"), 620, 740, 96); settings.Close();
        var paletteWindow=new Window(); ToolUiTheme(paletteWindow,colors);
        paletteWindow.Content=new ColorPalette {Value="#2563EB",Margin=new Thickness(16)};
        RenderWindowContent(paletteWindow,Path.Combine(directory,"color-palette.png"),292,354,96);paletteWindow.Close();
        using(var montage = new System.Drawing.Bitmap(8*80,100)) using(var graphics=System.Drawing.Graphics.FromImage(montage))
        {
            graphics.Clear(System.Drawing.Color.FromArgb(247,249,247)); int index=0;
            foreach(var id in new[]{"SheetPilot_Show","SheetPilot_Refresh","GridPilot_Rename","GridPilot_Focus","GridPilot_Text","GridPilot_Number","GridPilot_Undo","GridPilot_Settings"})
            { using(var icon=RibbonIcons.Draw(id,32)) { icon.Save(Path.Combine(directory,id+"-32.png"),System.Drawing.Imaging.ImageFormat.Png);graphics.DrawImageUnscaled(icon,index*80+24,10); } using(var icon=RibbonIcons.Draw(id,48))graphics.DrawImageUnscaled(icon,index*80+16,48); index++; }
            montage.Save(Path.Combine(directory,"ribbon-icons.png"),System.Drawing.Imaging.ImageFormat.Png);
        }
        Check(true,"icons: 32px and 48px custom assets rendered");
    }
    private static void RenderWindowContent(System.Windows.Window window, string path, int width, int height, int dpi)
    {
        var content = (System.Windows.UIElement)window.Content; window.Content = null;
        var host = new System.Windows.Controls.Border { Child = content, Resources = window.Resources, Background = (System.Windows.Media.Brush)window.FindResource("CanvasBrush") };
        System.Windows.Documents.TextElement.SetForeground(host, (System.Windows.Media.Brush)window.FindResource("TextBrush"));
        System.Windows.Documents.TextElement.SetFontFamily(host, window.FontFamily); System.Windows.Documents.TextElement.SetFontSize(host, window.FontSize);
        Render(host, path, width, height, dpi); host.Child = null; window.Content = content;
    }
    private static void FeatureExcelTests(object application, string directory)
    {
        dynamic app = application; dynamic book = null;
        var owned = new List<object>(); bool visible = app.Visible;
        try
        {
            book = app.Workbooks.Add(-4167); owned.Add((object)book);
            dynamic alpha = book.Sheets.Item[1]; alpha.Name="Alpha"; owned.Add((object)alpha);
            dynamic beta = book.Worksheets.Add(); beta.Name="Beta"; owned.Add((object)beta);
            alpha.Activate();
            var gateway = new ExcelGateway(() => application); var snapshot = gateway.Capture();
            var swap = snapshot.Sheets.Select(s => new RenameEntry(s,true,s.Name=="Alpha"?"Beta":"Alpha")).ToArray();
            gateway.ApplyRename(swap);
            Check((string)alpha.Name=="Beta" && (string)beta.Name=="Alpha", "rename COM: cyclic names applied through temporary names");
            gateway.UndoLast(); Check((string)alpha.Name=="Alpha" && (string)beta.Name=="Beta", "rename COM: undo restores both names");
            snapshot=gateway.Capture(); var plan=RenamePlanner.Build(snapshot.Sheets,new RenameOptions { Prefix="FY_", MatchesOnly=false });
            book.Protect("test",true,false);
            try { gateway.ApplyRename(plan); Check(false,"protected rename"); } catch(InvalidOperationException){Check((string)alpha.Name=="Alpha","rename COM: structure protection blocks all writes");} finally {book.Unprotect("test");}
            beta.Name="Changed";
            try { gateway.ApplyRename(plan); Check(false,"stale rename"); } catch(InvalidOperationException){Check((string)alpha.Name=="Alpha","rename COM: stale preview blocks all writes");}
            dynamic first=alpha.Range["A1"]; first.Value2=1234.5; first.NumberFormat="#,##0.00"; first.HorizontalAlignment=-4108; first.VerticalAlignment=-4108; first.IndentLevel=2; first.WrapText=true; first.Font.Bold=true; first.Interior.Color=0xAAFFEE;
            alpha.Range["A2"].Formula="=1+1"; alpha.Range["A3"].Formula="'00123"; alpha.Range["A4"].Formula="'123.50"; alpha.Range["A4"].NumberFormat="@"; alpha.Range["A5"].Formula="'1234567890123456789"; alpha.Range["A6"].Formula="=NA()"; alpha.Range["A7"].Value2=444; alpha.Range["A8"].Value2=888;
            string format=FormatSnapshot(first); var settings=new ToolPreferences { DecimalCulture=1 };
            app.Union(alpha.Range["A1"],alpha.Range["A8"]).Select(); gateway.ConvertSelection(true,settings);
            Check((object)first.Value2 is string && (string)first.Value2=="1234.5" && (object)alpha.Range["A8"].Value2 is string,"conversion COM: discontiguous numeric constants become real text");
            Check((double)alpha.Range["A7"].Value2==444 && (string)alpha.Range["A2"].Formula=="=1+1", "conversion COM: cells between areas and formulas untouched");
            Check(FormatSnapshot(first)==format,"conversion COM: number format alignment indent font and fill unchanged");
            Check((int)alpha.Range["A8"].HorizontalAlignment==-4152,"conversion COM: General numeric alignment is fixed right for text");
            Check((string)first.PrefixCharacter=="'", "conversion COM: apostrophe text supports Excel numeric-text indicator");
            gateway.UndoLast(); Check((object)first.Value2 is double && (object)alpha.Range["A8"].Value2 is double && FormatSnapshot(first)==format,"conversion COM: undo restores values and retains formatting");
            Check((int)alpha.Range["A8"].HorizontalAlignment==1,"conversion COM: undo restores General alignment");
            alpha.Range["A3:A5"].Select(); gateway.ConvertSelection(false,settings);
            Check((object)alpha.Range["A3"].Value2 is string && (object)alpha.Range["A5"].Value2 is string,"conversion COM: leading zeros and long IDs preserved");
            Check((object)alpha.Range["A4"].Value2 is double && (double)alpha.Range["A4"].Value2==123.5 && (string)alpha.Range["A4"].NumberFormat=="@","conversion COM: number stored without changing text number format");
            Check((int)alpha.Range["A4"].HorizontalAlignment==-4131,"conversion COM: General text alignment is fixed left for number");
            gateway.UndoLast(); Check((object)alpha.Range["A4"].Value2 is string,"conversion COM: undo numeric conversion");
            alpha.Range["A4"].Select(); gateway.ConvertSelection(false,settings); Check((object)alpha.Range["A3"].Value2 is string,"conversion COM: one-cell selection cannot expand into neighbors");
            alpha.Range["A4"].Value2=99;
            try{gateway.UndoLast();Check(false,"stale conversion undo");}catch(InvalidOperationException){Check((double)alpha.Range["A4"].Value2==99,"conversion COM: undo refuses changed content");}
            var multiPlan=RenamePlanner.Build(gateway.Capture(true).Sheets,new RenameOptions { Regex=true, Find="^(Alpha|Changed|新总览)$", Replacement="$1_2026" });
            Check(multiPlan.Count(p=>p.Changed)==3,"rename COM: all-workbook preview includes both test workbooks");
            gateway.ApplyRename(multiPlan); Check((string)alpha.Name=="Alpha_2026" && (string)beta.Name=="Changed_2026","rename COM: cross-workbook apply"); gateway.UndoLast();
            app.Visible=true; app.WindowState=-4137; SetForegroundWindow(new IntPtr((int)book.Windows.Item[1].Hwnd)); app.Goto(alpha.Range["C3"], true);
            Console.WriteLine("Focus target: " + book.Name + " actual: " + app.ActiveWorkbook.Name + " / " + app.ActiveCell.Address);
            Check((string)app.ActiveWorkbook.Name==(string)book.Name && (string)app.ActiveCell.Address=="$C$3", "focus COM: active workbook and selected center verified");
            alpha.Range["B1"].Value2="项目"; alpha.Range["C1"].Value2="本期"; alpha.Range["D1"].Value2="预算";
            for(int row=2;row<=10;row++){alpha.Cells[row,2].Value2="项目 "+(row-1);alpha.Cells[row,3].Value2=row*1280;alpha.Cells[row,4].Value2=row*1400;}
            alpha.Range["B1:D10"].ColumnWidth=16; alpha.Range["B1:D10"].RowHeight=22; alpha.Range["B1:D1"].Font.Bold=true; alpha.Range["B1:D1"].Interior.Color=0xE0F0E5;
            PumpUntil(()=>{var frame=FocusOverlay.Capture(application,settings);return frame.Bands.Length>0;});
            string before=FormatSnapshot(alpha.Range["C3"]); bool saved=book.Saved; int conditions=alpha.Cells.FormatConditions.Count;
            var geometry=FocusOverlay.Capture(application,settings);
            Check(geometry.Bands.Length>=2 && geometry.Bands.All(b=>geometry.Viewport.Contains(b)) && geometry.Bands.All(b => app.ActiveWindow.RangeFromPoint(b.Right-1,b.Bottom-1) != null),"focus COM: row and column bands clipped to visible grid");
            using(var focus = new FocusOverlay(()=>application,action=>action(),settings))
            {
                SetForegroundWindow(FocusOverlay.ActiveWindowHandle(application)); focus.Toggle(); PumpUntil(()=>focus.IsVisible,3000);
                Check(focus.IsVisible, "focus COM: owned transparent overlay actually visible");
                int updates=focus.UpdateCount; alpha.Range["D4"].Select();
                Check(focus.UpdateCount>updates,"focus COM: selection event renders synchronously without timer or queued macro");
                alpha.Range["C3"].Select();
                Console.WriteLine(focus.PerformanceSummary);
                using (var screenshot = new System.Drawing.Bitmap(geometry.Viewport.Width, geometry.Viewport.Height)) using (var graphics = System.Drawing.Graphics.FromImage(screenshot)) { graphics.CopyFromScreen(geometry.Viewport.Location, System.Drawing.Point.Empty, geometry.Viewport.Size); screenshot.Save(Path.Combine(directory,"focus-overlay.png"),System.Drawing.Imaging.ImageFormat.Png); }
                Check(focus.Enabled && FormatSnapshot(alpha.Range["C3"])==before && (bool)book.Saved==saved && (int)alpha.Cells.FormatConditions.Count==conditions,"focus COM: overlay leaves formatting rules and dirty state unchanged");
                focus.Toggle(); Check(!focus.Enabled,"focus COM: toggle removes overlay");
            }
            dynamic window=app.ActiveWindow; window.SplitRow=1; window.SplitColumn=1; window.FreezePanes=true;
            geometry=FocusOverlay.Capture(application,settings); Check(geometry.Bands.Length>0 && geometry.Bands.All(b=>geometry.Viewport.Contains(b)),"focus COM: frozen panes generate bounded overlays");
            window.FreezePanes=false; window.SplitRow=0; window.SplitColumn=0;
        }
        finally { app.Visible=visible; if(book!=null)book.Close(false); foreach(var item in owned.AsEnumerable().Reverse()) if(Marshal.IsComObject(item))Marshal.FinalReleaseComObject(item); }
    }
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
    private static void ToolUiTheme(Window window,NavigationSettings colors)
    {window.Resources.MergedDictionaries.Add(new ResourceDictionary {Source=new Uri("/SheetPilot.UI;component/Styles.xaml",UriKind.Relative)});using(var model=new NavigationModel(colors)) using(var view=new NavigationView(model)) window.Resources=view.Resources;window.FontFamily=new System.Windows.Media.FontFamily("Microsoft YaHei UI");window.FontSize=12;}
    private static string FormatSnapshot(dynamic cell) => string.Join("|",new object[]{cell.NumberFormat,cell.HorizontalAlignment,cell.VerticalAlignment,cell.IndentLevel,cell.WrapText,cell.Font.Bold,cell.Interior.Color});
}
