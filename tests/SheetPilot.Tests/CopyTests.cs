using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using SheetPilot.Core;
using SheetPilot.Excel;
using SheetPilot.UI;

internal static partial class Program
{
    private static void CopyCoreTests()
    {
        var areas = new[] { new CellRect(1, 1, 10, 4), new CellRect(1, 6, 10, 1) };
        foreach (string formula in new[] { "=A1+$B$2", "=SUM(A1:D10)", "=IF(A1>0,\"!outside[book]A99\",B2)", "='O''Brien'!A1*2", "=LOG10(100)+1E-3", "=SUM(A1:A10,F1:F10)", "=ROW()+COLUMN()", "=INDEX(A1:D10,2,1)", "=TRUE", "=SUM({1,2;3,4})", "=IF(TRUE,\"a\"\"b\",A1)" })
            Check(InternalFormula.CanKeep(formula, "O'Brien", areas, out _), "copy formula accepts " + formula);
        foreach (string formula in new[] { "=A11", "=SUM(A1:F10)", "=Other!A1", "='[Book]O''Brien'!A1", "=SUM(MyName)", "=INDIRECT(\"A1\")", "=OFFSET(A1,0,0)", "=SUM(Table1[Value])", "=A1#", "=@A1", "=NOW()", "=RAND()", "=CELL(\"filename\")", "=MyUdf(A1)", "=XFE1", "=A1048577", "=A:A", "=1:1", "=Sheet1:Sheet3!A1", "=SUM(A1:Unknown)", "=HYPERLINK(\"https://example.com\")" })
            Check(!InternalFormula.CanKeep(formula, "O'Brien", areas, out string reason) && !string.IsNullOrEmpty(reason), "copy formula rejects " + formula);
        Check(CellRect.Covered(new CellRect(1, 1, 10, 2), new[] { new CellRect(1, 1, 5, 2), new CellRect(6, 1, 5, 2) }), "copy coverage combines adjacent areas");
        Check(!CellRect.Covered(new CellRect(1, 1, 10, 2), new[] { new CellRect(1, 1, 4, 2), new CellRect(6, 1, 5, 2) }), "copy coverage detects missing row");
        var file = Path.Combine(Path.GetTempPath(), "GridPilot-copy-settings-" + Guid.NewGuid().ToString("N") + ".json");
        try
        {
            File.WriteAllText(file, "{\"SchemaVersion\":1}"); var store = new ToolPreferencesStore(file); var options = store.Load();
            Check(options.Copy.CellFormats && options.Copy.MergedCells && options.Copy.CellLimit == 1000000, "copy old configuration gets defaults");
            options.Copy.VisibleCellsOnly = true; options.Copy.FormulaFallback = 1; options.Copy.RowHeights = false; options.Copy.CellLimit = 200000;
            Check(store.Save(options), "copy settings save"); var loaded = store.Load().Copy;
            Check(loaded.VisibleCellsOnly && loaded.FormulaFallback == 1 && !loaded.RowHeights && loaded.CellLimit == 200000, "copy settings round trip");
            loaded.CellLimit = int.MaxValue; loaded.AppearanceLimit = 0; loaded.Normalize();
            Check(loaded.CellLimit == 5000000 && loaded.AppearanceLimit == 100, "copy limits normalized");
        }
        finally { File.Delete(file); if (File.Exists(file + ".bak")) File.Delete(file + ".bak"); }
        var window = new PreferencesWindow(new ToolPreferences(), new NavigationSettings(), () => "Saved", () => { });
        Check(window.FindName("Category_copy") != null, "copy unified settings category exists"); window.Close();
    }

    private static int CopyExcelTests(string addIn)
    {
        dynamic app = null; int pid = 0;
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            app.Visible = false; app.DisplayAlerts = false; app.AutomationSecurity = 3;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out pid);
            Console.WriteLine("Copy test owned Excel pid=" + pid + "; version=" + app.Version);
            var gateway = new ExcelGateway(() => (object)app);
            dynamic book = app.Workbooks.Add(-4167), sheet = book.Worksheets.Item[1]; sheet.Name = "O'Brien";
            book.Styles.Item["Normal"].Font.Name = "Arial"; book.Styles.Item["Normal"].Font.Size = 12;
            book.Date1904 = true;
            sheet.Range["B2"].Value2 = 12.5; sheet.Range["C2"].Value2 = 7.5; sheet.Range["D2"].Formula = "=B2+C2";
            sheet.Range["D3"].Formula = "=SUM($B$2:C2)"; sheet.Range["D4"].Formula = "='O''Brien'!B2*2";
            sheet.Range["D5"].Formula = "=Z99+1"; sheet.Range["Z99"].Value2 = 9;
            sheet.Range["D6"].Formula = "=INDIRECT(\"B2\")";
            sheet.Range["B3"].Formula = "'00123"; sheet.Range["B4"].Formula = "'=1+1";
            sheet.Range["C3"].Value2 = true; sheet.Range["C4"].Formula = "=NA()";
            sheet.Range["B5:C5"].Merge(); sheet.Range["B5"].Value2 = "合并标题";
            sheet.Range["B2:D6"].Font.Name = "Arial"; sheet.Range["B2:D6"].Font.Size = 13;
            sheet.Range["B2:D6"].Interior.Color = 0xDDFFDD; sheet.Range["B2:D6"].Borders.LineStyle = 1;
            sheet.Range["B2"].NumberFormat = "0.00%"; sheet.Range["B2"].HorizontalAlignment = -4108;
            sheet.Rows[2].RowHeight = 35; sheet.Rows[3].RowHeight = 27; sheet.Columns[2].ColumnWidth = 20; sheet.Columns[3].ColumnWidth = 25;
            // Office 365 build 20326 can crash when AddComment targets an inactive cell.
            sheet.Range["B2"].Select(); sheet.Range["B2"].AddComment("复制备注"); sheet.Range["B2:D6"].Select(); sheet.Calculate();
            app.Calculation = -4135; app.EnableEvents = false; app.ScreenUpdating = false; app.StatusBar = "original status"; book.Saved = true;
            var options = new CopyOptions();
            var sourceFormulas = ((object[,])sheet.Range["B2:D6"].Formula).Cast<object>().ToArray();
            var result = gateway.CopySelection(true, options); dynamic copied = result.Workbook, output = copied.Worksheets.Item[1];
            Check(result.KeptFormulas == 4 && result.ValueFormulas == 2, "copy counts internal and fallback formulas");
            Check((string)output.Range["D2"].Formula == "=B2+C2" && (double)output.Range["D2"].Value2 == 20, "copy internal formula and result");
            Check((string)output.Range["D4"].Formula == "='O''Brien'!B2*2", "copy same-sheet quoted formula preserved");
            Check(!(bool)output.Range["D5"].HasFormula && (double)output.Range["D5"].Value2 == 10 && !(bool)output.Range["D6"].HasFormula, "copy external and dynamic formulas frozen");
            Check((string)output.Range["B3"].Value2 == "00123" && (string)output.Range["B4"].Value2 == "=1+1" && !(bool)output.Range["B4"].HasFormula, "copy literal text and formula-looking strings exact");
            Check((bool)output.Range["C3"].Value2 && (bool)app.WorksheetFunction.IsError(output.Range["C4"]), "copy boolean and error type");
            Check((bool)output.Range["B5"].MergeCells && (string)output.Range["B5"].Value2 == "合并标题", "copy merged layout and value");
            Check((double)output.Rows[2].RowHeight == 35 && (double)output.Rows[3].RowHeight == 27 && Math.Abs((double)output.Columns[3].ColumnWidth - 25) < .2, "copy mixed row heights and widths");
            Check(Math.Abs((double)output.Columns[3].Width - (double)sheet.Columns[3].Width) < 1, "copy physical width under customized Normal font");
            Check((string)output.Range["B2"].NumberFormat == "0.00%" && (int)output.Range["B2"].HorizontalAlignment == -4108 && (string)output.Range["B2"].Comment.Text() == "复制备注", "copy number format alignment and note");
            Check((bool)copied.Date1904 && copied.LinkSources(1) == null, "copy date system and no links");
            Check(sourceFormulas.SequenceEqual(((object[,])sheet.Range["B2:D6"].Formula).Cast<object>()) && (int)app.Calculation == -4135 && !(bool)app.EnableEvents && !(bool)app.ScreenUpdating && (string)app.StatusBar == "original status", "copy leaves source contents and restores all host state");
            copied.Close(false); book.Activate(); sheet.Activate(); sheet.Range["B2:D6"].Select();
            result = gateway.CopySelection(false, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check(result.KeptFormulas == 0 && IsNoFormulas(output.Range["B2:D6"]) && (double)output.Range["D2"].Value2 == 20, "values mode removes all formulas"); copied.Close(false);
            book.Activate(); sheet.Activate(); sheet.Range["B2:D6"].Select(); int before = app.Workbooks.Count;
            try { gateway.CopySelection(true, new CopyOptions { FormulaFallback = 1 }); Check(false, "strict policy"); }
            catch (InvalidOperationException) { Check((int)app.Workbooks.Count == before && (int)app.Calculation == -4135 && !(bool)app.EnableEvents, "strict policy creates no partial output and restores host"); }
            try { gateway.CopySelection(false, options, (stage, done, total) => { if (stage == "复制数值与排版") throw new OperationCanceledException(); }); Check(false, "cancel"); }
            catch (OperationCanceledException) { Check((int)app.Workbooks.Count == before && (string)app.StatusBar == "original status", "cancellation after output creation rolls back"); }
            sheet.Columns[3].Hidden = true; sheet.Range["B4:C6"].Select();
            try { gateway.CopySelection(false, new CopyOptions { VisibleCellsOnly = true }); Check(false, "partial merge"); } catch (InvalidOperationException) { Check((int)app.Workbooks.Count == before, "partial merge rejected before output"); }
            sheet.Columns[3].Hidden = false;
            sheet.Range["A1:Z99"].Select();
            try { gateway.CopySelection(false, new CopyOptions { CellLimit = 100 }); Check(false, "limit"); } catch (InvalidOperationException) { Check((int)app.Workbooks.Count == before, "cell limit rejects before mutation"); }

            // Discontiguous selection must not copy the gap and formulas may reference both areas.
            sheet.Range["F2"].Formula = "=B2+C2"; app.Union(sheet.Range["B2:C2"], sheet.Range["F2"]).Select();
            result = gateway.CopySelection(true, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check(output.Range["D2"].Value2 == null && (string)output.Range["F2"].Formula == "=B2+C2", "discontiguous positions and cross-area internal formula"); copied.Close(false);
            book.Activate(); sheet.Activate();
            sheet.Rows[3].Hidden = true; sheet.Range["B2:D6"].Select();
            result = gateway.CopySelection(false, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check((bool)output.Rows[3].Hidden && (string)output.Range["B3"].Value2 == "00123", "hidden rows included with data and state"); copied.Close(false);
            book.Activate(); sheet.Activate(); sheet.Range["B2:D6"].Select();
            result = gateway.CopySelection(false, new CopyOptions { HiddenState = false }); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check(!(bool)output.Rows[3].Hidden && (string)output.Range["B3"].Value2 == "00123", "hidden state option off keeps hidden-row data visible"); copied.Close(false);
            book.Activate(); sheet.Activate(); sheet.Range["B2:D4"].Select();
            result = gateway.CopySelection(true, new CopyOptions { VisibleCellsOnly = true }); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check(output.Range["B3"].Value2 == null && !(bool)output.Rows[3].Hidden && (string)output.Range["B4"].Value2 == "=1+1", "visible-only leaves hidden-row gaps at original positions"); copied.Close(false);
            book.Activate(); sheet.Activate(); sheet.Rows[3].Hidden = false;
            sheet.Range["H1:H3"].Value2 = 10;
            dynamic rule = sheet.Range["H1:H3"].FormatConditions.Add(1, 5, "5"); rule.Interior.Color = 255; rule.Font.Bold = true;
            sheet.Range["H1:H3"].Select(); result = gateway.CopySelection(false, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check((int)output.Range["H1:H3"].FormatConditions.Count == 0 && (double)output.Range["H2"].Interior.Color == 255 && (bool)output.Range["H2"].Font.Bold, "conditional formatting frozen without rules"); copied.Close(false);
            book.Activate(); sheet.Activate(); sheet.Range["H1:H3"].Select();
            result = gateway.CopySelection(false, new CopyOptions { ConditionalFormatting = 1, Notes = false, CellFormats = false, NumberFormats = false }); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check((int)output.Range["H1:H3"].FormatConditions.Count == 0 && (string)output.Range["H2"].NumberFormat == "General", "copy options disable formatting"); copied.Close(false);
            book.Activate(); sheet.Activate();
            sheet.Range["J1:J3"].FormulaArray = "=B2:B4"; sheet.Range["J1:J3"].Select();
            result = gateway.CopySelection(true, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check(IsNoFormulas(output.Range["J1:J3"]) && result.ValueFormulas == 3, "legacy array formulas safely frozen"); copied.Close(false);
            book.Activate(); sheet.Activate();
            sheet.Range["L1:M5"].Value2 = new object[,] { { "Key", "Value" }, { "yes", 1 }, { "no", 2 }, { "yes", 3 }, { "no", 4 } };
            sheet.Range["L3:M3"].Interior.Color = 255; sheet.Range["L4:M4"].Interior.Color = 65280; sheet.Range["L5:M5"].Interior.Color = 16711680;
            sheet.Range["L1:M5"].AutoFilter(1, "yes"); sheet.Range["L1:M5"].Select();
            result = gateway.CopySelection(false, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check((double)output.Range["M3"].Value2 == 2 && (double)output.Range["M5"].Value2 == 4, "filtered rows retain data when visible-only is off");
            Check((double)output.Range["M3"].Interior.Color == 255 && (double)output.Range["M4"].Interior.Color == 65280 && (double)output.Range["M5"].Interior.Color == 16711680, "filtered row formats remain at original coordinates"); copied.Close(false);
            book.Activate(); sheet.Activate(); sheet.AutoFilterMode = false;
            sheet.Protect("test"); sheet.Range["B2:C2"].Select(); result = gateway.CopySelection(false, options); copied = result.Workbook;
            Check((bool)sheet.ProtectContents && (double)copied.Worksheets.Item[1].Range["B2"].Value2 == 12.5, "protected source readable without changing protection"); copied.Close(false); sheet.Unprotect("test");

            book.Activate(); sheet.Activate(); sheet.Range["B2:D6"].Select();
            result = gateway.CopySelection(false, new CopyOptions { CellFormats = false, NumberFormats = true, MergedCells = true, Notes = false, RowHeights = false, ColumnWidths = false, KeepSourceActive = true }); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check((string)output.Range["B2"].NumberFormat == "0.00%" && (bool)output.Range["B5"].MergeCells && output.Range["B2"].Comment == null, "independent number format merge and notes options");
            Check((string)app.ActiveWorkbook.Name == (string)book.Name && (double)output.Rows[2].RowHeight != 35, "keep source active and disable row height"); copied.Close(false);
            sheet.Range["B2:D6"].Select(); result = gateway.CopySelection(false, new CopyOptions { MergedCells = false, NumberFormats = false }); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check(!(bool)output.Range["B5"].MergeCells && (string)output.Range["B2"].NumberFormat == "General", "disable merge and number format independently"); copied.Close(false);
            book.Activate(); sheet.Activate(); sheet.Range["P1"].Value2 = "富文本测试"; sheet.Range["P1"].Characters(1, 2).Font.Bold = true; sheet.Range["P1"].Select();
            result = gateway.CopySelection(false, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check((string)output.Range["P1"].Value2 == "富文本测试" && (bool)output.Range["P1"].Characters(1, 2).Font.Bold && !(bool)output.Range["P1"].Characters(3, 2).Font.Bold, "rich text character formatting preserved"); copied.Close(false);
            book.Activate(); sheet.Activate();
            dynamic other = book.Worksheets.Add(); other.Name = "Other"; other.Range["A1"].Value2 = 99;
            book.Names.Add("InputValue", "='O''Brien'!$B$2"); sheet.Range["P2"].Formula = "=Other!A1"; sheet.Range["P3"].Formula = "=InputValue*2";
            sheet.Activate(); sheet.Range["P2:P3"].Select(); sheet.Calculate(); result = gateway.CopySelection(true, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check(IsNoFormulas(output.Range["P2:P3"]) && (double)output.Range["P2"].Value2 == 99 && (double)output.Range["P3"].Value2 == 25 && copied.LinkSources(1) == null, "cross-sheet and named formulas frozen and links removed"); copied.Close(false);
            book.Activate(); sheet.Activate(); sheet.Range["Q1:Q3"].Value2 = 10; sheet.Range["Q1:Q3"].FormatConditions.AddDatabar(); sheet.Range["Q1:Q3"].Select();
            try { gateway.CopySelection(false, options); Check(false, "databar"); } catch (InvalidOperationException) { Check((int)app.Workbooks.Count == before, "data bars reject explicitly before mutation"); }
            result = gateway.CopySelection(false, new CopyOptions { ConditionalFormatting = 1 }); copied = result.Workbook;
            Check((int)copied.Worksheets.Item[1].Range["Q1:Q3"].FormatConditions.Count == 0, "data bars may be deliberately removed by setting"); copied.Close(false);
            book.Activate(); sheet.Activate(); sheet.Range["H1:H3"].Select();
            try { gateway.CopySelection(false, options, (stage, done, total) => { if (stage == "固定条件格式外观") throw new InvalidOperationException("Injected failure"); }); Check(false, "appearance fault"); }
            catch (InvalidOperationException) { Check((int)app.Workbooks.Count == before && !(bool)app.EnableEvents, "mid-write failure closes result and restores events"); }
            sheet.Range["R1:S3"].Value2 = new object[,] { { "Input", "Double" }, { 2, 4 }, { 3, 6 } };
            dynamic table = sheet.ListObjects.Add(1, sheet.Range["R1:S3"], Type.Missing, 1); table.Name = "GPTestTable";
            sheet.Range["S2:S3"].Formula = "=[@Input]*2"; sheet.Calculate(); sheet.Range["R1:S3"].Select();
            result = gateway.CopySelection(true, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check((int)output.ListObjects.Count == 0 && IsNoFormulas(output.Range["S2:S3"]) && (double)output.Range["S3"].Value2 == 6 && copied.LinkSources(1) == null, "table converted to ordinary cells and structured formulas frozen"); copied.Close(false);
            book.Activate(); sheet.Activate(); bool spills = false;
            try { sheet.Range["U1"].Formula2 = "=SEQUENCE(3)"; sheet.Calculate(); spills = sheet.Range["U1"].Value2 is double; } catch (COMException) { }
            if (spills)
            {
                sheet.Range["U1:U3"].Select(); result = gateway.CopySelection(true, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
                Check(IsNoFormulas(output.Range["U1:U3"]) && (double)output.Range["U3"].Value2 == 3, "modern spill formula and children frozen together"); copied.Close(false);
                book.Activate(); sheet.Activate(); sheet.Range["U2:U3"].Select(); result = gateway.CopySelection(false, options); copied = result.Workbook;
                Check((double)copied.Worksheets.Item[1].Range["U3"].Value2 == 3, "partial spill children copied as values"); copied.Close(false);
                book.Activate(); sheet.Activate();
            }

            // Whole-column selection is bounded to UsedRange; used-range mode is independently tested.
            dynamic tiny = book.Worksheets.Add(); tiny.Name = "Small"; tiny.Range["A1:B2"].Value2 = 2; tiny.Columns[1].Select();
            result = gateway.CopySelection(false, options); copied = result.Workbook;
            Check(result.Cells == 2 && copied.Worksheets.Item[1].Range["B1"].Value2 == null, "whole column trimmed without copying neighboring column"); copied.Close(false);
            book.Activate(); tiny.Activate(); tiny.Range["Z99"].Select(); result = gateway.CopySelection(false, new CopyOptions { EmptySelectionUsesUsedRange = true }); copied = result.Workbook;
            Check(result.Cells == 4, "blank-cell fallback uses actual UsedRange"); copied.Close(false);
            book.Activate(); tiny.Activate(); tiny.Range["A1"].Select(); result = gateway.CopySelection(false, new CopyOptions { Scope = 1 }); copied = result.Workbook;
            Check(result.Cells == 4, "explicit used-range scope"); copied.Close(false);

            // Performance and correctness on a realistic rectangular table.
            dynamic perf = book.Worksheets.Add(); perf.Name = "Performance";
            var data = new object[20000, 5]; for (int r = 0; r < 20000; r++) for (int c = 0; c < 5; c++) data[r, c] = r + c;
            perf.Range["A1:E20000"].Value2 = data; perf.Range["F1:F20000"].FormulaR1C1 = "=RC[-5]+RC[-4]"; perf.Calculate(); perf.Range["A1:F20000"].Select();
            result = gateway.CopySelection(true, options); copied = result.Workbook; output = copied.Worksheets.Item[1];
            Check(result.Cells == 120000 && result.KeptFormulas == 20000 && (double)output.Range["F20000"].Value2 == 39999, "120k cell / 20k formula copy correctness");
            Console.WriteLine("PERF internal cells=120000 formulas=20000 milliseconds=" + result.ElapsedMilliseconds); copied.Close(false);
            book.Activate(); perf.Activate(); perf.Range["A1:F20000"].Select(); result = gateway.CopySelection(false, options); copied = result.Workbook;
            Check(IsNoFormulas(copied.Worksheets.Item[1].Range["F1:F20000"]), "120k values output contains no formulas"); Console.WriteLine("PERF values cells=120000 milliseconds=" + result.ElapsedMilliseconds); copied.Close(false);
            book.Activate(); perf.Activate(); perf.Range["A20001:E20001"].Merge(); perf.Range["A20001"].Value2 = "Footer"; perf.Range["A1:F20001"].Select();
            result = gateway.CopySelection(false, options); copied = result.Workbook;
            Check((bool)copied.Worksheets.Item[1].Range["A20001"].MergeCells && (double)copied.Worksheets.Item[1].Range["F20000"].Value2 == 39999, "large range with merged footer avoids per-cell boundary scan"); Console.WriteLine("PERF merged cells=120006 milliseconds=" + result.ElapsedMilliseconds); copied.Close(false);
            if (!string.IsNullOrEmpty(addIn))
            {
                app.EnableEvents = true; app.ScreenUpdating = true; book.Activate(); perf.Activate(); perf.Range["A1:F2"].Select();
                Check((bool)app.RegisterXLL(addIn), "copy XLL packed registration");
                // The queued macro creates a workbook before it finishes copying. Wait for its
                // completion announcement, not merely Workbooks.Count (COM can interleave calls).
                app.StatusBar = "waiting for copy command";
                app.Run("GridPilot_CopyValues"); PumpUntil(() => (int)app.Workbooks.Count > before && Convert.ToString((object)app.StatusBar).StartsWith("GridPilot · 已原样复制", StringComparison.Ordinal), 20000);
                Check(IsNoFormulas(app.ActiveSheet.Range["A1:F2"]), "copy values registered button command"); app.ActiveWorkbook.Close(false);
                book.Activate(); perf.Activate(); perf.Range["A1:F2"].Select(); app.StatusBar = "waiting for copy command";
                app.Run("GridPilot_CopyInternal"); PumpUntil(() => (int)app.Workbooks.Count > before && Convert.ToString((object)app.StatusBar).StartsWith("GridPilot · 已原样复制", StringComparison.Ordinal), 20000);
                Check((bool)app.ActiveSheet.Range["F1"].HasFormula, "copy internal registered button command"); app.ActiveWorkbook.Close(false);
            }
            // Above-32767 row regression for the old VBA Integer loop limit.
            book.Activate(); perf.Activate(); perf.Range["H1:H40000"].Value2 = 7; perf.Range["H1:H40000"].Select();
            result = gateway.CopySelection(false, options); copied = result.Workbook;
            Check((double)copied.Worksheets.Item[1].Range["H40000"].Value2 == 7, "copy exceeds VBA Integer row limit"); copied.Close(false);
            book.Close(false);
        }
        finally
        {
            if (app != null) { try { app.DisplayAlerts = false; app.Quit(); } catch(COMException) { } Marshal.FinalReleaseComObject((object)app); app = null; }
        }
        // Return from the COM scope before testing finalizer-based cleanup (caller does this).
        Console.WriteLine("COPY EXCEL PASS " + checks + " assertions; owned pid=" + pid);
        return pid;
    }
    private static void CopyExcelExitCheck(int pid)
    {
        for (int i = 0; i < 2; i++) { GC.Collect(); var finalizers = Task.Run(() => GC.WaitForPendingFinalizers()); PumpUntil(() => finalizers.IsCompleted, 10000); }
        try { using (var process = Process.GetProcessById(pid)) { PumpUntil(() => process.HasExited, 15000); Check(true, "copy owned Excel exits after Quit"); } }
        catch (ArgumentException) { Check(true, "copy owned Excel already exited"); }
        Console.WriteLine("COPY TOTAL PASS " + checks + " assertions");
    }
    private static bool IsNoFormulas(dynamic range) => range.HasFormula is bool && !(bool)range.HasFormula;
}
