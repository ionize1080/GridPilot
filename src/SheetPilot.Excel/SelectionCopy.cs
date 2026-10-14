using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Xml.Linq;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed class CopyResult
    {
        public object Workbook { get; internal set; }
        public long Cells { get; internal set; }
        public int KeptFormulas { get; internal set; }
        public int ValueFormulas { get; internal set; }
        public long ElapsedMilliseconds { get; internal set; }
        public Dictionary<string, int> FallbackReasons { get; } = new Dictionary<string, int>();
        public string Summary => "已原样复制 " + Cells.ToString("N0") + " 个单元格到新工作簿；保留公式 " + KeptFormulas + "，公式转值 " + ValueFormulas + "。耗时 " + (ElapsedMilliseconds / 1000.0).ToString("0.00") + " 秒；未保存文件。" +
            (FallbackReasons.Count == 0 ? "" : " 转值原因：" + string.Join("；", FallbackReasons.Select(x => x.Key + " " + x.Value)));
    }

    public sealed partial class ExcelGateway
    {
        private const int CopyBlockCells = 32768;
        private sealed class FormulaBlock { public CellRect Rect; public object[,] Formulas, CachedValues; }

        /// <summary>Creates only a new, unsaved workbook. On any failure/cancellation closes that workbook.
        /// All COM stays on the owning STA. The progress hook may throw OperationCanceledException.</summary>
        public CopyResult CopySelection(bool keepInternalFormulas, CopyOptions preferences, Action<string, long, long> progress = null)
        {
            CheckThread(); var options = preferences.Clone(); options.Normalize();
            dynamic app = application(), selection = app.Selection;
            if (selection == null || Microsoft.VisualBasic.Information.TypeName((object)selection).TrimStart('_') != "Range")
                throw new InvalidOperationException("请先选择普通工作表中的单元格区域。");
            dynamic source = selection.Worksheet, sourceBook = source.Parent, sourceWindow = app.ActiveWindow;
            dynamic range = options.Scope == 1 ? source.UsedRange : selection;
            if (options.Scope == 0 && options.EmptySelectionUsesUsedRange && (double)selection.CountLarge == 1 && selection.Value2 == null) range = source.UsedRange;
            var areaObjects = new List<object>(); var areas = new List<CellRect>();
            int count = (int)range.Areas.Count;
            if (count > 4096) throw new InvalidOperationException("选择区域过于零散，请缩小范围（最多4096块）。");
            for (int i = 1; i <= count; i++)
            {
                dynamic area = range.Areas.Item[i];
                if (options.TrimWholeRowsColumns && ((int)area.Rows.Count == (int)source.Rows.Count || (int)area.Columns.Count == (int)source.Columns.Count))
                    area = app.Intersect(area, source.UsedRange);
                if (area == null) continue;
                if (options.VisibleCellsOnly)
                {
                    object visible = null; try { visible = area.SpecialCells(12); } catch (COMException) { }
                    area = visible == null ? null : app.Intersect(area, visible);
                }
                if (area == null) continue;
                for (int j = 1; j <= (int)area.Areas.Count; j++)
                {
                    dynamic part = area.Areas.Item[j]; areaObjects.Add((object)part); areas.Add(CopyRect(part));
                    if (areas.Count > 4096) throw new InvalidOperationException("可见区域过于零散，请缩小范围（最多4096块）。");
                }
            }
            long cells = areas.Sum(x => x.Count);
            if (cells == 0) throw new InvalidOperationException("所选范围没有可复制的单元格。");
            if (cells > options.CellLimit) throw new InvalidOperationException("本次包含 " + cells.ToString("N0") + " 个单元格，超过原样复制设置的上限 " + options.CellLimit.ToString("N0") + "。请缩小范围或调整设置。");
            var timer = Stopwatch.StartNew(); var result = new CopyResult { Cells = cells };
            bool events = app.EnableEvents, screen = app.ScreenUpdating, alerts = app.DisplayAlerts;
            bool copyObjects = app.CopyObjectsWithCells;
            object status = app.StatusBar; int cancel = app.EnableCancelKey;
            dynamic output = null; bool succeeded = false;
            long lastStatus = -100;
            Action<string, long> report = (stage, done) => { progress?.Invoke(stage, done, cells); if (timer.ElapsedMilliseconds - lastStatus >= 80) { app.StatusBar = "GridPilot · " + stage + " " + done.ToString("N0") + " / " + cells.ToString("N0") + " · Esc 取消"; lastStatus = timer.ElapsedMilliseconds; } };
            try
            {
                app.EnableEvents = false; app.ScreenUpdating = false; app.DisplayAlerts = false; app.EnableCancelKey = 2;
                app.CopyObjectsWithCells = false;
                report("检查范围", 0);
                foreach (var item in areaObjects) ValidateCopyMerges(item, areas, report);
                var blocks = CaptureCopyFormulas(source, areaObjects, areas, keepInternalFormulas, options, result, report);
                var appearances = CaptureCopyAppearance(app, source, areaObjects, options, report);
                output = app.Workbooks.Add(-4167); dynamic target = output.Worksheets.Item[1]; target.Name = source.Name;
                report("初始化结果", 0);
                target.EnableCalculation = false;
                // A1 positions and original sheet name are invariant, including absolute references.
                // Match source workbook date system to preserve dates stored as serial values.
                output.Date1904 = sourceBook.Date1904;
                report("复制主题", 0);
                if (options.CellFormats) CopyTheme(sourceBook, output);
                long done = 0;
                for (int i = 0; i < areas.Count; i++)
                {
                    dynamic from = areaObjects[i], to = target.Range[areas[i].Address];
                    report("复制数值与排版", done);
                    CopyLayout(source, target, areas[i], report);
                    report("清理单元格规则", done);
                    to.FormatConditions.Delete(); to.Validation.Delete(); to.Hyperlinks.Delete();
                    while ((int)target.ListObjects.Count > 0) target.ListObjects.Item[1].Unlist();
                    if (!options.CellFormats)
                    {
                        to.ClearFormats();
                        if (options.NumberFormats) CopyNumberFormats(source, target, areas[i], report);
                        if (options.MergedCells) CopyMerges(from, target, report);
                    }
                    if (!options.MergedCells) to.UnMerge();
                    if (!options.NumberFormats) to.NumberFormat = "General";
                    if (!options.Notes) to.ClearComments();
                    FreezeCopyFormulas(source, target, areas[i], report);
                    done += areas[i].Count;
                }
                ApplyCopyAppearance(target, appearances, options, report);
                report("移除名称与链接", cells);
                // Range.Copy can import names or link bookkeeping even after contents are cleared.
                // No retained formula depends on names, validation, connections or other workbooks.
                var importedNames = new List<object>(); foreach (dynamic name in output.Names) importedNames.Add((object)name);
                foreach (dynamic name in importedNames)
                {
                    string nameText = name.Name, refersTo = name.RefersTo;
                    // Modern Excel exposes non-deletable compatibility function markers.
                    if (nameText.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase) && refersTo == "=#NAME?") continue;
                    name.Delete();
                }
                report("移除连接", cells);
                while ((int)output.Connections.Count > 0) output.Connections.Item(1).Delete();
                report("移除外部链接", cells);
                object links = output.LinkSources(1);
                if (links is Array linkArray) foreach (object link in linkArray) output.BreakLink(Convert.ToString(link), 1);
                foreach (var block in blocks)
                {
                    report("恢复内部公式", cells);
                    target.Range[block.Rect.Address].Formula = block.Formulas;
                }
                foreach (var rect in areas)
                {
                    if (options.RowHeights) CopyDimension(source, target, rect.Row, rect.Bottom, true, false, report);
                    if (options.ColumnWidths) CopyDimension(source, target, rect.Column, rect.Right, false, false, report);
                    if (options.HiddenState && !options.VisibleCellsOnly)
                    {
                        CopyHidden(source, target, rect.Row, rect.Bottom, true, report);
                        CopyHidden(source, target, rect.Column, rect.Right, false, report);
                    }
                    else
                    {
                        target.Rows[rect.Row + ":" + rect.Bottom].Hidden = false;
                        target.Columns[CellRect.ColumnName(rect.Column) + ":" + CellRect.ColumnName(rect.Right)].Hidden = false;
                    }
                }
                if (options.ViewSettings)
                {
                    dynamic window = output.Windows.Item[1];
                    window.DisplayGridlines = sourceWindow.DisplayGridlines; window.DisplayHeadings = sourceWindow.DisplayHeadings; window.Zoom = sourceWindow.Zoom;
                }
                report("完成校验", cells);
                if (output.LinkSources(1) != null) throw new InvalidOperationException("结果中检测到外部工作簿链接，已取消复制，请检查源范围的格式或批注。");
                target.EnableCalculation = true; target.Calculate();
                // Freeze every drifting retained result, then recalculate to a fixed point:
                // freezing a nested subtotal can in turn change its parent subtotal.
                int remaining=result.KeptFormulas;
                while(true)
                {
                    int drift=0;
                    foreach(var block in blocks)
                    {
                        report("复核保留公式",cells);object raw=target.Range[block.Rect.Address].Value2;var values=raw as object[,];
                        for(int r=0;r<block.Rect.Rows;r++)for(int c=0;c<block.Rect.Columns;c++)
                        {
                            if(block.Formulas[r,c]==null)continue;
                            object actual=values==null?raw:values[r+1,c+1],expected=block.CachedValues[r,c];
                            if(CopyEquivalent(expected,actual))continue;
                            if(options.FormulaFallback==1)throw new InvalidOperationException("保留公式在目标工作簿重算后结果改变，已停止复制。");
                            target.Range[DataCell.Address(block.Rect.Row+r,block.Rect.Column+c)].Value2=LiteralCopyValue(expected);
                            block.Formulas[r,c]=null;drift++;result.KeptFormulas--;result.ValueFormulas++;
                            const string reason="重算结果变化（嵌套小计、可见性等）";result.FallbackReasons.TryGetValue(reason,out int n);result.FallbackReasons[reason]=n+1;
                        }
                    }
                    if(drift==0)break;remaining-=drift;if(remaining<0)throw new InvalidOperationException("公式校验未收敛。");target.Calculate();
                }
                if (options.KeepSourceActive) { sourceBook.Activate(); source.Activate(); }
                else { output.Activate(); target.Activate(); target.Range[areas[0].Address].Select(); }
                result.Workbook = (object)output; succeeded = true;
            }
            finally
            {
                if (!succeeded)
                {
                    if (output != null) { try { output.Close(false); } catch (COMException) { } }
                    try { sourceBook.Activate(); source.Activate(); } catch (COMException) { }
                }
                // Each restoration runs even if another COM restoration fails.
                try { app.CutCopyMode = false; }
                finally { try { app.CopyObjectsWithCells = copyObjects; }
                finally { try { app.EnableEvents = events; }
                finally { try { app.DisplayAlerts = alerts; }
                finally { try { app.ScreenUpdating = screen; }
                finally { try { app.EnableCancelKey = cancel; } finally { app.StatusBar = status; } } } } } }
            }
            result.ElapsedMilliseconds = timer.ElapsedMilliseconds;
            return result;
        }

        private static bool CopyEquivalent(object a,object b)
        {if(a is ErrorWrapper x && b is ErrorWrapper y)return x.ErrorCode==y.ErrorCode;return MaskPlanner.Equivalent(a,b);}
        private static CellRect CopyRect(dynamic range) => new CellRect((int)range.Row, (int)range.Column, (int)range.Rows.Count, (int)range.Columns.Count);
        private static void CopyLayout(dynamic source, dynamic target, CellRect rect, Action<string, long> report)
        {
            if (!(bool)source.FilterMode) { source.Range[rect.Address].Copy(target.Range[rect.Address]); return; }
            // Excel also compacts formatting in Range.Copy(Destination) when an AutoFilter hides
            // rows. Copy contiguous visibility runs separately, at their original coordinates.
            var bounds = new SortedSet<int> { rect.Row, rect.Bottom + 1 };
            object visible = null;
            try { visible = source.Rows[rect.Row + ":" + rect.Bottom].SpecialCells(12); } catch (COMException) { }
            if (visible != null)
            {
                dynamic parts = ((dynamic)visible).Areas;
                for (int i = 1; i <= (int)parts.Count; i++) { CellRect part = CopyRect(parts.Item[i]); bounds.Add(Math.Max(rect.Row, part.Row)); bounds.Add(Math.Min(rect.Bottom + 1, part.Bottom + 1)); }
            }
            if (bounds.Count > 4097) throw new InvalidOperationException("筛选后的行过于零散，复制排版需要超过4096个分块。请缩小选择范围。");
            var rows = bounds.ToArray();
            for (int i = 0; i < rows.Length - 1; i++)
            {
                report("复制筛选行排版", rows[i] - rect.Row);
                var run = new CellRect(rows[i], rect.Column, rows[i + 1] - rows[i], rect.Columns);
                dynamic from = source.Range[run.Address];
                ValidateCopyMerges(from, new[] { run }, report);
                from.Copy(target.Range[run.Address]);
            }
        }
        private static void FreezeCopyFormulas(dynamic source, dynamic target, CellRect rect, Action<string, long> report)
        {
            dynamic range = source.Range[rect.Address];
            if (!IsFalse((object)range.HasFormula))
            {
                object formulas = null; try { formulas = range.SpecialCells(-4123); } catch (COMException) { }
                dynamic intersection = formulas == null ? null : source.Application.Intersect(range, formulas);
                if (intersection != null) for (int i = 1; i <= (int)intersection.Areas.Count; i++)
                {
                    CellRect block = CopyRect(intersection.Areas.Item[i]);
                    target.Range[block.Address].ClearContents();
                    CopyValues(source, target, block, report);
                }
            }
            FreezeSpills(source, target, rect, report);
        }
        private static void FreezeSpills(dynamic source, dynamic target, CellRect rect, Action<string, long> report)
        {
            object spilled;
            try { spilled = source.Range[rect.Address].HasSpill; }
            catch (COMException) { return; }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { return; }
            if (IsFalse(spilled)) return;
            report("固定溢出结果", 0);
            if (IsTrue(spilled))
            {
                target.Range[rect.Address].ClearContents(); CopyValues(source, target, rect, report); return;
            }
            if (rect.Rows > 1)
            {
                int n = rect.Rows / 2;
                FreezeSpills(source, target, new CellRect(rect.Row, rect.Column, n, rect.Columns), report);
                FreezeSpills(source, target, new CellRect(rect.Row + n, rect.Column, rect.Rows - n, rect.Columns), report);
            }
            else if (rect.Columns > 1)
            {
                int n = rect.Columns / 2;
                FreezeSpills(source, target, new CellRect(rect.Row, rect.Column, 1, n), report);
                FreezeSpills(source, target, new CellRect(rect.Row, rect.Column + n, 1, rect.Columns - n), report);
            }
        }
        private static void CopyValues(dynamic source, dynamic target, CellRect rect, Action<string, long> report)
        {
            int step = Math.Max(1, CopyBlockCells / rect.Columns);
            for (int start = 0; start < rect.Rows; start += step)
            {
                report("写入数值", (long)start * rect.Columns);
                var block = new CellRect(rect.Row + start, rect.Column, Math.Min(step, rect.Rows - start), rect.Columns);
                object raw = source.Range[block.Address].Value2;
                if (raw is object[,] values)
                {
                    for (int r = values.GetLowerBound(0); r <= values.GetUpperBound(0); r++)
                    for (int c = values.GetLowerBound(1); c <= values.GetUpperBound(1); c++) values[r, c] = LiteralCopyValue(values[r, c]);
                    target.Range[block.Address].Value2 = values;
                }
                else target.Range[block.Address].Value2 = LiteralCopyValue(raw);
            }
        }
        private static object LiteralCopyValue(object value)
        {
            if (value is string text) return "'" + text;
            // Excel numeric cells marshal as Double; VT_ERROR marshals as its HRESULT Int32.
            if (value is int error && (error & unchecked((int)0xffff0000)) == unchecked((int)0x800a0000)) return new ErrorWrapper(error);
            return value;
        }
        private static void CopyNumberFormats(dynamic source, dynamic target, CellRect rect, Action<string, long> report)
        {
            object format = source.Range[rect.Address].NumberFormat;
            if (format is string text) { target.Range[rect.Address].NumberFormat = text; return; }
            report("复制数字格式", 0);
            if (rect.Rows > 1) { int n = rect.Rows / 2; CopyNumberFormats(source, target, new CellRect(rect.Row, rect.Column, n, rect.Columns), report); CopyNumberFormats(source, target, new CellRect(rect.Row + n, rect.Column, rect.Rows - n, rect.Columns), report); }
            else if (rect.Columns > 1) { int n = rect.Columns / 2; CopyNumberFormats(source, target, new CellRect(rect.Row, rect.Column, 1, n), report); CopyNumberFormats(source, target, new CellRect(rect.Row, rect.Column + n, 1, rect.Columns - n), report); }
        }
        private static bool IsTrue(object value) => value is bool b && b;
        private static bool IsFalse(object value) => value is bool b && !b;
        private static void ValidateCopyMerges(dynamic area, IReadOnlyList<CellRect> areas, Action<string, long> report)
        {
            if (IsFalse((object)area.MergeCells)) return;
            CellRect rect = CopyRect(area);
            VisitMerges((object)area.Worksheet, new CellRect(rect.Row, rect.Column, rect.Rows, 1), merge => ValidateMerge(merge, areas), report);
            if (rect.Columns > 1) VisitMerges((object)area.Worksheet, new CellRect(rect.Row, rect.Right, rect.Rows, 1), merge => ValidateMerge(merge, areas), report);
            if (rect.Columns > 2)
            {
                VisitMerges((object)area.Worksheet, new CellRect(rect.Row, rect.Column + 1, 1, rect.Columns - 2), merge => ValidateMerge(merge, areas), report);
                if (rect.Rows > 1) VisitMerges((object)area.Worksheet, new CellRect(rect.Bottom, rect.Column + 1, 1, rect.Columns - 2), merge => ValidateMerge(merge, areas), report);
            }
        }
        private static void ValidateMerge(CellRect merge, IReadOnlyList<CellRect> areas)
        {
            if (!areas.Any(x => x.Contains(merge))) throw new InvalidOperationException("选区只包含合并单元格的一部分（" + merge.Address + "）。请完整选择该合并区域。");
        }
        private static void VisitMerges(dynamic sheet, CellRect rect, Action<CellRect> visit, Action<string, long> report)
        {
            dynamic range = sheet.Range[rect.Address];
            if (IsFalse((object)range.MergeCells)) return;
            report("检查合并布局", 0);
            dynamic first = range.Cells[1, 1];
            if (IsTrue((object)first.MergeCells))
            {
                CellRect merge = CopyRect(first.MergeArea);
                if (merge.Contains(rect)) { visit(merge); return; }
            }
            if (rect.Rows > 1)
            {
                int n = rect.Rows / 2;
                VisitMerges(sheet, new CellRect(rect.Row, rect.Column, n, rect.Columns), visit, report);
                VisitMerges(sheet, new CellRect(rect.Row + n, rect.Column, rect.Rows - n, rect.Columns), visit, report);
            }
            else if (rect.Columns > 1)
            {
                int n = rect.Columns / 2;
                VisitMerges(sheet, new CellRect(rect.Row, rect.Column, 1, n), visit, report);
                VisitMerges(sheet, new CellRect(rect.Row, rect.Column + n, 1, rect.Columns - n), visit, report);
            }
        }

        private static List<FormulaBlock> CaptureCopyFormulas(dynamic source, List<object> areaObjects, List<CellRect> areas, bool keep, CopyOptions options, CopyResult result, Action<string, long> report)
        {
            var blocks = new List<FormulaBlock>(); long done = 0; string sheetName = source.Name;
            foreach (var item in areaObjects)
            {
                dynamic area = item; CellRect areaRect = CopyRect(area);
                if (IsFalse((object)area.HasFormula)) { done += areaRect.Count; continue; }
                int step = Math.Max(1, CopyBlockCells / areaRect.Columns);
                for (int offset = 0; offset < areaRect.Rows; offset += step)
                {
                    report("分析公式", done);
                    var rect = new CellRect(areaRect.Row + offset, areaRect.Column, Math.Min(step, areaRect.Rows - offset), areaRect.Columns);
                    dynamic chunk = source.Range[rect.Address]; object formulaCells = null;
                    try { formulaCells = chunk.SpecialCells(-4123); } catch (COMException) { }
                    if (formulaCells == null) { done += rect.Count; continue; }
                    dynamic intersection = source.Application.Intersect(chunk, formulaCells);
                    if (intersection == null) { done += rect.Count; continue; }
                    object raw = chunk.Formula, cachedRaw=chunk.Value2; var values = raw as object[,];var cached=cachedRaw as object[,];
                    var retained = new object[rect.Rows, rect.Columns];
                    for (int a = 1; a <= (int)intersection.Areas.Count; a++)
                    {
                        dynamic part = intersection.Areas.Item[a]; CellRect partRect = CopyRect(part);
                        bool special = !IsFalse((object)part.HasArray);
                        try { special |= !IsFalse((object)part.HasSpill); } catch (COMException) { } catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                        for (int r = partRect.Row; r <= partRect.Bottom; r++)
                        for (int c = partRect.Column; c <= partRect.Right; c++)
                        {
                            int row = r - rect.Row, col = c - rect.Column;
                            string formula = (values == null ? raw : values[row + values.GetLowerBound(0), col + values.GetLowerBound(1)]) as string;
                            string reason = "仅数值模式";
                            bool allowed = keep && !special && InternalFormula.CanKeep(formula, sheetName, areas, out reason);
                            if (keep && special) reason = "数组或溢出公式";
                            if (allowed) { retained[row, col] = formula; result.KeptFormulas++; }
                            else
                            {
                                result.ValueFormulas++;
                                if (keep)
                                {
                                    if (options.FormulaFallback == 1) throw new InvalidOperationException(CellRect.ColumnName(c) + r + " 的公式无法保留：" + reason + "。未创建结果；可在设置中改为将这类公式转值。");
                                    result.FallbackReasons.TryGetValue(reason, out int n); result.FallbackReasons[reason] = n + 1;
                                }
                            }
                        }
                    }
                    // Coalesce equal horizontal runs across adjacent rows, so full formula columns
                    // or tables are restored in blocks instead of one COM call per cell.
                    var active = new Dictionary<string, CellRect>();
                    for (int r = 0; r <= rect.Rows; r++)
                    {
                        var next = new Dictionary<string, CellRect>();
                        if (r < rect.Rows) for (int c = 0; c < rect.Columns; c++)
                        {
                            if (retained[r, c] == null) continue;
                            int start = c; while (c + 1 < rect.Columns && retained[r, c + 1] != null) c++;
                            string key = start + ":" + c;
                            next[key] = active.TryGetValue(key, out var old) ? new CellRect(old.Row, old.Column, old.Rows + 1, old.Columns) : new CellRect(r, start, 1, c - start + 1);
                        }
                        foreach (var pair in active.Where(x => !next.ContainsKey(x.Key)))
                        {
                            var run = pair.Value; var formulas = new object[run.Rows, run.Columns];var expected=new object[run.Rows,run.Columns];
                            for (int rr = 0; rr < run.Rows; rr++) for (int cc = 0; cc < run.Columns; cc++) {formulas[rr, cc] = retained[run.Row + rr, run.Column + cc];expected[rr,cc]=cached==null?cachedRaw:cached[run.Row+rr+1,run.Column+cc+1];}
                            blocks.Add(new FormulaBlock { Rect = new CellRect(rect.Row + run.Row, rect.Column + run.Column, run.Rows, run.Columns), Formulas = formulas, CachedValues=expected });
                        }
                        active = next;
                    }
                    done += rect.Count;
                }
            }
            return blocks;
        }

        private static void CopyDimension(dynamic source, dynamic target, int first, int last, bool rows, bool hidden, Action<string, long> report)
        {
            report(rows ? "复制行布局" : "复制列布局", 0);
            string address = rows ? first + ":" + last : CellRect.ColumnName(first) + ":" + CellRect.ColumnName(last);
            dynamic from = rows ? source.Rows[address] : source.Columns[address], to = rows ? target.Rows[address] : target.Columns[address];
            object value;
            try { value = hidden ? from.Hidden : rows ? from.RowHeight : from.ColumnWidth; }
            catch (COMException ex) { if (first == last) throw new InvalidOperationException("无法读取布局 " + address + " / " + from.Address + "：" + ex.Message, ex); value = null; }
            if (value != null && value != DBNull.Value)
            {
                if (hidden) to.Hidden = value; else if (rows) to.RowHeight = value; else to.ColumnWidth = value;
            }
            else if (first < last)
            {
                int mid = first + (last - first) / 2;
                CopyDimension(source, target, first, mid, rows, hidden, report); CopyDimension(source, target, mid + 1, last, rows, hidden, report);
            }
        }
        private static void CopyHidden(dynamic source, dynamic target, int first, int last, bool rows, Action<string, long> report)
        {
            report("复制隐藏状态", 0);
            string address = rows ? first + ":" + last : CellRect.ColumnName(first) + ":" + CellRect.ColumnName(last);
            dynamic from = rows ? source.Rows[address] : source.Columns[address], to = rows ? target.Rows[address] : target.Columns[address];
            object visible = null; try { visible = from.SpecialCells(12); } catch (COMException) { }
            to.Hidden = true;
            if (visible == null) return;
            dynamic parts = source.Application.Intersect(from, visible).Areas;
            var spans = new List<Tuple<int, int>>();
            for (int i = 1; i <= (int)parts.Count; i++)
            {
                CellRect rect = CopyRect(parts.Item[i]);
                spans.Add(Tuple.Create(rows ? rect.Row : rect.Column, rows ? rect.Bottom : rect.Right));
            }
            int start = -1, end = -1;
            Action flush = () =>
            {
                if (start < 0) return;
                dynamic range = rows ? target.Rows[start + ":" + end] : target.Columns[CellRect.ColumnName(start) + ":" + CellRect.ColumnName(end)]; range.Hidden = false;
            };
            foreach (var span in spans.OrderBy(x => x.Item1))
            {
                if (span.Item1 <= end + 1) end = Math.Max(end, span.Item2);
                else { flush(); start = span.Item1; end = span.Item2; }
            }
            flush();
        }
        private static void CopyMerges(dynamic from, dynamic target, Action<string, long> report)
        {
            var known = new HashSet<string>();
            VisitMerges((object)from.Worksheet, (CellRect)CopyRect(from), merge => { if (known.Add(merge.Address)) target.Range[merge.Address].Merge(); }, report);
        }

        private sealed class Appearance
        {
            public readonly List<string> Addresses = new List<string>();
            public object Color, Pattern, PatternColor, FontColor, Bold, Italic, Underline, Strike, NumberFormat;
            public readonly object[] Lines = new object[4], Weights = new object[4], Colors = new object[4];
            public string Key => string.Join("|", new[] { Color, Pattern, PatternColor, FontColor, Bold, Italic, Underline, Strike, NumberFormat }.Concat(Lines).Concat(Weights).Concat(Colors));
        }
        private static List<Appearance> CaptureCopyAppearance(dynamic app, dynamic source, List<object> areas, CopyOptions options, Action<string, long> report)
        {
            var groups = new Dictionary<string, Appearance>();
            if (!options.CellFormats || options.ConditionalFormatting == 1) return groups.Values.ToList();
            long total = 0;
            foreach (var item in areas)
            {
                dynamic area = item; object conditional = null;
                try { conditional = area.SpecialCells(-4172); } catch (COMException) { }
                if (conditional == null) continue;
                dynamic cells = app.Intersect(area, conditional); if (cells == null) continue;
                for (int rule = 1; rule <= (int)cells.FormatConditions.Count; rule++)
                {
                    int type = cells.FormatConditions.Item[rule].Type;
                    if (type == 4 || type == 6) throw new InvalidOperationException("复制范围含条件格式数据条或图标集，不能固定为单元格外观。请在设置中选择仅保留基础格式，或缩小范围。");
                }
                total += (long)(double)cells.CountLarge;
                if (total > options.AppearanceLimit) throw new InvalidOperationException("条件格式范围超过外观固定上限 " + options.AppearanceLimit.ToString("N0") + "。请缩小范围，或在原样复制设置中调整上限/选择仅基础格式。");
                int index = 0;
                foreach (dynamic cell in cells.Cells)
                {
                    if (++index % 64 == 0) report("读取条件格式外观", index);
                    dynamic display = cell.DisplayFormat;
                    var a = new Appearance { Color = display.Interior.Color, Pattern = display.Interior.Pattern, PatternColor = display.Interior.PatternColor,
                        FontColor = display.Font.Color, Bold = display.Font.Bold, Italic = display.Font.Italic, Underline = display.Font.Underline, Strike = display.Font.Strikethrough, NumberFormat = display.NumberFormat };
                    for (int b = 0; b < 4; b++) { dynamic border = display.Borders.Item[7 + b]; a.Lines[b] = border.LineStyle; a.Weights[b] = border.Weight; a.Colors[b] = border.Color; }
                    string key = a.Key;
                    if (!groups.TryGetValue(key, out var existing)) groups.Add(key, existing = a);
                    existing.Addresses.Add((string)cell.Address);
                }
            }
            return groups.Values.ToList();
        }
        private static void ApplyCopyAppearance(dynamic target, List<Appearance> appearances, CopyOptions options, Action<string, long> report)
        {
            foreach (var a in appearances)
            for (int start = 0; start < a.Addresses.Count;)
            {
                report("固定条件格式外观", start);
                string address = a.Addresses[start++];
                while (start < a.Addresses.Count && address.Length + a.Addresses[start].Length < 220) address += "," + a.Addresses[start++];
                dynamic range = target.Range[address];
                range.Interior.Color = a.Color; range.Interior.PatternColor = a.PatternColor; range.Interior.Pattern = a.Pattern;
                range.Font.Color = a.FontColor; range.Font.Bold = a.Bold; range.Font.Italic = a.Italic; range.Font.Underline = a.Underline; range.Font.Strikethrough = a.Strike;
                if (options.NumberFormats) range.NumberFormat = a.NumberFormat;
                // Each cell needs its own perimeter; a multi-cell Range's outer border is not equivalent.
                foreach (dynamic cell in range.Cells) for (int b = 0; b < 4; b++)
                { dynamic border = cell.Borders.Item[7 + b]; border.Weight = a.Weights[b]; border.Color = a.Colors[b]; border.LineStyle = a.Lines[b]; }
            }
        }
        private static void CopyTheme(dynamic sourceBook, dynamic targetBook)
        {
            string id = Guid.NewGuid().ToString("N");
            string colorPath = Path.Combine(Path.GetTempPath(), "GridPilot-color-" + id + ".xml");
            string fontPath = Path.Combine(Path.GetTempPath(), "GridPilot-font-" + id + ".xml");
            string snapshotPath = Path.Combine(Path.GetTempPath(), "GridPilot-theme-" + id + ".tmp");
            try
            {
                try
                {
                    sourceBook.Theme.ThemeColorScheme.Save(colorPath);
                    sourceBook.Theme.ThemeFontScheme.Save(fontPath);
                }
                catch (COMException)
                {
                    // Excel blocks Theme access on protected sheets. Read the exact theme from
                    // a temporary SaveCopyAs snapshot, never unprotect/save over the source file.
                    // Binary .xls and encrypted packages cannot be read this way and fail closed.
                    sourceBook.SaveCopyAs(snapshotPath);
                    try
                    {
                        using (var package = Package.Open(snapshotPath, FileMode.Open, FileAccess.Read))
                        using (var stream = package.GetPart(new Uri("/xl/theme/theme1.xml", UriKind.Relative)).GetStream())
                        {
                            var theme = XDocument.Load(stream); XNamespace ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
                            theme.Descendants(ns + "clrScheme").First().Save(colorPath);
                            theme.Descendants(ns + "fontScheme").First().Save(fontPath);
                        }
                    }
                    catch (Exception ex) when (ex is FileFormatException || ex is InvalidOperationException || ex is IOException)
                    { throw new InvalidOperationException("无法读取此受保护文件的原始主题。请关闭设置中的字体/填充/边框复制，或使用未加密的 xlsx/xlsm/xlsb 副本。未修改源文件。", ex); }
                }
                targetBook.Theme.ThemeColorScheme.Load(colorPath);
                targetBook.Theme.ThemeFontScheme.Load(fontPath);
                targetBook.Styles.Item["Normal"].Font.Name = sourceBook.Styles.Item["Normal"].Font.Name;
                targetBook.Styles.Item["Normal"].Font.Size = sourceBook.Styles.Item["Normal"].Font.Size;
            }
            finally
            {
                if (File.Exists(colorPath)) File.Delete(colorPath);
                if (File.Exists(fontPath)) File.Delete(fontPath);
                if (File.Exists(snapshotPath)) File.Delete(snapshotPath);
            }
        }
    }
}
