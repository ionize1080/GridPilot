using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class ExcelGateway
    {
        private Func<string> undo;
        private ExcelEvents undoEvents;
        private bool undoInvalidated;
        public void ReleaseUndoEvents(){undoEvents?.Dispose();undoEvents=null;undo=null;}
        private void WatchUndo(){if(undoEvents==null){undoEvents=new ExcelEvents(application());undoEvents.Change(()=>{if(undo!=null)undoInvalidated=true;undo=null;});}}
        public bool CanUndo => undo != null;
        public string UndoLast()
        {
            CheckThread(); if (undo == null) {if(undoInvalidated)throw new InvalidOperationException("工作表已发生外部编辑，上次撤销已失效。");return "没有可撤销的 GridPilot 操作。";}
            var operation = undo; string result = operation(); undo = null; return result;
        }
        private object FindBook(string key)
        {
            dynamic app = application(), books = app.Workbooks;
            for (int i = 1; i <= (int)books.Count; i++) { dynamic book = books.Item[i]; if (WorkbookKey(book) == key) return book; }
            throw new InvalidOperationException("工作簿已关闭或另存为，请刷新预览。");
        }
        private object FindPage(dynamic book, string key)
        {
            dynamic pages = book.Sheets;
            for (int i = 1; i <= (int)pages.Count; i++) { dynamic page = pages.Item[i]; if (SheetKey(page) == key) return page; }
            throw new InvalidOperationException("页面标识已改变，请刷新预览。");
        }
        public string ApplyRename(IReadOnlyList<RenameEntry> plan, bool remember = true)
        {
            CheckThread();
            var changed = plan.Where(p => p.Changed).ToArray();
            if (changed.Length == 0) return "没有需要重命名的页面。";
            if (changed.Any(p => p.Error != null || RenamePlanner.ValidateName(p.NewName) != null)) throw new InvalidOperationException("请先解决预览中的名称错误。");
            var live = Capture(true).Sheets;
            foreach (var book in plan.GroupBy(p => p.Source.WorkbookKey))
            {
                var expected = book.ToDictionary(p => p.Source.Key);
                var current = live.Where(s => s.WorkbookKey == book.Key).ToArray();
                foreach (var entry in book)
                    if (!current.Any(s => s.Key == entry.Source.Key && s.Name == entry.Source.Name)) throw new InvalidOperationException("预览后页面发生变化，请刷新预览再应用。");
                var names = current.Select(s => expected.TryGetValue(s.Key, out var item) && item.Changed ? item.NewName : s.Name);
                if (names.Distinct(StringComparer.OrdinalIgnoreCase).Count() != current.Length) throw new InvalidOperationException("页面名称冲突，请刷新预览。");
            }
            var targets = new List<object>();
            foreach (var entry in changed)
            {
                dynamic book = FindBook(entry.Source.WorkbookKey);
                if ((bool)book.ReadOnly || (bool)book.ProtectStructure) throw new InvalidOperationException(entry.Source.WorkbookName + " 为只读或结构受保护，未应用任何修改。");
                targets.Add(FindPage(book, entry.Source.Key));
            }
            dynamic app = application(); bool events = app.EnableEvents, screen = app.ScreenUpdating;
            try
            {
                app.EnableEvents = false; app.ScreenUpdating = false;
                try
                {
                    for (int i = 0; i < targets.Count; i++) ((dynamic)targets[i]).Name = TemporaryName();
                    for (int i = 0; i < targets.Count; i++) ((dynamic)targets[i]).Name = changed[i].NewName;
                }
                catch (Exception ex)
                {
                    var failures = new List<string>();
                    for (int i = 0; i < targets.Count; i++) { try { ((dynamic)targets[i]).Name = TemporaryName(); } catch { failures.Add(changed[i].Source.Name); } }
                    for (int i = 0; i < targets.Count; i++) { try { ((dynamic)targets[i]).Name = changed[i].Source.Name; } catch { failures.Add(changed[i].Source.Name); } }
                    throw new InvalidOperationException(failures.Count == 0 ? "重命名未完成，已恢复原名称。" : "重命名失败且部分恢复失败，请检查：" + string.Join("、", failures.Distinct()), ex);
                }
            }
            finally { try { app.EnableEvents = events; } finally { app.ScreenUpdating = screen; } }
            if (remember)
            {
                var reverse = changed.Select(p => new RenameEntry(new SheetItem(p.Source.Key, p.NewName, p.Source.Index, p.Source.Visibility, p.Source.IsActive, p.Source.Kind, p.Source.WorkbookKey, p.Source.WorkbookName), true, p.Source.Name)).ToArray();
                undoInvalidated=false;undo = () => ApplyRename(reverse, false);
            }
            return "已重命名 " + changed.Length + " 个页面。可通过「撤销上次操作」恢复；未自动保存工作簿。";
        }
        private static string TemporaryName() => "_GP_" + Guid.NewGuid().ToString("N").Substring(0, 24);
        private sealed class CellChange
        {
            public string Book, Sheet, Address; public object Before, After; public bool IgnoreBefore;
            public int BeforeAlignment, AfterAlignment;
            public bool AlignmentChanged => BeforeAlignment != AfterAlignment;
        }
        public string ConvertSelection(bool toText, ToolPreferences options)
        {
            CheckThread(); options.Normalize();WatchUndo();
            dynamic app = application(), selection = app.Selection;
            if (selection == null || Microsoft.VisualBasic.Information.TypeName((object)selection).TrimStart('_') != "Range") throw new InvalidOperationException("请先选中单元格区域。");
            dynamic sheet = selection.Worksheet, book = sheet.Parent;
            if ((bool)book.ReadOnly) throw new InvalidOperationException("当前工作簿为只读，未转换。");
            object constants;
            try { constants = selection.SpecialCells(2, toText ? 1 : 2); }
            catch (COMException) { return "所选区域没有可转换的常量；公式、空白和错误值保持不变。"; }
            dynamic cells = app.Intersect(selection, constants);
            if (options.VisibleCellsOnly && cells != null)
            {
                try { cells = app.Intersect(cells, selection.SpecialCells(12)); }
                catch (COMException) { return "所选区域没有可转换的可见单元格。"; }
            }
            if (cells == null) return "所选区域没有可转换的单元格。";
            if (Convert.ToDouble(cells.CountLarge, CultureInfo.InvariantCulture) > options.ConversionLimit) throw new InvalidOperationException("可转换单元格超过设置的数量上限，请缩小范围或调整设置。");
            var culture = (CultureInfo)(options.DecimalCulture == 1 ? CultureInfo.InvariantCulture : CultureInfo.CurrentCulture).Clone();
            if (options.DecimalCulture == 0) { culture.NumberFormat.NumberDecimalSeparator = (string)app.DecimalSeparator; culture.NumberFormat.NumberGroupSeparator = (string)app.ThousandsSeparator; }
            var changes = new List<CellChange>(); var skipped = new Dictionary<string,int>();
            string bookKey = WorkbookKey(book), sheetKey = SheetKey(sheet);
            for (int areaIndex = 1; areaIndex <= (int)cells.Areas.Count; areaIndex++)
            {
            dynamic area = cells.Areas.Item[areaIndex];
            for (int i = 1; i <= (int)area.Cells.Count; i++)
            {
                dynamic cell = area.Cells.Item[i];
                if ((bool)cell.HasFormula) { Skip(skipped, "公式"); continue; }
                if ((bool)cell.MergeCells) { Skip(skipped, "合并单元格"); continue; }
                if ((bool)sheet.ProtectContents && (bool)cell.Locked) { Skip(skipped, "受保护单元格"); continue; }
                object before = cell.Value2, after;
                if (toText)
                {
                    if (!(before is double)) { Skip(skipped, "非数值"); continue; }
                    after = NumericConversion.AsText((double)before, culture);
                }
                else
                {
                    if (!(before is string text) || !NumericConversion.TryNumber(text, options, culture, out double number, out string reason))
                    { Skip(skipped, before is string s ? GetSkipReason(s, options, culture) : "非文本"); continue; }
                    after = number;
                }
                bool ignore = false; try { ignore = (bool)cell.Errors.Item[3].Ignore; } catch (COMException) { }
                int alignment = (int)cell.HorizontalAlignment;
                int afterAlignment = options.PreserveVisualAlignment && alignment == 1 ? (toText ? -4152 : -4131) : alignment;
                changes.Add(new CellChange { BeforeAlignment = alignment, AfterAlignment = afterAlignment, Book = bookKey, Sheet = sheetKey, Address = (string)cell.Address, Before = before, After = after, IgnoreBefore = ignore });
            }
            }
            if (changes.Count > 0)
            {
                WriteCells(changes,false,options.NumberAsTextHint);undoInvalidated=false;
                if(!(bool)app.EnableEvents){undo=null;undoInvalidated=true;}
                else undo=()=>{WriteCells(changes,true,false);return "已恢复 "+changes.Count+" 个单元格的内容。";};
            }
            string details = skipped.Count == 0 ? "" : "；跳过 " + string.Join("、", skipped.Select(p => p.Key + " " + p.Value));
            return "已" + (toText ? "转为文本 " : "转为数值 ") + changes.Count + " 个单元格" + details + (options.PreserveVisualAlignment ? "。保留视觉对齐与其他格式，公式保持不变。" : "。公式与格式属性保持不变。");
        }
        private static string GetSkipReason(string text, ToolPreferences options, CultureInfo culture) { NumericConversion.TryNumber(text, options, culture, out _, out string reason); return reason; }
        private static void Skip(Dictionary<string,int> counts, string reason) { counts.TryGetValue(reason, out int n); counts[reason] = n + 1; }
        private void WriteCells(IReadOnlyList<CellChange> changes, bool reverse, bool hints)
        {
            var targets = new List<object>();
            var sheets=new Dictionary<string,object>(StringComparer.Ordinal);
            foreach (var change in changes)
            {
                string key=change.Book+"|"+change.Sheet;
                if(!sheets.TryGetValue(key,out var target))
                {dynamic foundBook=FindBook(change.Book);if((bool)foundBook.ReadOnly)throw new InvalidOperationException("目标工作簿只读。");target=FindPage(foundBook,change.Sheet);sheets.Add(key,target);}
                dynamic sheet=target,cell=sheet.Range[change.Address];
                if (((bool)sheet.ProtectContents && (bool)cell.Locked) || (bool)cell.HasFormula || !Equals((object)cell.Value2, reverse ? change.After : change.Before) || (change.AlignmentChanged && (int)cell.HorizontalAlignment != (reverse ? change.AfterAlignment : change.BeforeAlignment)))
                    throw new InvalidOperationException("目标内容或保护状态已改变，未写入；请重新选择范围。");
                targets.Add((object)cell);
            }
            dynamic app = application(); bool events = app.EnableEvents, screen = app.ScreenUpdating;
            int done = 0;
            try
            {
                app.EnableEvents = false; app.ScreenUpdating = false;
                try
                {
                    for (int i = 0; i < changes.Count; i++)
                    {
                        dynamic cell = targets[i]; object value = reverse ? changes[i].Before : changes[i].After;
                        SetCellValue(cell, value); done = i + 1;
                        if (changes[i].AlignmentChanged) cell.HorizontalAlignment = reverse ? changes[i].BeforeAlignment : changes[i].AfterAlignment;
                        try { if (reverse) cell.Errors.Item[3].Ignore = changes[i].IgnoreBefore; else if (hints && value is string) cell.Errors.Item[3].Ignore = false; } catch (COMException) { }
                    }
                }
                catch (Exception ex)
                {
                    int failures = 0;
                    for (int i = done - 1; i >= 0; i--) { try { SetCellValue((dynamic)targets[i], reverse ? changes[i].After : changes[i].Before); if (changes[i].AlignmentChanged) ((dynamic)targets[i]).HorizontalAlignment = reverse ? changes[i].AfterAlignment : changes[i].BeforeAlignment; ((dynamic)targets[i]).Errors.Item[3].Ignore = changes[i].IgnoreBefore; } catch { failures++; } }
                    throw new InvalidOperationException(failures == 0 ? "转换失败，已恢复已写入的内容。" : "转换失败，部分内容无法恢复，请检查所选区域。", ex);
                }
            }
            finally { try { app.EnableEvents = events; } finally { app.ScreenUpdating = screen; } }
        }
        private static void SetCellValue(dynamic cell, object value)
        {
            // Apostrophe input stores a numeric string without changing NumberFormat or alignment.
            if (value is string text) cell.Formula = "'" + text; else cell.Value2 = value;
        }
    }
}
