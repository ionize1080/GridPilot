using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace SheetPilot.Core
{
    public sealed class RenameEntry
    {
        public SheetItem Source { get; }
        public string NewName { get; internal set; }
        public string Error { get; internal set; }
        public bool Selected { get; }
        public bool Changed => Selected && Source.Name != NewName;
        public RenameEntry(SheetItem source, bool selected, string name) { Source = source; Selected = selected; NewName = name; }
    }
    public static class RenamePlanner
    {
        public static IReadOnlyList<RenameEntry> Build(IReadOnlyList<SheetItem> sources, RenameOptions options, ISet<string> excluded = null)
        {
            string find = options.Find ?? "";
            if (find.Length > 512 || (options.Replacement ?? "").Length > 512) throw new ArgumentException("查找或替换内容最多512个字符。");
            Regex regex = find.Length == 0 ? null : new Regex(options.Regex ? find : Regex.Escape(find), RegexOptions.CultureInvariant | (options.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase), TimeSpan.FromMilliseconds(40));
            var rows = new List<RenameEntry>(); var counts = new Dictionary<string, int>(); int global = 0;
            var watch = Stopwatch.StartNew();
            foreach (var sheet in sources)
            {
                if (watch.ElapsedMilliseconds > 1000) throw new ArgumentException("预览耗时过长，请简化表达式。");
                bool selected = (excluded == null || !excluded.Contains(Key(sheet))) && (options.IncludeHidden || sheet.CanNavigate) && (options.IncludeCharts || sheet.Kind != "图表页");
                if (selected && options.MatchesOnly && regex != null && !regex.IsMatch(sheet.Name)) selected = false;
                string name = sheet.Name;
                if (selected)
                {
                    if (regex != null)
                    {
                        int count = options.FirstOnly ? 1 : int.MaxValue;
                        name = options.Regex ? regex.Replace(name, options.Replacement ?? "", count) : regex.Replace(name, m => options.Replacement ?? "", count);
                    }
                    if (options.Trim) name = name.Trim();
                    if (options.LetterCase == 1) name = name.ToUpperInvariant(); else if (options.LetterCase == 2) name = name.ToLowerInvariant();
                    name = (options.Prefix ?? "") + name + (options.Suffix ?? "");
                    if (options.Numbering)
                    {
                        counts.TryGetValue(sheet.WorkbookKey, out int n); counts[sheet.WorkbookKey] = n + 1;
                        long number = options.Start + (long)(options.RestartPerBook ? n : global) * options.Step; global++;
                        string serial = number.ToString("D" + Math.Max(1, Math.Min(10, options.Digits)), CultureInfo.InvariantCulture);
                        name = options.NumberAtEnd ? name + (options.Separator ?? "") + serial : serial + (options.Separator ?? "") + name;
                    }
                    if (options.Truncate && name.Length > 31) name = SafeCut(name, 31);
                }
                var entry = new RenameEntry(sheet, selected, name);
                if (entry.Changed) entry.Error = ValidateName(name);
                rows.Add(entry);
            }
            foreach (var group in rows.GroupBy(r => r.Source.WorkbookKey))
                foreach (var collision in group.GroupBy(r => r.NewName, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
                    foreach (var row in collision.Where(r => r.Changed)) row.Error = "与同一工作簿中的页面重名";
            return rows.AsReadOnly();
        }
        public static string Key(SheetItem sheet) => sheet.WorkbookKey + "|" + sheet.Key;
        private static string SafeCut(string value, int length) => value.Substring(0, char.IsHighSurrogate(value[length - 1]) ? length - 1 : length);
        public static string ValidateName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "名称不能为空";
            if (name.Length > 31) return "名称超过31个字符";
            if (name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0 || name.Any(char.IsControl)) return "名称含不允许的字符";
            if (name.StartsWith("'", StringComparison.Ordinal) || name.EndsWith("'", StringComparison.Ordinal)) return "名称不能以单引号开头或结尾";
            if (string.Equals(name, "History", StringComparison.OrdinalIgnoreCase)) return "History 是 Excel 保留名称";
            return null;
        }
    }
}
