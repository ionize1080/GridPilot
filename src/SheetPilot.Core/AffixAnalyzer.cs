using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace SheetPilot.Core
{
    public static class AffixAnalyzer
    {
        private static readonly string[] Endings = { "股份有限公司", "有限责任公司", "集团有限公司", "有限公司", "支行", "分行", "工程项目", "建设项目", "改造项目", "项目", "工程", "中心", "事业部", "研究所", "基金", "合同", "Co., Ltd.", "Ltd.", "Inc.", "Corp." };
        // Reservoir sampling covers the entire column, including records far below the header.
        public static string[] Sample(IEnumerable<DataCell> cells, CancellationToken cancel)
        {
            var values = new List<string>(2000); var random = new Random(271828); int count = 0;
            foreach (var cell in cells)
            {
                cancel.ThrowIfCancellationRequested();
                if (!(cell.Value is string text) || string.IsNullOrWhiteSpace(text) || MaskPlanner.IsPlaceholder(text)) continue;
                count++; if (values.Count < 2000) values.Add(text);
                else { int index = random.Next(count); if (index < values.Count) values[index] = text; }
            }
            return values.ToArray();
        }
        public static void Apply(MaskRule rule, IEnumerable<string> input)
        {
            var values = input.Where(s => !string.IsNullOrWhiteSpace(s) && !MaskPlanner.IsPlaceholder(s)).ToArray();
            rule.Prefix = rule.Suffix = ""; rule.Digits = rule.Kind == MaskKind.Identifier ? 8 : 4; rule.ShortName = false;
            if (values.Length == 0) return;
            if (rule.Kind == MaskKind.Identifier)
            {
                // Last numeric run preserves structured prefixes such as HT-2026- and suffixes such as 号.
                var shapes = values.Select(v => Regex.Match(v, @"^(.*?)([0-9]+)([^0-9]*)$", RegexOptions.CultureInvariant))
                    .Where(m => m.Success).Select(m => new { Prefix = m.Groups[1].Value, Suffix = m.Groups[3].Value, Digits = m.Groups[2].Length })
                    .GroupBy(s => s).OrderByDescending(g => g.Count()).ThenBy(g => g.Key.Prefix, StringComparer.Ordinal).FirstOrDefault();
                if (shapes != null && shapes.Count() >= Math.Max(2, values.Length * .4) && shapes.Key.Digits <= 18)
                { rule.Prefix = shapes.Key.Prefix; rule.Suffix = shapes.Key.Suffix; rule.Digits = shapes.Key.Digits; rule.Reason += "；编号结构覆盖 " + (100 * shapes.Count() / values.Length) + "%"; }
                return;
            }
            if (rule.Kind != MaskKind.ChineseName && rule.Kind != MaskKind.EnglishName) return;
            var suffix = Endings.Select(s => new { Text = s, Hits = values.Count(v => v.EndsWith(s, StringComparison.OrdinalIgnoreCase) && v.Length > s.Length) })
                .Where(s => s.Hits >= 2 && s.Hits >= values.Length * .28)
                .OrderByDescending(s => s.Hits * (1 + .35 * s.Text.Length)).ThenByDescending(s => s.Text.Length).FirstOrDefault();
            if (suffix != null) { rule.Suffix = suffix.Text; rule.Reason += "；词尾「" + suffix.Text + "」覆盖 " + (100 * suffix.Hits / values.Length) + "%"; }
            else
            {
                // Require diverse stems before suggesting an unknown common tail; duplicated names do not count as evidence.
                var distinct = values.Distinct(StringComparer.Ordinal).ToArray();
                var tail = distinct.SelectMany(v => Enumerable.Range(2, 3).Where(n => v.Length > n + 2).Select(n => v.Substring(v.Length - n)))
                    .Where(t => t.All(c => c >= '\u4e00' && c <= '\u9fff')).GroupBy(t => t)
                    .Where(g => g.Count() >= 3 && g.Count() >= distinct.Length * .6)
                    .OrderByDescending(g => g.Count() * (1 + .15 * g.Key.Length)).ThenBy(g => g.Key, StringComparer.Ordinal).FirstOrDefault();
                if (tail != null) { rule.Suffix = tail.Key; rule.Reason += "；高频词尾建议，请确认是否含敏感信息"; }
            }
            rule.ShortName = rule.Kind == MaskKind.ChineseName && rule.Suffix.Length == 0 && values.Average(v => v.Length) <= 4;
            if (rule.ShortName) rule.Digits = 0;
            // Pure text prefixes can disclose location or organization and remain opt-in.
            var tag = values.Select(v => Regex.Match(v, @"^(?:【[^】]{1,12}】|（[^）]{1,12}）|\([^)]{1,12}\))").Value)
                .Where(v => v.Length > 0).GroupBy(v => v).OrderByDescending(g => g.Count()).FirstOrDefault();
            if (tag != null && tag.Count() >= Math.Max(2, values.Length * .6)) { rule.Prefix = tag.Key; rule.Reason += "；公共标签建议，请确认"; }
        }
    }
}
