using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace SheetPilot.Core
{
    [DataContract]
    public sealed class CopyOptions
    {
        [DataMember] public int Scope { get; set; }
        [DataMember] public bool EmptySelectionUsesUsedRange { get; set; }
        [DataMember] public bool TrimWholeRowsColumns { get; set; } = true;
        [DataMember] public bool VisibleCellsOnly { get; set; }
        [DataMember] public bool CellFormats { get; set; } = true;
        [DataMember] public bool NumberFormats { get; set; } = true;
        [DataMember] public bool RowHeights { get; set; } = true;
        [DataMember] public bool ColumnWidths { get; set; } = true;
        [DataMember] public bool HiddenState { get; set; } = true;
        [DataMember] public bool MergedCells { get; set; } = true;
        [DataMember] public bool Notes { get; set; } = true;
        [DataMember] public bool ViewSettings { get; set; } = true;
        [DataMember] public bool KeepSourceActive { get; set; }
        [DataMember] public int ConditionalFormatting { get; set; }
        [DataMember] public int FormulaFallback { get; set; }
        [DataMember] public int CellLimit { get; set; } = 1000000;
        [DataMember] public int AppearanceLimit { get; set; } = 10000;
        public CopyOptions Clone() => (CopyOptions)MemberwiseClone();
        [OnDeserializing] private void Initialize(StreamingContext c)
        { TrimWholeRowsColumns = CellFormats = NumberFormats = RowHeights = ColumnWidths = HiddenState = MergedCells = Notes = ViewSettings = true; CellLimit = 1000000; AppearanceLimit = 10000; }
        public void Normalize()
        { Scope = Math.Max(0, Math.Min(1, Scope)); ConditionalFormatting = Math.Max(0, Math.Min(1, ConditionalFormatting)); FormulaFallback = Math.Max(0, Math.Min(1, FormulaFallback)); CellLimit = Math.Max(100, Math.Min(5000000, CellLimit)); AppearanceLimit = Math.Max(100, Math.Min(100000, AppearanceLimit)); }
    }

    public struct CellRect
    {
        public readonly int Row, Column, Rows, Columns;
        public int Bottom => Row + Rows - 1;
        public int Right => Column + Columns - 1;
        public long Count => (long)Rows * Columns;
        public CellRect(int row, int column, int rows, int columns) { Row = row; Column = column; Rows = rows; Columns = columns; }
        public bool Contains(int row, int column) => row >= Row && row <= Bottom && column >= Column && column <= Right;
        public bool Contains(CellRect other) => Contains(other.Row, other.Column) && Contains(other.Bottom, other.Right);
        public string Address => ColumnName(Column) + Row + ":" + ColumnName(Right) + Bottom;
        public static string ColumnName(int column) { string value = ""; while (column > 0) { column--; value = (char)('A' + column % 26) + value; column /= 26; } return value; }
        public static bool Covered(CellRect target, IReadOnlyList<CellRect> areas)
        {
            var remaining = new List<CellRect> { target };
            foreach (var area in areas)
            {
                var next = new List<CellRect>();
                foreach (var r in remaining)
                {
                    int top = Math.Max(r.Row, area.Row), left = Math.Max(r.Column, area.Column), bottom = Math.Min(r.Bottom, area.Bottom), right = Math.Min(r.Right, area.Right);
                    if (top > bottom || left > right) { next.Add(r); continue; }
                    if (r.Row < top) next.Add(new CellRect(r.Row, r.Column, top - r.Row, r.Columns));
                    if (r.Bottom > bottom) next.Add(new CellRect(bottom + 1, r.Column, r.Bottom - bottom, r.Columns));
                    if (r.Column < left) next.Add(new CellRect(top, r.Column, bottom - top + 1, left - r.Column));
                    if (r.Right > right) next.Add(new CellRect(top, right + 1, bottom - top + 1, r.Right - right));
                }
                remaining = next; if (remaining.Count == 0) return true;
            }
            return false;
        }
    }

    // Conservative lexical analysis: no Evaluate, DirectPrecedents, regex substitution of string literals,
    // or discovery of dependencies through the active workbook. Unknown constructs become cached values.
    public static class InternalFormula
    {
        private static readonly Regex Reference = new Regex(@"\G(?:(?:'(?<quoted>(?:[^']|'')+)'|(?<sheet>[\p{L}_][\p{L}\p{N}_.]*))!)?(?<first>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})(?::(?<last>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}))?", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
        private static readonly Regex Number = new Regex(@"\G(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(100));
        private static readonly HashSet<string> Functions = new HashSet<string>((
            "SUM SUMIF SUMIFS SUMPRODUCT SUBTOTAL AGGREGATE PRODUCT AVERAGE AVERAGEIF AVERAGEIFS MIN MAX MEDIAN COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS " +
            "IF IFERROR IFNA AND OR NOT XOR TRUE FALSE ABS SIGN INT TRUNC MOD ROUND ROUNDUP ROUNDDOWN MROUND CEILING FLOOR POWER SQRT EXP LN LOG LOG10 " +
            "LEFT RIGHT MID LEN LOWER UPPER PROPER TRIM CLEAN SUBSTITUTE REPLACE FIND SEARCH EXACT CONCATENATE TEXT VALUE NUMBERVALUE CHAR CODE UNICHAR UNICODE " +
            "DATE DATEVALUE DAY MONTH YEAR DAYS EDATE EOMONTH TIME TIMEVALUE HOUR MINUTE SECOND WEEKDAY WEEKNUM NETWORKDAYS WORKDAY DATEDIF YEARFRAC " +
            "ISBLANK ISNUMBER ISTEXT ISERROR ISERR ISNA ISLOGICAL ISEVEN ISODD N T NA MATCH INDEX LOOKUP VLOOKUP HLOOKUP ROW COLUMN ROWS COLUMNS " +
            "STDEV STDEVP STDEV.S STDEV.P VAR VARP VAR.S VAR.P LARGE SMALL RANK RANK.EQ RANK.AVG PERCENTILE QUARTILE CHOOSE").Split(' '), StringComparer.OrdinalIgnoreCase);

        public static bool CanKeep(string formula, string sourceSheet, IReadOnlyList<CellRect> areas, out string reason)
        {
            reason = null;
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') { reason = "非普通公式"; return false; }
            try
            {
                for (int i = 1; i < formula.Length;)
                {
                    char ch = formula[i];
                    if (char.IsWhiteSpace(ch) || "+-*/^&=<>%,;(){}".IndexOf(ch) >= 0) { i++; continue; }
                    if (ch == '"')
                    {
                        i++; bool closed = false;
                        while (i < formula.Length) { if (formula[i++] != '"') continue; if (i < formula.Length && formula[i] == '"') { i++; continue; } closed = true; break; }
                        if (!closed) { reason = "未闭合字符串"; return false; } continue;
                    }
                    // Recognize a function before cell tokens, e.g. LOG10(100).
                    if (char.IsLetter(ch) || ch == '_')
                    {
                        int end = i; while (end < formula.Length && (char.IsLetterOrDigit(formula[end]) || formula[end] == '_' || formula[end] == '.')) end++;
                        int next = end; while (next < formula.Length && char.IsWhiteSpace(formula[next])) next++;
                        string word = formula.Substring(i, end - i);
                        if (next < formula.Length && formula[next] == '(')
                        {
                            if (!Functions.Contains(word)) { reason = "动态引用、未知或依赖环境的函数：" + word; return false; }
                            i = next + 1; continue;
                        }
                        if ((word.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || word.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) && (next == formula.Length || formula[next] != '!')) { i = end; continue; }
                    }
                    var match = Reference.Match(formula, i);
                    if (match.Success)
                    {
                        int end = i + match.Length;
                        if (end < formula.Length && (char.IsLetterOrDigit(formula[end]) || "_.$![:#".IndexOf(formula[end]) >= 0)) { reason = "命名区域、结构化或动态引用"; return false; }
                        string sheet = match.Groups["quoted"].Success ? match.Groups["quoted"].Value.Replace("''", "'") : match.Groups["sheet"].Value;
                        if (sheet.Length > 0 && !string.Equals(sheet, sourceSheet, StringComparison.OrdinalIgnoreCase)) { reason = "跨工作表引用"; return false; }
                        if (!ParseCell(match.Groups["first"].Value, out int r1, out int c1) || !ParseCell(match.Groups["last"].Success ? match.Groups["last"].Value : match.Groups["first"].Value, out int r2, out int c2)) { reason = "超出工作表边界"; return false; }
                        var rect = new CellRect(Math.Min(r1, r2), Math.Min(c1, c2), Math.Abs(r2 - r1) + 1, Math.Abs(c2 - c1) + 1);
                        if (!CellRect.Covered(rect, areas)) { reason = "引用复制范围之外的单元格"; return false; }
                        i = end; continue;
                    }
                    var number = Number.Match(formula, i); if (number.Success) { i += number.Length; continue; }
                    reason = "外部引用、名称或未支持的公式语法"; return false;
                }
                return true;
            }
            catch (RegexMatchTimeoutException) { reason = "公式分析超时"; return false; }
        }
        private static bool ParseCell(string text, out int row, out int column)
        {
            text = text.Replace("$", ""); column = 0; int i = 0;
            while (i < text.Length && char.IsLetter(text[i])) column = column * 26 + char.ToUpperInvariant(text[i++]) - 'A' + 1;
            return int.TryParse(text.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out row) && column <= 16384 && row <= 1048576;
        }
    }
}
