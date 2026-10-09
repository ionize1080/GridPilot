using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace SheetPilot.Core
{
    public static class NumericConversion
    {
        public static string AsText(double value, CultureInfo culture) => value.ToString("G15", culture);
        public static bool TryNumber(string source, ToolPreferences options, CultureInfo culture, out double value, out string reason)
        {
            value = 0; reason = "非数字文本";
            if (source == null) return false;
            string text = options.TrimNumericText ? source.Trim() : source;
            if (!options.TrimNumericText && source != source.Trim()) { reason = "含首尾空格"; return false; }
            bool percent = options.AllowPercent && text.EndsWith(culture.NumberFormat.PercentSymbol, StringComparison.Ordinal);
            if (percent) text = text.Substring(0, text.Length - culture.NumberFormat.PercentSymbol.Length).TrimEnd();
            string unsigned = text.TrimStart('+', '-');
            if (options.PreserveLeadingZeros && unsigned.Length > 1 && unsigned[0] == '0' && char.IsDigit(unsigned[1])) { reason = "保留前导零"; return false; }
            string mantissa = Regex.Split(unsigned, "[eE]")[0];
            string digits = new string(mantissa.Where(c => c >= '0' && c <= '9').ToArray()).TrimStart('0');
            if (options.ProtectPrecision && digits.Length > 15) { reason = "超过15位有效数字"; return false; }
            string group = culture.NumberFormat.NumberGroupSeparator;
            if (!string.IsNullOrEmpty(group) && mantissa.Contains(group))
            {
                if (!options.AllowThousands) { reason = "千位分隔符未启用"; return false; }
                string integer = mantissa.Split(new[] { culture.NumberFormat.NumberDecimalSeparator }, StringSplitOptions.None)[0];
                string[] groups = integer.Split(new[] { group }, StringSplitOptions.None);
                int[] sizes = culture.NumberFormat.NumberGroupSizes;
                for (int i = groups.Length - 1, level = 0; i > 0; i--, level++)
                    if (groups[i].Length != sizes[Math.Min(level, sizes.Length - 1)] || !groups[i].All(char.IsDigit)) { reason = "千位分组不完整"; return false; }
                int firstSize = sizes[Math.Min(groups.Length - 2, sizes.Length - 1)];
                if (groups[0].Length < 1 || groups[0].Length > firstSize || !groups[0].All(char.IsDigit)) { reason = "千位分组不完整"; return false; }
            }
            var style = NumberStyles.Float | (options.AllowThousands ? NumberStyles.AllowThousands : NumberStyles.None);
            if (!double.TryParse(text, style, culture, out value) || double.IsInfinity(value) || double.IsNaN(value)) return false;
            if (percent) value /= 100;
            reason = null; return true;
        }
    }
}
