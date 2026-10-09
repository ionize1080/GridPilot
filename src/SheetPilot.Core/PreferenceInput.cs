using System;
using System.Globalization;

namespace SheetPilot.Core
{
    public static class PreferenceInput
    {
        public static bool TryInteger(string text, int minimum, int maximum, out int value)
        {
            if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value)) return false;
            value = Math.Max(minimum, Math.Min(maximum, value));
            return true;
        }
        public static bool TryColor(string text, out string color)
        {
            color = null;
            if (text == null) return false;
            string digits = text.Trim();
            if (digits.StartsWith("#", StringComparison.Ordinal)) digits = digits.Substring(1);
            if (digits.Length == 3) digits = string.Concat(digits[0], digits[0], digits[1], digits[1], digits[2], digits[2]);
            if (digits.Length != 6 || !int.TryParse(digits, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out _)) return false;
            color = "#" + digits.ToUpperInvariant();
            return true;
        }
    }

    // Observed Office geometry must never be fed back as a new geometry command.
    public sealed class NativeLayoutGuard
    {
        private int depth;
        public bool IsObserving => depth != 0;
        public void Observe(Action update)
        {
            depth++;
            try { update(); }
            finally { depth--; }
        }
    }
}
