using System;
using System.Windows;
using System.Windows.Media;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    internal static class ViewTheme
    {
        public static void Apply(ResourceDictionary resources, NavigationSettings settings)
        {
            var background = Parse(settings.BackgroundColor); var text = Parse(settings.TextColor); var accent = Parse(settings.AccentColor);
            Put(resources, "CanvasBrush", background); Put(resources, "SurfaceBrush", Parse(settings.SurfaceColor));
            Put(resources, "TextBrush", text); Put(resources, "MutedBrush", Parse(settings.MutedTextColor)); Put(resources, "AccentBrush", accent);
            Put(resources, "BorderBrush", Blend(background, text, .22)); Put(resources, "HoverBrush", Blend(background, accent, .10));
            Put(resources, "SelectedBrush", Blend(background, accent, .20));
            double luminance = .2126 * Linear(accent.R) + .7152 * Linear(accent.G) + .0722 * Linear(accent.B);
            Put(resources, "AccentTextBrush", luminance > .179 ? Colors.Black : Colors.White);
        }
        private static double Linear(byte b) { double n = b / 255.0; return n <= .04045 ? n / 12.92 : Math.Pow((n + .055) / 1.055, 2.4); }
        private static Color Parse(string s) => (Color)ColorConverter.ConvertFromString(s);
        private static Color Blend(Color a, Color b, double ratio) => Color.FromRgb((byte)(a.R + (b.R-a.R)*ratio), (byte)(a.G + (b.G-a.G)*ratio), (byte)(a.B + (b.B-a.B)*ratio));
        private static void Put(ResourceDictionary r, string key, Color color) { var brush = new SolidColorBrush(color); brush.Freeze(); r[key] = brush; }
    }
}
