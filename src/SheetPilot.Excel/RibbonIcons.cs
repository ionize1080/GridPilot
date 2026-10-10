using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace SheetPilot.Excel
{
    public static class RibbonIcons
    {
        private static readonly Dictionary<string, Bitmap> cache = new Dictionary<string, Bitmap>();
        public static Bitmap Get(string id, IntPtr window)
        {
            int dpi = 96; try { uint value = GetDpiForWindow(window); if (value > 0) dpi = (int)value; } catch (EntryPointNotFoundException) { }
            int size = Math.Max(32, Math.Min(96, (int)Math.Round(32 * dpi / 96.0)));
            string key = id + ":" + size;
            if (!cache.TryGetValue(key, out var bitmap)) cache[key] = bitmap = Draw(id, size);
            return bitmap;
        }
        public static Bitmap Draw(string id, int size)
        {
            var image = new Bitmap(size, size, PixelFormat.Format32bppArgb); image.SetResolution(96, 96);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(Color.FromArgb(40, 79, 72), 2) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            using (var accent = new SolidBrush(Color.FromArgb(38, 150, 122)))
            using (var faint = new SolidBrush(Color.FromArgb(160, 214, 196)))
            {
                g.Clear(Color.Transparent); g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.HighQuality; g.ScaleTransform(size / 32f, size / 32f);
                if (id.Contains("Refresh")) { g.DrawArc(pen, 5, 5, 22, 22, 40, 285); g.FillPolygon(accent, new[] { new Point(24, 3), new Point(29, 11), new Point(20, 11) }); }
                else if (id.Contains("Focus")) { g.FillRectangle(faint, 2, 12, 28, 8); g.FillRectangle(faint, 12, 2, 8, 28); g.DrawRectangle(pen, 3, 3, 26, 26); g.DrawLine(pen, 3, 12, 29, 12); g.DrawLine(pen, 3, 20, 29, 20); g.DrawLine(pen, 12, 3, 12, 29); g.DrawLine(pen, 20, 3, 20, 29); g.FillRectangle(accent, 13, 13, 6, 6); }
                else if (id.Contains("Rename")) { g.DrawRectangle(pen, 4, 3, 20, 25); g.DrawLine(pen, 8, 9, 19, 9); g.DrawLine(pen, 8, 14, 16, 14); g.FillPolygon(accent, new[] { new Point(13, 25), new Point(15, 19), new Point(25, 9), new Point(29, 13), new Point(19, 23) }); }
                else if (id.Contains("Copy")) { g.FillRectangle(faint, 3, 3, 18, 22); g.DrawRectangle(pen, 3, 3, 18, 22); g.FillRectangle(Brushes.White, 10, 9, 18, 21); g.DrawRectangle(pen, 10, 9, 18, 21); using (var font = new Font("Segoe UI", id.Contains("Internal") ? 12 : 10, FontStyle.Bold, GraphicsUnit.Pixel)) g.DrawString(id.Contains("Internal") ? "fx" : "123", font, accent, 11, 14); }
                else if (id.Contains("Text") || id.Contains("Number")) { g.DrawRectangle(pen, 3, 3, 26, 26); g.FillPolygon(accent, new[] { new Point(4, 4), new Point(11, 4), new Point(4, 11) }); using (var font = new Font("Segoe UI", id.Contains("Text") ? 17 : 13, FontStyle.Bold, GraphicsUnit.Pixel)) using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center }) g.DrawString(id.Contains("Text") ? "A" : "12", font, accent, new RectangleF(3, 4, 26, 26), format); }
                else if (id.Contains("Undo")) { g.DrawArc(pen, 7, 9, 20, 18, 200, 260); g.FillPolygon(accent, new[] { new Point(3, 13), new Point(12, 5), new Point(12, 20) }); }
                else if (id.Contains("Settings")) { g.FillEllipse(faint, 6, 6, 20, 20); for (int i=0;i<8;i++) { double a = i*Math.PI/4; g.DrawLine(pen, 16+(float)Math.Cos(a)*10,16+(float)Math.Sin(a)*10,16+(float)Math.Cos(a)*14,16+(float)Math.Sin(a)*14); } g.DrawEllipse(pen, 7, 7, 18, 18); g.FillEllipse(accent, 12, 12, 8, 8); }
                else { g.DrawRectangle(pen, 5, 3, 21, 26); g.DrawLine(pen, 10, 3, 10, 29); g.FillPolygon(accent, new[] { new Point(17, 3), new Point(23, 3), new Point(23, 17), new Point(20, 14), new Point(17, 17) }); g.DrawLine(pen, 14, 22, 22, 22); }
            }
            return image;
        }
        public static void Release() { foreach (var image in cache.Values) image.Dispose(); cache.Clear(); }
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hwnd);
    }
}
