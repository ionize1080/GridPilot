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
        private sealed class PictureConverter : System.Windows.Forms.AxHost
        { private PictureConverter():base(""){} public static object Convert(Bitmap image)=>GetIPictureDispFromPicture(image); }
        public static object Picture(string id,IntPtr window,int theme)=>PictureConverter.Convert(Get(id,window,theme));
        private static readonly Dictionary<string, Bitmap> cache = new Dictionary<string, Bitmap>();
        public static Bitmap Get(string id, IntPtr window, int theme=0)
        {
            int dpi = 96; try { uint value = GetDpiForWindow(window); if (value > 0) dpi = (int)value; } catch (EntryPointNotFoundException) { }
            int size = Math.Max(32, Math.Min(96, (int)Math.Round(32 * dpi / 96.0)));
            string key = id + ":" + size + ":" + theme;
            if (!cache.TryGetValue(key, out var bitmap)) cache[key] = bitmap = Draw(id, size, theme);
            return bitmap;
        }
        public static Bitmap Draw(string id, int size, int theme=0)
        {
            var image = new Bitmap(size*4, size*4, PixelFormat.Format32bppArgb); image.SetResolution(96, 96);
            using (var g = Graphics.FromImage(image))
            using (var pen = new Pen(theme==1?Color.FromArgb(43,67,112):theme==2?Color.FromArgb(49,56,67):Color.FromArgb(38,83,74), 1.65f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            using (var accent = new SolidBrush(theme==1?Color.FromArgb(70,113,214):theme==2?Color.FromArgb(81,94,111):Color.FromArgb(35,153,125)))
            using (var faint = new SolidBrush(theme==1?Color.FromArgb(211,223,249):theme==2?Color.FromArgb(216,222,228):Color.FromArgb(201,234,224)))
            {
                g.Clear(Color.Transparent); g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.HighQuality; g.ScaleTransform(size*4 / 32f, size*4 / 32f);
                if (id == "GridPilot_AI")
                {
                    // Open grid + a four-point spark; shared palette, legible at Ribbon size.
                    using(var outline=new GraphicsPath())
                    {
                        outline.AddArc(3,7,6,6,180,90);outline.AddLine(6,7,17,7);outline.AddLine(17,7,17,11);
                        outline.AddLine(17,11,25,11);outline.AddLine(25,11,25,25);outline.AddArc(19,22,6,6,0,90);
                        outline.AddLine(22,28,6,28);outline.AddArc(3,22,6,6,90,90);outline.CloseFigure();
                        g.FillPath(faint,outline);g.DrawPath(pen,outline);
                    }
                    g.DrawLine(pen,4,17,24,17);g.DrawLine(pen,11,8,11,27);g.DrawLine(pen,4,22,24,22);
                    g.FillPolygon(accent,new[]{new Point(24,1),new Point(26,6),new Point(31,8),new Point(26,10),new Point(24,15),new Point(22,10),new Point(17,8),new Point(22,6)});
                }
                else if (id.Contains("Refresh")) { g.DrawArc(pen, 5, 5, 22, 22, 40, 285); g.FillPolygon(accent, new[] { new Point(24, 3), new Point(29, 11), new Point(20, 11) }); }
                else if (id.Contains("Focus")) { g.FillRectangle(faint, 2, 12, 28, 8); g.FillRectangle(faint, 12, 2, 8, 28); g.DrawRectangle(pen, 3, 3, 26, 26); g.DrawLine(pen, 3, 12, 29, 12); g.DrawLine(pen, 3, 20, 29, 20); g.DrawLine(pen, 12, 3, 12, 29); g.DrawLine(pen, 20, 3, 20, 29); g.FillRectangle(accent, 13, 13, 6, 6); }
                else if (id.Contains("Rename")) { g.DrawRectangle(pen, 4, 3, 20, 25); g.DrawLine(pen, 8, 9, 19, 9); g.DrawLine(pen, 8, 14, 16, 14); g.FillPolygon(accent, new[] { new Point(13, 25), new Point(15, 19), new Point(25, 9), new Point(29, 13), new Point(19, 23) }); }
                else if (id.Contains("Database")) {g.FillRectangle(faint,5,8,22,17);g.FillEllipse(faint,5,3,22,10);g.DrawEllipse(pen,5,3,22,10);g.DrawLine(pen,5,8,5,25);g.DrawLine(pen,27,8,27,25);g.DrawArc(pen,5,12,22,10,0,180);g.DrawArc(pen,5,20,22,10,0,180);g.FillEllipse(accent,20,23,8,8);}
                else if (id.Contains("Sort")) {g.DrawLine(pen,6,4,6,27);g.DrawLines(pen,new[]{new Point(2,23),new Point(6,27),new Point(10,23)});g.FillRectangle(accent,14,5,7,4);g.FillRectangle(accent,14,14,11,4);g.FillRectangle(accent,14,23,15,4);}
                else if (id.Contains("Reconcile")) {g.FillRectangle(faint,3,4,11,22);g.DrawRectangle(pen,3,4,11,22);g.DrawRectangle(pen,18,7,11,22);g.DrawLine(pen,6,10,11,10);g.DrawLine(pen,6,15,10,15);g.DrawLine(pen,21,13,26,13);g.DrawLines(pen,new[]{new Point(10,23),new Point(14,27),new Point(23,18)});}
                else if (id.Contains("SimpleMask")) {g.FillRectangle(faint,3,5,26,22);g.DrawRectangle(pen,3,5,26,22);g.DrawLine(pen,3,12,29,12);DrawLabel(g,"A",accent,new RectangleF(5,15,8,9),10);g.DrawLine(pen,14,20,19,20);g.DrawLines(pen,new[]{new Point(17,18),new Point(19,20),new Point(17,22)});g.FillEllipse(accent,23,18,4,4);}
                else if (id.Contains("Mask")) { g.FillPolygon(faint,new[]{new Point(16,2),new Point(28,7),new Point(26,21),new Point(16,30),new Point(6,21),new Point(4,7)});g.DrawArc(pen,11,8,10,13,180,180);g.FillRectangle(accent,9,14,14,10);g.FillEllipse(Brushes.White,14,17,4,4); }
                else if (id.Contains("Segment")) { g.DrawRectangle(pen,3,4,26,24);g.DrawLine(pen,16,5,16,27);DrawLabel(g,"中",accent,new RectangleF(5,8,9,16),11,"Microsoft YaHei");DrawLabel(g,"A",accent,new RectangleF(18,8,9,16),12); }
                else if (id.Contains("Copy")) { g.FillRectangle(faint, 3, 3, 18, 22); g.DrawRectangle(pen, 3, 3, 18, 22); g.FillRectangle(Brushes.White, 10, 9, 18, 21); g.DrawRectangle(pen, 10, 9, 18, 21); DrawLabel(g,id.Contains("Internal")?"fx":"123",accent,new RectangleF(12,12,14,15),id.Contains("Internal")?12:10); }
                else if (id.Contains("Text") || id.Contains("Number")) { g.DrawRectangle(pen, 3, 3, 26, 26); g.FillPolygon(accent, new[] { new Point(4, 4), new Point(11, 4), new Point(4, 11) }); DrawLabel(g,id.Contains("Text")?"A":"12",accent,new RectangleF(5,6,22,21),id.Contains("Text")?17:13); }
                else if (id.Contains("Undo")) { g.DrawArc(pen, 7, 9, 20, 18, 200, 260); g.FillPolygon(accent, new[] { new Point(3, 13), new Point(12, 5), new Point(12, 20) }); }
                else if (id.Contains("Settings")) { g.FillEllipse(faint, 6, 6, 20, 20); for (int i=0;i<8;i++) { double a = i*Math.PI/4; g.DrawLine(pen, 16+(float)Math.Cos(a)*10,16+(float)Math.Sin(a)*10,16+(float)Math.Cos(a)*14,16+(float)Math.Sin(a)*14); } g.DrawEllipse(pen, 7, 7, 18, 18); g.FillEllipse(accent, 12, 12, 8, 8); }
                else { g.DrawRectangle(pen, 5, 3, 21, 26); g.DrawLine(pen, 10, 3, 10, 29); g.FillPolygon(accent, new[] { new Point(17, 3), new Point(23, 3), new Point(23, 17), new Point(20, 14), new Point(17, 17) }); g.DrawLine(pen, 14, 22, 22, 22); }
            }
            var output=new Bitmap(size,size,PixelFormat.Format32bppArgb);
            using(var scaled=Graphics.FromImage(output)){scaled.InterpolationMode=InterpolationMode.HighQualityBicubic;scaled.PixelOffsetMode=PixelOffsetMode.HighQuality;scaled.DrawImage(image,new Rectangle(0,0,size,size));}
            image.Dispose();return output;
        }
        private static void DrawLabel(Graphics graphics,string text,Brush brush,RectangleF box,float size,string family="Segoe UI")
        {
            // Center the actual glyph outlines, not font side bearings or DrawString layout padding.
            using(var font=new FontFamily(family))using(var glyph=new GraphicsPath())using(var matrix=new Matrix())
            {
                glyph.AddString(text,font,(int)FontStyle.Bold,size,PointF.Empty,StringFormat.GenericTypographic);
                var bounds=glyph.GetBounds();float scale=Math.Min(1,Math.Min(box.Width/bounds.Width,box.Height/bounds.Height));
                matrix.Translate(-bounds.X,-bounds.Y);matrix.Scale(scale,scale,MatrixOrder.Append);
                matrix.Translate(box.X+(box.Width-bounds.Width*scale)/2,box.Y+(box.Height-bounds.Height*scale)/2,MatrixOrder.Append);glyph.Transform(matrix);graphics.FillPath(brush,glyph);
            }
        }
        public static void Release() { foreach (var image in cache.Values) image.Dispose(); cache.Clear(); }
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hwnd);
    }
}
