using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed class FocusFrame
    {
        public Rectangle Viewport { get; set; }
        public Rectangle[] Bands { get; set; } = new Rectangle[0];
    }
    public sealed class FocusOverlay : IDisposable
    {
        private readonly Func<object> application;
        private readonly Action<Action> queue;
        private readonly ToolPreferences settings;
        private readonly Timer timer = new Timer();
        private Layer layer;
        private IntPtr owner;
        private bool pending, disposed;
        private string signature;
        public bool Enabled { get; private set; }
        public bool IsVisible => layer?.Visible == true;
        public FocusOverlay(Func<object> application, Action<Action> queue, ToolPreferences settings)
        { this.application = application; this.queue = queue; this.settings = settings; timer.Tick += (s,e) => Pulse(); }
        public void Toggle() { Enabled = !Enabled; if (Enabled) { timer.Interval = settings.FocusInterval; timer.Start(); Pulse(); } else { timer.Stop(); layer?.Hide(); } }
        private void Pulse()
        {
            if (disposed || !Enabled) return;
            if (owner != IntPtr.Zero && GetForegroundWindow() != owner) layer?.Hide();
            if (pending) return; pending = true;
            try { queue(() => { pending = false; if (disposed || !Enabled) return; try { Update(); } catch { layer?.Hide(); } }); }
            catch { pending = false; layer?.Hide(); }
        }
        private void Update()
        {
            dynamic app = application(); IntPtr hwnd = ActiveWindowHandle(app);
            if (GetForegroundWindow() != hwnd || IsIconic(hwnd) || !(bool)app.Ready) { layer?.Hide(); return; }
            timer.Interval = Math.Max(50, Math.Min(500, settings.FocusInterval));
            var frame = Capture(application(), settings);
            if (frame.Bands.Length == 0) { layer?.Hide(); return; }
            if (owner != hwnd) { layer?.Dispose(); layer = null; owner = hwnd; signature = null; }
            if (layer == null) layer = new Layer();
            string next = frame.Viewport + "|" + string.Join(";", frame.Bands.Select(r => r.ToString())) + settings.FocusColor + settings.FocusOpacity;
            if (signature != next)
            {
                var region = new Region(); region.MakeEmpty();
                foreach (var band in frame.Bands) { var local = band; local.Offset(-frame.Viewport.X, -frame.Viewport.Y); region.Union(local); }
                var previous = layer.Region; layer.Bounds = frame.Viewport; layer.Region = region; previous?.Dispose();
                layer.BackColor = ColorTranslator.FromHtml(settings.FocusColor); layer.Opacity = Math.Max(.05, Math.Min(.6, settings.FocusOpacity / 100.0)); signature = next;
            }
            if (!layer.Visible) layer.Show(new Owner(hwnd));
        }
        public static FocusFrame Capture(object application, ToolPreferences settings)
        {
            dynamic app = application, sheet = app.ActiveSheet, window = app.ActiveWindow;
            if (sheet == null || window == null || Microsoft.VisualBasic.Information.TypeName((object)sheet).TrimStart('_') != "Worksheet" || (int)window.View != 1) return new FocusFrame();
            IntPtr hwnd = ActiveWindowHandle(app); Rectangle grid = GridBounds(hwnd);
            if (grid.Width == 0 || grid.Height == 0) return new FocusFrame();
            dynamic cell = app.ActiveCell.MergeArea, panes = window.Panes;
            var regions = new List<Tuple<Rectangle, Rectangle>>();
            for (int i = 1; i <= (int)panes.Count; i++)
            {
                dynamic pane = panes.Item[i], visible = pane.VisibleRange;
                var viewport = ClipToCells(window, Rectangle.Intersect(grid, ScreenRect(pane, visible)));
                regions.Add(Tuple.Create(viewport, ScreenRect(pane, cell)));
            }
            var bands = new List<Rectangle>();
            foreach (var pair in regions)
            {
                Rectangle viewport = pair.Item1, active = pair.Item2;
                // VisibleRange includes partially visible edge cells. Clip at adjacent pane origins.
                foreach (var other in regions.Where(p => !ReferenceEquals(p, pair)))
                {
                    Rectangle r = other.Item1;
                    if (r.Left > viewport.Left && r.Left < viewport.Right && r.Top < viewport.Bottom && r.Bottom > viewport.Top) viewport.Width = r.Left - viewport.Left;
                    if (r.Top > viewport.Top && r.Top < viewport.Bottom && r.Left < viewport.Right && r.Right > viewport.Left) viewport.Height = r.Top - viewport.Top;
                }
                if (settings.FocusRow)
                {
                    var row = Rectangle.Intersect(viewport, new Rectangle(viewport.Left, active.Top, viewport.Width, active.Height));
                    Add(bands, settings.ClearActiveCell ? Rectangle.FromLTRB(row.Left, row.Top, Math.Min(row.Right, active.Left), row.Bottom) : row);
                    if (settings.ClearActiveCell) Add(bands, Rectangle.FromLTRB(Math.Max(row.Left, active.Right), row.Top, row.Right, row.Bottom));
                }
                if (settings.FocusColumn)
                {
                    var column = Rectangle.Intersect(viewport, new Rectangle(active.Left, viewport.Top, active.Width, viewport.Height));
                    Add(bands, settings.ClearActiveCell ? Rectangle.FromLTRB(column.Left, column.Top, column.Right, Math.Min(column.Bottom, active.Top)) : column);
                    if (settings.ClearActiveCell) Add(bands, Rectangle.FromLTRB(column.Left, Math.Max(column.Top, active.Bottom), column.Right, column.Bottom));
                }
            }
            return new FocusFrame { Viewport = grid, Bands = bands.ToArray() };
        }
        internal static IntPtr ActiveWindowHandle(dynamic app)
        {
            try { return GetAncestor(new IntPtr((int)app.ActiveWindow.Hwnd), 2); }
            catch { return new IntPtr((int)app.Hwnd); }
        }
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
        private static Rectangle ClipToCells(dynamic window, Rectangle bounds)
        {
            if (bounds.Width <= 2 || bounds.Height <= 2) return Rectangle.Empty;
            // EXCEL7 also contains tab/scroll bars. RangeFromPoint gives the actual data edge.
            int x = bounds.Left + Math.Min(8, bounds.Width / 2), y = bounds.Top + Math.Min(8, bounds.Height / 2);
            if (!HasCell(window, x, y)) return Rectangle.Empty;
            int right = bounds.Right, bottom = bounds.Bottom;
            if (!HasCell(window, right - 1, y))
            {
                int low = x, high = right;
                while (high - low > 1) { int mid = low + (high - low)/2; if (HasCell(window,mid,y)) low=mid; else high=mid; }
                right = high;
            }
            if (!HasCell(window, x, bottom - 1))
            {
                int low = y, high = bottom;
                while (high - low > 1) { int mid = low + (high - low)/2; if (HasCell(window,x,mid)) low=mid; else high=mid; }
                bottom = high;
            }
            return Rectangle.FromLTRB(bounds.Left,bounds.Top,right,bottom);
        }
        private static bool HasCell(dynamic window, int x, int y)
        { try { return window.RangeFromPoint(x,y) != null; } catch (COMException) { return false; } }
        private static void Add(List<Rectangle> target, Rectangle rectangle) { if (rectangle.Width > 0 && rectangle.Height > 0) target.Add(rectangle); }
        private static Rectangle ScreenRect(dynamic pane, dynamic range)
        {
            double left = range.Left, top = range.Top, width = range.Width, height = range.Height;
            int x1 = pane.PointsToScreenPixelsX((int)Math.Round(left)), x2 = pane.PointsToScreenPixelsX((int)Math.Round(left + width));
            int y1 = pane.PointsToScreenPixelsY((int)Math.Round(top)), y2 = pane.PointsToScreenPixelsY((int)Math.Round(top + height));
            return Rectangle.FromLTRB(Math.Min(x1,x2), Math.Min(y1,y2), Math.Max(x1,x2), Math.Max(y1,y2));
        }
        private static Rectangle GridBounds(IntPtr owner)
        {
            Rectangle result = Rectangle.Empty;
            EnumChildWindows(owner, (handle, p) => { var name = new StringBuilder(64); GetClassName(handle, name, name.Capacity); if (name.ToString() == "EXCEL7" && IsWindowVisible(handle)) { GetClientRect(handle, out Rect r); var point = new Point(r.Left, r.Top); ClientToScreen(handle, ref point); var candidate = new Rectangle(point.X, point.Y, r.Right-r.Left, r.Bottom-r.Top); if ((long)candidate.Width*candidate.Height > (long)result.Width*result.Height) result = candidate; } return true; }, IntPtr.Zero);
            return result;
        }
        public void Dispose() { if (disposed) return; disposed = true; Enabled = false; timer.Dispose(); layer?.Dispose(); }
        private sealed class Owner : IWin32Window { public Owner(IntPtr handle) { Handle = handle; } public IntPtr Handle { get; } }
        private sealed class Layer : Form
        {
            public Layer() { FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; StartPosition = FormStartPosition.Manual; AutoScaleMode = AutoScaleMode.None; }
            protected override bool ShowWithoutActivation => true;
            protected override CreateParams CreateParams { get { var p = base.CreateParams; p.ExStyle |= 0x08000000 | 0x80 | 0x20; return p; } }
            protected override void WndProc(ref Message m) { if (m.Msg == 0x84) { m.Result = new IntPtr(-1); return; } if (m.Msg == 0x21) { m.Result = new IntPtr(3); return; } base.WndProc(ref m); }
        }
        [StructLayout(LayoutKind.Sequential)] private struct Rect { public int Left, Top, Right, Bottom; }
        private delegate bool Enumerate(IntPtr handle, IntPtr param);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, Enumerate callback, IntPtr param);
        [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr handle, StringBuilder name, int count);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr handle, out Rect rect);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr handle, ref Point point);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr handle);
    }
}
