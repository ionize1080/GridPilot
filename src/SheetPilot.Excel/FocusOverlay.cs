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
    public sealed class RippleBand { public Rectangle Bounds; public int Opacity; }
    public sealed class FocusFrame
    {
        public RippleBand[] Ripple { get; set; } = new RippleBand[0];
        public Rectangle Viewport { get; set; }
        public Rectangle[] Bands { get; set; } = new Rectangle[0];
        public Rectangle[] Panes { get; set; } = new Rectangle[0];
        public Rectangle[] Cells { get; set; } = new Rectangle[0];
    }
    public sealed class FocusOverlay : IDisposable
    {
        private readonly Func<object> application;
        private readonly ToolPreferences settings;
        private readonly Timer timer = new Timer();
        private readonly ExcelEvents events;
        private readonly uint processId;
        private readonly Dictionary<Rectangle, Rectangle> edges = new Dictionary<Rectangle, Rectangle>();
        private Layer layer;
        private IntPtr owner;
        private bool updating, disposed;
        private string signature, motion;
        private int lastMotion, lastPoll;
        private FocusFrame displayed, destination, departure;
        private readonly System.Diagnostics.Stopwatch animation = new System.Diagnostics.Stopwatch();
        private string layout;
        private int lastSelection;
        private readonly System.Diagnostics.Stopwatch ripple=new System.Diagnostics.Stopwatch();
        public int PresentCount { get; private set; }
        public int ShowCount { get; private set; }
        public int HideCount { get; private set; }
        public int UpdateCount { get; private set; }
        public double LastUpdateMilliseconds { get; private set; }
        private readonly Queue<double> timings = new Queue<double>();
        public string PerformanceSummary
        {
            get { var values=timings.OrderBy(v=>v).ToArray(); return values.Length==0?"No focus samples":"Focus updates="+UpdateCount+"; recent="+values.Length+"; p50="+values[values.Length/2].ToString("F2",System.Globalization.CultureInfo.InvariantCulture)+"ms; p95="+values[Math.Min(values.Length-1,(int)(values.Length*.95))].ToString("F2",System.Globalization.CultureInfo.InvariantCulture)+"ms; max="+values.Last().ToString("F2",System.Globalization.CultureInfo.InvariantCulture)+"ms; presents="+PresentCount+"; shows="+ShowCount+"; hides="+HideCount+"; visible="+IsVisible; }
        }
        public bool Enabled { get; private set; }
        public bool IsVisible => layer?.Visible == true;
        internal bool IsAnimating => animation.IsRunning || ripple.IsRunning;
        internal IntPtr OwnerHandle => owner;
        public FocusOverlay(Func<object> application, Action<Action> queue, ToolPreferences settings)
        {
            this.application = application; this.settings = settings; timer.Interval = 16; timer.Tick += (s,e) => Pulse(false);
            GetWindowThreadProcessId(ActiveWindowHandle(application()), out processId);
            events = new ExcelEvents(application());
            try { events.Selection(() => Pulse(true)); events.Activation(() => { motion = null; edges.Clear(); Pulse(true); }); }
            catch (COMException) { /* The lightweight timer remains a fallback on restricted hosts. */ }
        }
        public void Toggle() { Enabled = !Enabled; if (Enabled) { timer.Start(); Pulse(true); } else { timer.Stop(); HideLayer(); } }
        private void Pulse(bool immediate)
        {
            if (disposed || !Enabled || updating) return;
            updating = true;
            try
            {
                // Native visibility checks keep running even when Excel rejects COM while a menu is open.
                if (owner != IntPtr.Zero && !WorksheetExposed(owner))
                {
                    HideLayer(); motion = null;
                    IntPtr foreground=GetForegroundWindow(); GetWindowThreadProcessId(foreground,out uint activeProcess);
                    // An SDI workbook in the same Excel instance must be allowed to acquire a new owner.
                    if(foreground==owner || activeProcess!=processId || ClassName(foreground)!="XLMAIN") return;
                }
                if (!immediate && animation.IsRunning) Animate();
                else if(!immediate && ripple.IsRunning) AnimateRipple();
                int interval = unchecked(Environment.TickCount-lastMotion)>250 ? 100 : Math.Max(16, Math.Min(100, settings.FocusInterval));
                if (!immediate && IsVisible && unchecked(Environment.TickCount-lastPoll)<interval) return;
                lastPoll=Environment.TickCount;
                Update(immediate);
            }
            // Excel can transiently reject calls during a normal mouse click. Preserve the last frame;
            // the independent native guard still hides menus, editing, dialogs and background windows.
            catch (COMException) { motion = null; }
            catch { HideLayer(); motion = null; }
            finally { updating = false; }
        }
        private void Update(bool immediate)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            dynamic app = application(); IntPtr hwnd = ActiveWindowHandle(app);
            if (!WorksheetExposed(hwnd)) { HideLayer(); return; }
            if (!(bool)app.Ready) return;
            dynamic window = app.ActiveWindow;
            GetClientRect(hwnd, out Rect bounds);
            var origin = new Point(); ClientToScreen(hwnd, ref origin);
            string nextLayout = hwnd + "|" + origin + "|" + bounds.Right + ":" + bounds.Bottom + "|" + window.ScrollRow + ":" + window.ScrollColumn + ":" + window.Zoom + ":" + window.View + ":" + window.SplitRow + ":" + window.SplitColumn + ":" + window.FreezePanes + "|" + app.ActiveSheet.Name;
            string nextMotion = nextLayout + "|" + app.ActiveCell.Address + "|" + app.ActiveCell.Width + ":" + app.ActiveCell.Height + "|" + settings.FocusColor + settings.FocusOpacity + settings.FocusRow + settings.FocusColumn + settings.ClearActiveCell + settings.FocusAnimation;
            if (motion == nextMotion && IsVisible) return;
            lastMotion=Environment.TickCount;
            if (edges.Count > 64 || owner != hwnd) edges.Clear();
            // Geometry can change without a selection event (row resize, pane drag, DPI).
            if (layout != nextLayout) edges.Clear();
            var frame = Capture(application(), settings, edges);
            motion = nextMotion;
            if (frame.Bands.Length == 0) { HideLayer(); return; }
            if (owner != hwnd) { layer?.Dispose(); layer = null; owner = hwnd; signature = null; }
            if (layer == null) layer = new Layer();
            string next = frame.Viewport + "|" + string.Join(";", frame.Bands.Select(r => r.ToString())) + settings.FocusColor + settings.FocusOpacity;
            if (signature != next)
            {
                bool glide = immediate && layout == nextLayout && IsVisible && AnimationAllowed() &&
                    unchecked(Environment.TickCount - lastSelection) > 90 && CanInterpolate(displayed, frame);
                departure = displayed; destination = frame;
                animation.Reset();
                if (immediate && settings.FocusRipple && AnimationAllowed(true) && unchecked(Environment.TickCount-lastSelection)>90)ripple.Restart();else ripple.Reset();
                if (glide && !ripple.IsRunning) { animation.Start(); Animate(); }
                else Present(frame);
                signature = next; PresentCount++;
            }
            layout = nextLayout;
            if (immediate) lastSelection = Environment.TickCount;
            if (!layer.Visible) { layer.Show(new Owner(hwnd)); ShowCount++; }
            LastUpdateMilliseconds = watch.Elapsed.TotalMilliseconds; UpdateCount++;
            timings.Enqueue(LastUpdateMilliseconds); if(timings.Count>256)timings.Dequeue();
        }
        private bool AnimationAllowed(bool forRipple=false)
        {
            bool enabled = false;
            return (forRipple?settings.FocusRipple:settings.FocusAnimation) && !SystemInformation.HighContrast && !SystemInformation.TerminalServerSession &&
                SystemParametersInfo(0x1042, 0, ref enabled, 0) && enabled;
        }
        private void Present(FocusFrame frame)
        { layer.Present(frame, ColorTranslator.FromHtml(settings.FocusColor), settings.FocusOpacity); displayed = frame; }
        private void Animate()
        {
            double progress = Math.Min(1, animation.Elapsed.TotalMilliseconds / 120.0);
            if (!AnimationAllowed()) progress = 1;
            Present(Interpolate(departure, destination, progress, settings));
            if (progress >= 1) animation.Stop();
        }
        private void AnimateRipple()
        {
            double progress=Math.Min(1,ripple.Elapsed.TotalMilliseconds/(double)settings.RippleDuration);
            if(!AnimationAllowed(true))progress=1;
            if(progress>=1){ripple.Stop();Present(destination);return;}
            Present(RippleFrame(destination,progress,settings));
        }
        public static FocusFrame RippleFrame(FocusFrame frame,double progress,ToolPreferences options)
        {
            if(frame==null || progress>=1)return frame;
            var stripes=new List<RippleBand>();double distance=1-Math.Pow(1-Math.Max(0,progress),2);
            int width=options.RippleWidth;double fade=Math.Pow(1-progress,1.3);
            for(int i=0;i<frame.Cells.Length;i++)
            {
                var cell=frame.Cells[i];var pane=frame.Panes[i];
                if(!pane.IntersectsWith(cell))continue;
                int cx=cell.Left+cell.Width/2,cy=cell.Top+cell.Height/2;
                for(int trail=0;trail<8;trail++)
                {
                    int bandWidth=Math.Max(1,width/8),alpha=(int)(options.RippleStrength*fade*Math.Sin((trail+1)*Math.PI/9));
                    if(alpha<=0)continue;
                    foreach(int direction in new[]{-1,1})
                    {
                        if(options.FocusRow)
                        {int edge=direction<0?pane.Left:pane.Right;int x=cx+(int)((edge-cx)*distance)-direction*trail*bandWidth;var rect=Rectangle.Intersect(pane,new Rectangle(x,cell.Top,bandWidth,cell.Height));AddRipple(stripes,rect,cell,alpha);}
                        if(options.FocusColumn)
                        {int edge=direction<0?pane.Top:pane.Bottom;int y=cy+(int)((edge-cy)*distance)-direction*trail*bandWidth;var rect=Rectangle.Intersect(pane,new Rectangle(cell.Left,y,cell.Width,bandWidth));AddRipple(stripes,rect,cell,alpha);}
                    }
                }
            }
            return new FocusFrame {Viewport=frame.Viewport,Cells=frame.Cells,Panes=frame.Panes,Bands=frame.Bands,Ripple=stripes.ToArray()};
        }
        private static void AddRipple(List<RippleBand> output,Rectangle rect,Rectangle cell,int opacity)
        {var bands=new List<Rectangle>();AddClippedBand(bands,rect,cell,true);foreach(var band in bands)output.Add(new RippleBand {Bounds=band,Opacity=opacity});}
        public static bool CanInterpolate(FocusFrame a, FocusFrame b)
        {
            if (a == null || b == null || a.Viewport != b.Viewport || a.Cells.Length == 0 ||
                a.Cells.Length != b.Cells.Length || !a.Panes.SequenceEqual(b.Panes)) return false;
            for (int i = 0; i < a.Cells.Length; i++)
                if (Math.Abs(a.Cells[i].X-b.Cells[i].X) > b.Panes[i].Width/2 ||
                    Math.Abs(a.Cells[i].Y-b.Cells[i].Y) > b.Panes[i].Height/2) return false;
            return true;
        }
        public static FocusFrame Interpolate(FocusFrame a, FocusFrame b, double progress, ToolPreferences settings)
        {
            if (progress >= 1 || !CanInterpolate(a,b)) return b;
            double t = 1-Math.Pow(1-Math.Max(0,progress),3);
            var cells = new Rectangle[b.Cells.Length]; var bands = new List<Rectangle>();
            for (int i=0;i<cells.Length;i++)
            {
                var x=a.Cells[i]; var y=b.Cells[i];
                cells[i]=new Rectangle((int)Math.Round(x.X+(y.X-x.X)*t),(int)Math.Round(x.Y+(y.Y-x.Y)*t),
                    (int)Math.Round(x.Width+(y.Width-x.Width)*t),(int)Math.Round(x.Height+(y.Height-x.Height)*t));
                var pane=b.Panes[i]; var cell=cells[i];
                if(settings.FocusRow) AddClippedBand(bands,Rectangle.Intersect(pane,new Rectangle(pane.Left,cell.Top,pane.Width,cell.Height)),y,settings.ClearActiveCell);
                if(settings.FocusColumn) AddClippedBand(bands,Rectangle.Intersect(pane,new Rectangle(cell.Left,pane.Top,cell.Width,pane.Height)),y,settings.ClearActiveCell);
            }
            return new FocusFrame { Viewport=b.Viewport, Panes=b.Panes, Cells=cells, Bands=bands.ToArray() };
        }
        private static void AddClippedBand(List<Rectangle> bands, Rectangle band, Rectangle actual, bool clear)
        {
            var cut=Rectangle.Intersect(band,actual);
            if(!clear || cut.Width<=0 || cut.Height<=0) { Add(bands,band); return; }
            Add(bands,Rectangle.FromLTRB(band.Left,band.Top,band.Right,cut.Top));
            Add(bands,Rectangle.FromLTRB(band.Left,cut.Bottom,band.Right,band.Bottom));
            Add(bands,Rectangle.FromLTRB(band.Left,cut.Top,cut.Left,cut.Bottom));
            Add(bands,Rectangle.FromLTRB(cut.Right,cut.Top,band.Right,cut.Bottom));
        }
        [DllImport("user32.dll")] private static extern bool SystemParametersInfo(uint action,uint parameter,ref bool value,uint flags);
        public static FocusFrame Capture(object application, ToolPreferences settings) => Capture(application, settings, null);
        private static FocusFrame Capture(object application, ToolPreferences settings, Dictionary<Rectangle, Rectangle> edges)
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
                var raw = Rectangle.Intersect(grid, ScreenRect(pane, visible));
                Rectangle viewport = Rectangle.Empty;
                if (edges == null || !edges.TryGetValue(raw, out viewport))
                { viewport = ClipToCells(window, raw); if (edges != null) edges[raw] = viewport; }
                regions.Add(Tuple.Create(viewport, ScreenRect(pane, cell)));
            }
            var bands = new List<Rectangle>();
            var clippedPanes = new List<Rectangle>();
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
                clippedPanes.Add(viewport);
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
            return new FocusFrame { Viewport = grid, Bands = bands.ToArray(), Panes=clippedPanes.ToArray(), Cells=regions.Select(p=>p.Item2).ToArray() };
        }
        internal static IntPtr ActiveWindowHandle(dynamic app)
        {
            try { var root=GetAncestor(new IntPtr((int)app.ActiveWindow.Hwnd), 2); if(root!=IntPtr.Zero && IsWindow(root)) return root; } catch { }
            try { var fallback=new IntPtr((int)app.Hwnd); return IsWindow(fallback)?fallback:IntPtr.Zero; } catch { return IntPtr.Zero; }
        }
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);
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
        public void Dispose() { if (disposed) return; disposed = true; Enabled = false; events.Dispose(); timer.Dispose(); layer?.Dispose(); }
        private sealed class Owner : IWin32Window { public Owner(IntPtr handle) { Handle = handle; } public IntPtr Handle { get; } }
        private void HideLayer() { animation.Reset(); ripple.Reset(); displayed=null; layout=null; signature=null; motion=null; if (IsVisible) { layer.Hide(); HideCount++; } }
        internal static bool WorksheetExposed(IntPtr hwnd)
        {
            if (GetForegroundWindow()!=hwnd || IsIconic(hwnd) || !IsWindowEnabled(hwnd)) return false;
            var info = new GuiThreadInfo { Size = Marshal.SizeOf(typeof(GuiThreadInfo)) };
            if (GetGUIThreadInfo(0, ref info))
            {
                if ((info.Flags & 0x1C) != 0) return false; // menu, system menu, popup menu
                string focusClass = ClassName(info.Focus);
                if (focusClass == "EXCEL6" || focusClass == "EXCEL<") return false; // in-cell/formula editor
            }
            Rectangle grid=GridBounds(hwnd);
            if (grid.Width<=0 || grid.Height<=0) return false;
            // EXCEL7 may still have WS_VISIBLE behind Backstage. Follow the actual child z-order,
            // not just its visibility bit. Our owned overlay is not a child and cannot mask this hit test.
            var point=new Point(grid.Left+grid.Width/2,grid.Top+grid.Height/2);
            IntPtr current=hwnd;
            for(int i=0;i<16;i++)
            {
                if(ClassName(current)=="EXCEL7") return true;
                var local=point; ScreenToClient(current,ref local);
                IntPtr child=ChildWindowFromPointEx(current,local,1|4);
                if(child==IntPtr.Zero || child==current) break;
                current=child;
            }
            return false;
        }
        private static string ClassName(IntPtr window) { var value=new StringBuilder(128);GetClassName(window,value,value.Capacity);return value.ToString(); }
        // One per-pixel-alpha bitmap is committed atomically. No Region replacement, Opacity changes,
        // background erase or hide/show cycle between selection frames.
        private sealed class Layer : Form
        {
            public Layer() { FormBorderStyle=FormBorderStyle.None; ShowInTaskbar=false; StartPosition=FormStartPosition.Manual; AutoScaleMode=AutoScaleMode.None; }
            protected override bool ShowWithoutActivation => true;
            protected override CreateParams CreateParams { get { var p=base.CreateParams;p.ExStyle|=0x80000|0x08000000|0x80|0x20;return p; } }
            protected override void OnPaintBackground(PaintEventArgs e) { }
            protected override void OnPaint(PaintEventArgs e) { }
            private IntPtr memory, bitmap, previousBitmap, pixels;
            private Size bitmapSize;
            private Rectangle[] previousBands = new Rectangle[0];
            private int[] clearRow, colorRow, rippleRow;
            private int lastColor;
            public void Present(FocusFrame frame, Color color, int opacity)
            {
                if(bitmapSize!=frame.Viewport.Size || bitmap==IntPtr.Zero)
                {
                    ReleaseBitmap();
                    memory=CreateCompatibleDC(IntPtr.Zero);
                    var info=new BitmapInfo {Size=40,Width=frame.Viewport.Width,Height=-frame.Viewport.Height,Planes=1,Bits=32};
                    bitmap=CreateDIBSection(memory,ref info,0,out pixels,IntPtr.Zero,0);
                    if(bitmap==IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                    previousBitmap=SelectObject(memory,bitmap);bitmapSize=frame.Viewport.Size;
                    clearRow=new int[bitmapSize.Width];colorRow=new int[bitmapSize.Width];rippleRow=new int[bitmapSize.Width];lastColor=0;
                    Fill(new Rectangle(Point.Empty,bitmapSize),clearRow);
                }
                // Reuse the top-down DIB and update only the old/new stripes, not a full-screen
                // managed Bitmap plus GDI copy per click. Pixels are opaque BGRA; the blend sets opacity.
                foreach(var band in previousBands) Fill(band,clearRow);
                int packed=Premultiplied(color,opacity);if(lastColor!=packed){lastColor=packed;for(int x=0;x<colorRow.Length;x++)colorRow[x]=packed;}
                previousBands=frame.Bands.Select(b=>new Rectangle(b.X-frame.Viewport.X,b.Y-frame.Viewport.Y,b.Width,b.Height)).ToArray();
                foreach(var band in previousBands) Fill(band,colorRow);
                var dirty=previousBands.ToList();
                foreach(var band in frame.Ripple)
                {var bounds=new Rectangle(band.Bounds.X-frame.Viewport.X,band.Bounds.Y-frame.Viewport.Y,band.Bounds.Width,band.Bounds.Height);int value=Premultiplied(color,Math.Min(85,opacity+band.Opacity));for(int x=0;x<bounds.Width;x++)rippleRow[x]=value;Fill(bounds,rippleRow);dirty.Add(bounds);}
                previousBands=dirty.ToArray();
                var location=frame.Viewport.Location;var source=Point.Empty;var size=bitmapSize;
                var blend=new Blend {Alpha=255,Format=1};
                if(!UpdateLayeredWindow(Handle,IntPtr.Zero,ref location,ref size,memory,ref source,0,ref blend,2)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            }
            private static int Premultiplied(Color color,int opacity){int a=255*opacity/100;return (a<<24)|((color.R*a/255)<<16)|((color.G*a/255)<<8)|(color.B*a/255);}
            private void Fill(Rectangle rectangle,int[] row)
            {
                for(int y=rectangle.Top;y<rectangle.Bottom;y++)
                    Marshal.Copy(row,0,IntPtr.Add(pixels,(y*bitmapSize.Width+rectangle.Left)*4),rectangle.Width);
            }
            private void ReleaseBitmap()
            {
                if(memory!=IntPtr.Zero && previousBitmap!=IntPtr.Zero)SelectObject(memory,previousBitmap);
                if(bitmap!=IntPtr.Zero)DeleteObject(bitmap);
                if(memory!=IntPtr.Zero)DeleteDC(memory);
                bitmap=memory=previousBitmap=pixels=IntPtr.Zero;previousBands=new Rectangle[0];
            }
            protected override void Dispose(bool disposing) { ReleaseBitmap();base.Dispose(disposing); }
            protected override void WndProc(ref Message m) { if(m.Msg==0x84){m.Result=new IntPtr(-1);return;}if(m.Msg==0x21){m.Result=new IntPtr(3);return;}base.WndProc(ref m); }
            [StructLayout(LayoutKind.Sequential,Pack=1)] private struct Blend { public byte Operation, Flags, Alpha, Format; }
            [DllImport("user32.dll",SetLastError=true)] private static extern bool UpdateLayeredWindow(IntPtr hwnd,IntPtr dest,ref Point location,ref Size size,IntPtr source,ref Point origin,int color,ref Blend blend,uint flags);
            [StructLayout(LayoutKind.Sequential)] private struct BitmapInfo {public int Size,Width,Height;public short Planes,Bits;public int Compression,ImageSize,XResolution,YResolution,Colors,ImportantColors;}
            [DllImport("gdi32.dll",SetLastError=true)] private static extern IntPtr CreateDIBSection(IntPtr dc,ref BitmapInfo info,uint usage,out IntPtr pixels,IntPtr section,uint offset);
            [DllImport("gdi32.dll")] private static extern IntPtr CreateCompatibleDC(IntPtr dc);
            [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr dc,IntPtr item);
            [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr item);
            [DllImport("gdi32.dll")] private static extern bool DeleteDC(IntPtr dc);
        }
        [StructLayout(LayoutKind.Sequential)] private struct GuiThreadInfo { public int Size, Flags; public IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret; public Rect CaretRect; }
        [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread,ref GuiThreadInfo info);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint process);
        [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr hwnd,ref Point point);
        [DllImport("user32.dll")] private static extern IntPtr ChildWindowFromPointEx(IntPtr hwnd,Point point,uint flags);
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
