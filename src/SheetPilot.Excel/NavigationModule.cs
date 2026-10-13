using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SheetPilot.Core;
using SheetPilot.UI;

namespace SheetPilot.Excel
{
    [ComVisible(true)]
    [Guid("DE860CD1-8AF1-47B1-B953-C890F436EA29")]
    [ProgId("SheetPilot.NavigationPane")]
    public sealed class NavigationPane : UserControl
    {
        private ElementHost host;
        public NavigationPane()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
        }
        internal void SetBackground(string color)
        {
            BackColor = System.Drawing.ColorTranslator.FromHtml(color);
            if (host != null) host.BackColor = BackColor;
        }
        internal void Attach(NavigationView view)
        {
            host = new ElementHost { Dock = DockStyle.Fill, Child = view };
            Controls.Add(host);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && host != null) { host.Child = null; host.Dispose(); host = null; }
            base.Dispose(disposing);
        }
    }

    public sealed class NavigationModule : IFeatureModule
    {
        private readonly IWorkbookGateway gateway;
        private readonly Dictionary<IntPtr, PaneSession> sessions = new Dictionary<IntPtr, PaneSession>();
        private readonly Timer timer;
        private readonly Timer saveTimer;
        private readonly SettingsStore store;
        private ExcelEvents events;
        private readonly NavigationSettings settings;
        public NavigationSettings Settings => settings;
        private readonly NativeLayoutGuard nativeLayout = new NativeLayoutGuard();
        private bool settingsDirty;
        private bool queued, disposed;
        public NavigationModule(IWorkbookGateway gateway)
        {
            this.gateway = gateway;
            string directory = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            string current = Path.Combine(directory, "GridPilot.navigation.json"), legacy = Path.Combine(directory, "SheetPilot.settings.json");
            if (!File.Exists(current) && File.Exists(legacy)) { try { File.Copy(legacy, current); } catch { current = legacy; } }
            store = new SettingsStore(current);
            settings = store.Load();
            timer = new Timer { Interval = settings.RefreshSeconds * 1000 };
            timer.Tick += (s,e) => Refresh();
            saveTimer = new Timer { Interval = 400 };
            saveTimer.Tick += (s,e) => SaveSettings();
            settings.PropertyChanged += SettingsChanged;

        }
        private void SettingsChanged(object sender, PropertyChangedEventArgs e)
        {
            if (disposed) return;
            settingsDirty = true;
            saveTimer.Stop(); saveTimer.Start();
            foreach (var session in sessions.Values)
            {
                session.Model.SettingsStatus = "正在保存…";
                if (!nativeLayout.IsObserving)
                {
                    if (e.PropertyName == nameof(settings.DockSide)) session.RequestLayout(true, false);
                    else if (e.PropertyName == nameof(settings.PaneWidth)) session.RequestLayout(false, true);
                }
                if (e.PropertyName == nameof(settings.BackgroundColor)) session.SetBackground();
            }
            if (e.PropertyName == nameof(settings.RefreshSeconds) || e.PropertyName == nameof(settings.AutoRefresh))
            {
                timer.Interval = settings.RefreshSeconds * 1000;
                if (settings.AutoRefresh && sessions.Count > 0) timer.Start(); else timer.Stop();
            }
        }
        private void SaveSettings()
        {
            saveTimer.Stop();
            if (!settingsDirty) return;
            settingsDirty = false;
            store.Save(settings);
            foreach (var session in sessions.Values) { session.Model.SettingsStatus = store.Status; session.Model.SettingsSaveFailed = store.HasError; }
        }
        public string Id => "SheetPilot_Navigation";
        public string Label => "导航";
        public IEnumerable<FeatureCommand> Commands => new[] {
            new FeatureCommand("SheetPilot_Show", "工作表导航", "显示可搜索的页面目录，支持正则表达式与跨工作簿跳转。", "GoTo", Show),
            new FeatureCommand("SheetPilot_Refresh", "刷新目录", "重新读取页面、顺序及隐藏状态。", "RefreshAll", Refresh)
        };
        public void ShowSettings()
        {
            Show();
            if (sessions.TryGetValue(ExcelDnaUtil.WindowHandle, out var session)) session.Model.IsSettingsOpen = true;
        }
        public void Show()
        {
            if (disposed) return;
            if(events==null){events=new ExcelEvents(ExcelDnaUtil.Application);try{events.Activation(Refresh);}catch(Exception ex){Log.Error("Navigation events",ex);}}
            // Office 2010 is MDI: one application pane. Office 2013+ uses one pane per window.
            IntPtr window = ExcelDnaUtil.WindowHandle;
            if (!sessions.TryGetValue(window, out var session))
            {
                session = new PaneSession(settings, nativeLayout, Refresh, Navigate, () => { if (settings.AutoRefresh) timer.Start(); Refresh(); });
                session.Model.SettingsStatus = store.Status;
                session.Model.SettingsSaveFailed = store.HasError;
                sessions.Add(window, session);
            }
            session.Pane.Visible = true;
            session.View.FocusSearch();
            if (settings.AutoRefresh) timer.Start();
            Refresh();
        }
        public void Refresh()
        {
            if (disposed || queued || sessions.Count == 0) return;
            queued = true;
            try
            {
                // At most one pending refresh even while Excel is in cell-edit / modal mode.
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    queued = false;
                    if (disposed) return;
                    try
                    {
                        foreach (var dead in sessions.Keys.Where(h => !IsWindow(h)).ToArray()) { sessions[dead].Dispose(); sessions.Remove(dead); }
                        WorkbookSnapshot snapshot = null;
                        foreach (var session in sessions.Values.Where(p => p.Pane.Visible && !p.IsResizing))
                        {
                            if (snapshot == null) snapshot = gateway.Capture(settings.ScopeIndex == 1);
                            session.Model.Update(snapshot);
                        }
                        if (!sessions.Values.Any(p => p.Pane.Visible)) timer.Stop();
                    }
                    catch (Exception ex) { Log.Error("Refresh", ex); foreach (var item in sessions.Values) item.Model.Report("Excel 暂时无法读取目录。结束编辑或关闭对话框后重试。"); }
                });
            }
            catch (Exception ex) { queued = false; Log.Error("Queue refresh", ex); }
        }
        private void Navigate(PaneSession session, SheetItem item)
        {
            if (disposed || session.NavigationPending) return;
            session.NavigationPending = true;
            try
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    session.NavigationPending = false;
                    if (disposed || session.Disposed) return;
                    try
                    {
                        gateway.Navigate(item.WorkbookKey, item);
                        if (session.Pane.Visible && !sessions.ContainsKey(ExcelDnaUtil.WindowHandle))
                        { Show(); sessions[ExcelDnaUtil.WindowHandle].Model.Query = session.Model.Query; }
                        Refresh();
                        WorksheetFocus.Restore(ExcelDnaUtil.WindowHandle);
                    }
                    catch (Exception ex) { Log.Error("Navigate", ex); session.Model.Report(ex is InvalidOperationException ? ex.Message : "Excel 暂时无法跳转，请结束编辑后重试。"); }
                });
            }
            catch (Exception ex) { session.NavigationPending = false; Log.Error("Queue navigation", ex); session.Model.Report("Excel 暂时无法接受跳转，请稍后重试。"); }
        }
        public void Dispose()
        {
            if (disposed) return;
            foreach (var session in sessions.Values) session.FlushWidth();
            disposed = true;
            events?.Dispose();
            settings.PropertyChanged -= SettingsChanged;
            SaveSettings();
            saveTimer.Dispose();
            timer.Stop(); timer.Dispose();
            foreach (var session in sessions.Values) session.Dispose();
            sessions.Clear();
        }
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr handle);

        private sealed class PaneSession : IDisposable
        {
            public CustomTaskPane Pane { get; private set; }
            public NavigationModel Model { get; }
            public NavigationView View { get; }
            public bool NavigationPending { get; set; }
            public bool Disposed { get; private set; }
            private readonly NavigationPane control;
            private readonly NavigationSettings settings;
            private readonly NativeLayoutGuard nativeLayout;
            private readonly Timer resizeTimer = new Timer { Interval = 250 };
            private readonly Timer layoutTimer = new Timer { Interval = 75 };
            private bool pendingDock, pendingWidth;
            private readonly Action becameVisible;
            public bool IsResizing => resizeTimer.Enabled;
            private bool applyingLayout;
            public PaneSession(NavigationSettings settings, NativeLayoutGuard nativeLayout, Action refresh, Action<PaneSession, SheetItem> navigate, Action becameVisible)
            {
                this.settings = settings;
                this.nativeLayout = nativeLayout;
                this.becameVisible = becameVisible;
                resizeTimer.Tick += (s,e) => FlushWidth();
                layoutTimer.Tick += (s,e) => { layoutTimer.Stop(); bool dock = pendingDock, width = pendingWidth; pendingDock = pendingWidth = false; ApplyLayout(dock, width); };
                Model = new NavigationModel(settings);
                View = new NavigationView(Model);
                control = new NavigationPane();
                control.Attach(View);
                SetBackground();
                // Office may dispose its child while already tearing down the COM pane.
                // Release managed work here; never reenter Pane.Delete from this callback.
                control.Disposed += (s,e) => { resizeTimer.Stop(); layoutTimer.Stop(); View.Dispose(); Model.Dispose(); };
                try
                {
                    // Use Office's docking/parent mechanism; no SetParent, owner hacks, or embedded browser.
                    Pane = CustomTaskPaneFactory.CreateCustomTaskPane(control, "GridPilot");
                    if (Pane == null) throw new InvalidOperationException("Office 无法创建任务面板，请检查当前用户的加载项权限。");
                    ApplyLayout();
                    Pane.DockPositionStateChange += DockChanged;
                    Pane.VisibleStateChange += VisibleChanged;
                    control.SizeChanged += (s,e) =>
                    {
                        if (Disposed || applyingLayout) return;
                        // While the native splitter owns the interaction, never write geometry back.
                        layoutTimer.Stop(); pendingDock = pendingWidth = false;
                        resizeTimer.Stop(); resizeTimer.Start();
                    };
                    Model.ScopeChanged += all => refresh();
                    Model.RefreshRequested += refresh;
                    Model.NavigateRequested += item => navigate(this, item);
                    Log.Write("Task pane created; dock=" + Pane.DockPosition + "; width=" + settings.PaneWidth + "; nativeWidth=" + Pane.Width);
                }
                catch { resizeTimer.Dispose(); layoutTimer.Dispose(); View.Dispose(); Model.Dispose(); control.Dispose(); if (Pane != null) { try { Pane.Delete(); } catch { } } throw; }
            }
            public void SetBackground() => control.SetBackground(settings.BackgroundColor);
            private void VisibleChanged(CustomTaskPane pane) { if (!Disposed && pane.Visible) becameVisible(); }
            public void RequestLayout(bool dock, bool width)
            {
                if (Disposed) return;
                resizeTimer.Stop();
                pendingDock |= dock; pendingWidth |= width;
                if (!layoutTimer.Enabled) layoutTimer.Start();
            }
            private void ApplyLayout(bool dock = true, bool width = true)
            {
                if (Disposed || Pane == null || applyingLayout) return;
                applyingLayout = true;
                try
                {
                    var side = settings.DockSide == 1 ? MsoCTPDockPosition.msoCTPDockPositionLeft : MsoCTPDockPosition.msoCTPDockPositionRight;
                    if (dock && Pane.DockPosition != side) Pane.DockPosition = side;
                    int pixels = (int)Math.Round(settings.PaneWidth * DpiScale());
                    if (width && Pane.Width != pixels) Pane.Width = pixels;
                }
                catch (Exception ex) { Log.Error("Pane layout", ex); Model.Report("Excel 暂时无法调整停靠位置或宽度，请稍后重试。"); }
                finally { applyingLayout = false; }
            }
            private void DockChanged(CustomTaskPane pane)
            {
                if (Disposed || applyingLayout) return;
                try
                {
                    nativeLayout.Observe(() =>
                    {
                        if (pane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionLeft) settings.DockSide = 1;
                        else if (pane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionRight) settings.DockSide = 0;
                    });
                }
                catch (Exception ex) { Log.Error("Remember dock", ex); }
            }
            public void FlushWidth()
            {
                if (!resizeTimer.Enabled) return;
                resizeTimer.Stop();
                if (Disposed || applyingLayout || Pane == null) return;
                try
                {
                    if (Pane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionLeft || Pane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionRight)
                        nativeLayout.Observe(() => settings.PaneWidth = (int)Math.Round(Pane.Width / DpiScale()));
                }
                catch (Exception ex) { Log.Error("Remember width", ex); }
            }
            public void Dispose()
            {
                if (Disposed) return;
                Disposed = true;
                resizeTimer.Dispose(); layoutTimer.Dispose();
                View.Dispose();
                Model.Dispose();
                try { if (Pane != null) { Pane.DockPositionStateChange -= DockChanged; Pane.VisibleStateChange -= VisibleChanged; } }
                catch (Exception ex) { Log.Error("Detach pane events", ex); }
                try { Pane?.Delete(); } catch (Exception ex) { Log.Error("Delete pane", ex); }
                Pane = null;
                control.Dispose();
            }
            private double DpiScale()
            {
                try { uint dpi=GetDpiForWindow(control.Handle);if(dpi>0)return dpi/96.0; } catch(EntryPointNotFoundException) { }
                using(var graphics=control.CreateGraphics())return graphics.DpiX/96.0;
            }
            [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hwnd);
        }
    }
}
