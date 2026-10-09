using System;
using System.Collections.Generic;
using System.Linq;
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
        public NavigationPane() { AutoScaleMode = AutoScaleMode.Dpi; }
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
        private bool queued, disposed;
        public NavigationModule(IWorkbookGateway gateway)
        {
            this.gateway = gateway;
            timer = new Timer { Interval = 1200 };
            timer.Tick += (s,e) => Refresh();
        }
        public string Id => "SheetPilot_Navigation";
        public string Label => "工作簿工具";
        public IEnumerable<FeatureCommand> Commands => new[] {
            new FeatureCommand("SheetPilot_Show", "工作表导航", "显示可搜索的页面目录，支持正则表达式与跨工作簿跳转。", "GoTo", Show),
            new FeatureCommand("SheetPilot_Refresh", "刷新目录", "重新读取页面、顺序及隐藏状态。", "RefreshAll", Refresh)
        };
        public void Show()
        {
            if (disposed) return;
            // Office 2010 is MDI: one application pane. Office 2013+ uses one pane per window.
            IntPtr window = ExcelDnaUtil.WindowHandle;
            if (!sessions.TryGetValue(window, out var session))
            {
                session = new PaneSession(window, Refresh, Navigate);
                sessions.Add(window, session);
            }
            session.Pane.Visible = true;
            session.View.FocusSearch();
            timer.Start();
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
                        var current = ExcelDnaUtil.WindowHandle;
                        if (sessions.TryGetValue(current, out var session) && session.Pane.Visible)
                            session.Model.Update(gateway.Capture(session.Model.ScopeIndex == 1));
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
                    try { gateway.Navigate(item.WorkbookKey, item); Refresh(); }
                    catch (Exception ex) { Log.Error("Navigate", ex); session.Model.Report(ex is InvalidOperationException ? ex.Message : "Excel 暂时无法跳转，请结束编辑后重试。"); }
                });
            }
            catch (Exception ex) { session.NavigationPending = false; Log.Error("Queue navigation", ex); session.Model.Report("Excel 暂时无法接受跳转，请稍后重试。"); }
        }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
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
            public PaneSession(IntPtr handle, Action refresh, Action<PaneSession, SheetItem> navigate)
            {
                Model = new NavigationModel();
                View = new NavigationView(Model);
                control = new NavigationPane();
                control.Attach(View);
                control.Disposed += (s,e) => Model.Dispose();
                try
                {
                    // Use Office's docking/parent mechanism; no SetParent, owner hacks, or embedded browser.
                    Pane = CustomTaskPaneFactory.CreateCustomTaskPane(control, "SheetPilot · 工作表导航");
                    if (Pane == null) throw new InvalidOperationException("Office 无法创建任务面板，请检查当前用户的加载项权限。");
                    Pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft;
                    Pane.Width = 340;
                    Model.ScopeChanged += all => refresh();
                    Model.RefreshRequested += refresh;
                    Model.NavigateRequested += item => navigate(this, item);
                    Log.Write("Task pane created");
                }
                catch { Model.Dispose(); control.Dispose(); if (Pane != null) { try { Pane.Delete(); } catch { } } throw; }
            }
            public void Dispose()
            {
                if (Disposed) return;
                Disposed = true;
                Model.Dispose();
                try { Pane?.Delete(); } catch (Exception ex) { Log.Error("Delete pane", ex); }
                control.Dispose();
            }
        }
    }
}
