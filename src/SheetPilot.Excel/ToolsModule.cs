using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Forms;
using ExcelDna.Integration;
using SheetPilot.Core;
using SheetPilot.UI;

namespace SheetPilot.Excel
{
    public sealed class ToolsModule : IFeatureModule
    {
        private readonly ExcelGateway gateway;
        private readonly NavigationModule navigation;
        private readonly ToolPreferencesStore store;
        private readonly ToolPreferences settings;
        private readonly FocusOverlay focus;
        private readonly HashSet<Window> windows = new HashSet<Window>();
        private readonly Timer statusTimer = new Timer { Interval = 9000 };
        private object previousStatus;
        private string status;
        private bool disposed;
        public ToolsModule(ExcelGateway gateway, NavigationModule navigation)
        {
            this.gateway = gateway; this.navigation = navigation;
            store = new ToolPreferencesStore(Path.Combine(Path.GetDirectoryName(ExcelDnaUtil.XllPath), "GridPilot.tools.json")); settings = store.Load();
            focus = new FocusOverlay(() => ExcelDnaUtil.Application, action => ExcelAsyncUtil.QueueAsMacro(() => action()), settings);
            statusTimer.Tick += (s,e) => ClearStatus();
        }
        public string Id => "GridPilot_Tools";
        public string Label => "常用工具";
        public bool FocusEnabled => focus.Enabled;
        public bool CanUndo => gateway.CanUndo;
        internal string PerformanceSummary => focus.PerformanceSummary;
        public IEnumerable<FeatureCommand> Commands => new[] {
            new FeatureCommand("GridPilot_CopyInternal", "原样复制·内部公式", "复制到新工作簿，保留排版和范围内公式；外部及无法确认依赖的公式按设置转为当前值或停止。Esc 可取消。", "", () => Run(() => CopySelection(true))),
            new FeatureCommand("GridPilot_CopyValues", "原样复制·仅数值", "复制到新工作簿，所有公式转为当前值，保留文本、错误值与排版。复制选项位于统一设置。Esc 可取消。", "", () => Run(() => CopySelection(false))),
            new FeatureCommand("GridPilot_Rename", "批量重命名", "选择范围、查找替换与编号；检查预览后应用。", "", () => Run(OpenRename)),
            new FeatureCommand("GridPilot_Focus", "聚焦", "临时高亮活动单元格所在行列；再次点击关闭，不修改单元格格式。", "", () => { focus.Toggle(); Announce(focus.Enabled ? "聚焦已开启，在普通工作表视图中跟随活动单元格。" : "聚焦已关闭。"); Ribbon.Refresh(); }),
            new FeatureCommand("GridPilot_Text", "转为文本", "将选中的数值常量转为文本，保留视觉位置与其他格式。", "", () => Run(() => Announce(gateway.ConvertSelection(true, settings)))),
            new FeatureCommand("GridPilot_Number", "转为数值", "将选中的数字文本转为数值，默认保护前导零和长编号。", "", () => Run(() => Announce(gateway.ConvertSelection(false, settings)))),
            new FeatureCommand("GridPilot_Undo", "撤销上次操作", "恢复最近一次成功的 GridPilot 重命名或内容转换；不覆盖已被再次修改的目标。", "", () => Run(() => { Announce(gateway.UndoLast()); navigation.Refresh(); })),
            new FeatureCommand("GridPilot_Settings", "设置", "原样复制、聚焦、转换与重命名默认规则；设置随加载项文件夹保存。", "", () => Run(OpenSettings)) };
        private void Run(Action action) { if (disposed) return; ExcelAsyncUtil.QueueAsMacro(() => { if (disposed) return; AddIn.Guard(action); Ribbon.Refresh(); }); }
        private void CopySelection(bool formulas)
        {
            try
            {
                var result = gateway.CopySelection(formulas, settings.Copy, (stage, done, total) =>
                { if ((GetAsyncKeyState(27) & 0x8000) != 0) throw new OperationCanceledException(); });
                Announce(result.Summary); Log.Write("Copy: cells=" + result.Cells + "; kept=" + result.KeptFormulas + "; values=" + result.ValueFormulas + "; ms=" + result.ElapsedMilliseconds);
                navigation.Refresh();
            }
            catch (OperationCanceledException) { Announce("原样复制已取消，未保留未完成的结果工作簿。"); }
        }
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        public void Execute(string id) { foreach (var command in Commands) if (command.Id == id) { command.Execute(); return; } }
        private void OpenRename()
        {
            var working = new ToolPreferences { Rename = settings.Rename.Clone() };
            Show(new RenameWindow(working, navigation.Settings, all => gateway.Capture(all), plan => { var result = gateway.ApplyRename(plan); navigation.Refresh(); Ribbon.Refresh(); return result; }, () => { settings.Rename = working.Rename.Clone(); store.Save(settings); return store.Status; }));
        }
        private void OpenSettings()
        {
            Show(new PreferencesWindow(settings, navigation.Settings, () => { store.Save(settings); return store.Status; }, navigation.ShowSettings));
            if (!store.Save(settings)) Announce(store.Status);
        }
        private void Show(Window window)
        {
            new WindowInteropHelper(window).Owner = ExcelDnaUtil.WindowHandle;
            windows.Add(window); try { window.ShowDialog(); } finally { windows.Remove(window); }
        }
        private void Announce(string text)
        {
            dynamic app = ExcelDnaUtil.Application;
            if (status == null) previousStatus = app.StatusBar;
            status = "GridPilot · " + text; app.StatusBar = status; statusTimer.Stop(); statusTimer.Start();
        }
        private void ClearStatus()
        {
            statusTimer.Stop(); if (status == null) return;
            try { dynamic app = ExcelDnaUtil.Application; if (Equals((object)app.StatusBar, status)) app.StatusBar = previousStatus ?? false; } catch { }
            status = null;
        }
        public void Dispose() { if (disposed) return; disposed = true; focus.Dispose(); ClearStatus(); statusTimer.Dispose(); foreach (var window in new List<Window>(windows)) window.Close(); windows.Clear(); }
    }
}
