using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed class AddIn : IExcelAddIn
    {
        internal static FeatureRegistry Features;
        internal static NavigationModule Navigation;
        internal static ToolsModule Tools;
        public void AutoOpen()
        {
            Features = new FeatureRegistry();
            var gateway = new ExcelGateway(() => ExcelDnaUtil.Application);
            Navigation = new NavigationModule(gateway);
            Features.Add(Navigation);
            Tools = new ToolsModule(gateway, Navigation); Features.Add(Tools);
            Log.Write("Loaded " + (Environment.Is64BitProcess ? "x64" : "x86") + "; Excel " + ExcelDnaUtil.ExcelVersion);
        }
        public void AutoClose() => Shutdown();
        internal static void Shutdown()
        {
            if (Features == null) return;
            try { Features?.Dispose(); }
            catch (Exception ex) { Log.Error("Unload", ex); }
            finally
            {
                Features = null; Navigation = null; Tools = null; Log.Write("Unloaded");
                // Unreachable Office RCWs can otherwise survive in an idle automation host.
                // Do not block the Excel STA: COM finalizers may need to marshal back to it.
                System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                {
                    GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                });
            }
        }
        [ExcelCommand(Name = "SheetPilot_ShowNavigation", Description = "显示 GridPilot 工作表导航")]
        public static void ShowNavigation() => Guard(() => Navigation?.Show());
        [ExcelCommand(Name = "GridPilot_ShowNavigation")] public static void GridPilotNavigation() => ShowNavigation();
        [ExcelCommand(Name = "GridPilot_Rename")] public static void Rename() => Guard(() => Tools?.Execute("GridPilot_Rename"));
        [ExcelCommand(Name = "GridPilot_Focus")] public static void Focus() => Guard(() => Tools?.Execute("GridPilot_Focus"));
        [ExcelCommand(Name = "GridPilot_Text")] public static void Text() => Guard(() => Tools?.Execute("GridPilot_Text"));
        [ExcelCommand(Name = "GridPilot_Number")] public static void Number() => Guard(() => Tools?.Execute("GridPilot_Number"));
        [ExcelCommand(Name = "GridPilot_Segment")] public static void Segment() => Guard(() => Tools?.Execute("GridPilot_Segment"));
        [ExcelCommand(Name = "GridPilot_Database")] public static void Database() => Guard(() => Tools?.Execute("GridPilot_Database"));
        [ExcelCommand(Name = "GridPilot_Mask")] public static void Mask() => Guard(() => Tools?.Execute("GridPilot_Mask"));
        [ExcelCommand(Name = "GridPilot_Undo")] public static void Undo() => Guard(() => Tools?.Execute("GridPilot_Undo"));
        [ExcelCommand(Name = "GridPilot_Settings")] public static void Settings() => Guard(() => Tools?.Execute("GridPilot_Settings"));
        [ExcelCommand(Name = "GridPilot_CopyInternal")] public static void CopyInternal() => Guard(() => Tools?.Execute("GridPilot_CopyInternal"));
        [ExcelCommand(Name = "GridPilot_CopyValues")] public static void CopyValues() => Guard(() => Tools?.Execute("GridPilot_CopyValues"));
        [ExcelCommand(Name = "GridPilot_Diagnostics")] public static void Diagnostics() => Log.Write(Tools?.PerformanceSummary ?? "Tools unavailable");
        internal static void Guard(Action action)
        {
            try { action(); }
            catch (Exception ex)
            {
                Log.Error("Command", ex);
                System.Windows.Forms.MessageBox.Show("GridPilot 无法完成此操作。\n" + ex.Message + "\n\n诊断日志：" + Log.Path,
                    "GridPilot", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
        }
    }

    [ComVisible(true)]
    [Guid("74C45697-E294-4FA3-AF4B-2315D65F68C6")]
    public sealed class Ribbon : ExcelRibbon
    {
        private static IRibbonUI ui;
        public void OnLoad(IRibbonUI ribbon) { ui = ribbon; }
        public static void Refresh() { try { ui?.Invalidate(); } catch { } }
        public override object LoadImage(string imageId) => RibbonIcons.Get(imageId, ExcelDnaUtil.WindowHandle);
        public object GetImage(IRibbonControl control) => RibbonIcons.Picture(control.Id, ExcelDnaUtil.WindowHandle, AddIn.Tools?.IconTheme ?? 0);
        public bool GetPressed(IRibbonControl control) => AddIn.Tools?.FocusEnabled == true;
        public bool GetEnabled(IRibbonControl control) => control.Id != "GridPilot_Undo" || AddIn.Tools?.CanUndo == true;
        public void OnToggle(IRibbonControl control, bool pressed) => OnCommand(control);

        public override void OnBeginShutdown(ref Array custom)
        {
            AddIn.Shutdown();
            base.OnBeginShutdown(ref custom);
        }
        public override void OnDisconnection(ExcelDna.Integration.Extensibility.ext_DisconnectMode mode, ref Array custom)
        {
            AddIn.Shutdown();
            ui = null; RibbonIcons.Release();
            base.OnDisconnection(mode, ref custom);
        }
        public override string GetCustomUI(string ribbonId)
        {
            XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
            var tab = new XElement(ns + "tab", new XAttribute("id", "SheetPilot_Tab"), new XAttribute("label", "GridPilot"));
            if (AddIn.Features != null)
                foreach (var module in AddIn.Features.Modules)
                {
                    var group = new XElement(ns + "group", new XAttribute("id", module.Id), new XAttribute("label", module.Label));
                    foreach (var command in module.Commands)
                    {
                        bool toggle = command.Id == "GridPilot_Focus";
                        var button = new XElement(ns + (toggle ? "toggleButton" : "button"), new XAttribute("id", command.Id), new XAttribute("label", command.Label),
                            new XAttribute("size", "large"), new XAttribute("getImage", "GetImage"), new XAttribute("screentip", command.Label),
                            new XAttribute("supertip", command.Description), new XAttribute("getEnabled", "GetEnabled"), new XAttribute("onAction", toggle ? "OnToggle" : "OnCommand"));
                        if (toggle) button.Add(new XAttribute("getPressed", "GetPressed"));
                        group.Add(button);
                    }
                    tab.Add(group);
                }
            return new XElement(ns + "customUI", new XAttribute("loadImage", "LoadImage"), new XAttribute("onLoad", "OnLoad"), new XElement(ns + "ribbon", new XElement(ns + "tabs", tab))).ToString();
        }
        public void OnCommand(IRibbonControl control) => AddIn.Guard(() => AddIn.Features?.Execute(control.Id));
    }

    internal static class Log
    {
        internal static string Path => System.IO.Path.Combine(System.IO.Path.GetTempPath(), "GridPilot-" + System.Diagnostics.Process.GetCurrentProcess().Id + ".log");
        internal static void Write(string message)
        {
            try
            {
                if (File.Exists(Path) && new FileInfo(Path).Length > 1024 * 1024) File.WriteAllText(Path, "");
                File.AppendAllText(Path, DateTime.Now.ToString("s") + " " + message + Environment.NewLine);
            }
            catch { /* Logging must never break Excel or require a writable add-in folder. */ }
        }
        internal static void Error(string operation, Exception error) => Write(operation + ": " + error.GetType().Name + " (0x" + error.HResult.ToString("X8") + ")");
    }
}
