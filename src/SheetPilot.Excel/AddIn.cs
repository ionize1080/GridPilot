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
        public void AutoOpen()
        {
            Features = new FeatureRegistry();
            Navigation = new NavigationModule(new ExcelGateway(() => ExcelDnaUtil.Application));
            Features.Add(Navigation);
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
                Features = null; Navigation = null; Log.Write("Unloaded");
                // Unreachable Office RCWs can otherwise survive in an idle automation host.
                // Do not block the Excel STA: COM finalizers may need to marshal back to it.
                System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                {
                    GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                });
            }
        }
        [ExcelCommand(Name = "SheetPilot_ShowNavigation", Description = "显示 SheetPilot 工作表导航")]
        public static void ShowNavigation() => Guard(() => Navigation?.Show());
        internal static void Guard(Action action)
        {
            try { action(); }
            catch (Exception ex)
            {
                Log.Error("Command", ex);
                System.Windows.Forms.MessageBox.Show("SheetPilot 无法完成此操作。\n" + ex.Message + "\n\n诊断日志：" + Log.Path,
                    "SheetPilot", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
        }
    }

    [ComVisible(true)]
    [Guid("74C45697-E294-4FA3-AF4B-2315D65F68C6")]
    public sealed class Ribbon : ExcelRibbon
    {
        public override void OnBeginShutdown(ref Array custom)
        {
            AddIn.Shutdown();
            base.OnBeginShutdown(ref custom);
        }
        public override void OnDisconnection(ExcelDna.Integration.Extensibility.ext_DisconnectMode mode, ref Array custom)
        {
            AddIn.Shutdown();
            base.OnDisconnection(mode, ref custom);
        }
        public override string GetCustomUI(string ribbonId)
        {
            XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
            var tab = new XElement(ns + "tab", new XAttribute("id", "SheetPilot_Tab"), new XAttribute("label", "SheetPilot"));
            if (AddIn.Features != null)
                foreach (var module in AddIn.Features.Modules)
                {
                    var group = new XElement(ns + "group", new XAttribute("id", module.Id), new XAttribute("label", module.Label));
                    foreach (var command in module.Commands)
                        group.Add(new XElement(ns + "button", new XAttribute("id", command.Id), new XAttribute("label", command.Label),
                            new XAttribute("size", "large"), new XAttribute("imageMso", command.ImageMso), new XAttribute("screentip", command.Label),
                            new XAttribute("supertip", command.Description), new XAttribute("onAction", "OnCommand")));
                    tab.Add(group);
                }
            return new XElement(ns + "customUI", new XElement(ns + "ribbon", new XElement(ns + "tabs", tab))).ToString();
        }
        public void OnCommand(IRibbonControl control) => AddIn.Guard(() => AddIn.Features?.Execute(control.Id));
    }

    internal static class Log
    {
        internal static string Path => System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SheetPilot-" + System.Diagnostics.Process.GetCurrentProcess().Id + ".log");
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
