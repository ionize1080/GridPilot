using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    [ComVisible(true), Guid("62A071C0-F294-45D6-83BD-97A2244406E1"), ProgId("GridPilot.AiPane")]
    public sealed class AiPane : UserControl
    {
        private IntPtr child;
        [DllImport("user32.dll", SetLastError=true)] private static extern IntPtr SetParent(IntPtr window, IntPtr parent);
        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr window, int index);
        [DllImport("user32.dll")] private static extern int SetWindowLong(IntPtr window, int index, int value);
        [DllImport("user32.dll")] private static extern bool MoveWindow(IntPtr window, int x, int y, int width, int height, bool repaint);
        public AiPane() { AutoScaleMode = AutoScaleMode.Dpi; BackColor = System.Drawing.Color.White; }
        internal void Attach(IntPtr window)
        {
            child = window;
            int style = GetWindowLong(child, -16);
            SetWindowLong(child, -16, (style & ~unchecked((int)0x80000000) & ~0x00C00000) | 0x40000000);
            SetParent(child, Handle);
            Fit();
        }
        private void Fit() { if (child != IntPtr.Zero) MoveWindow(child, 0, 0, ClientSize.Width, ClientSize.Height, true); }
        protected override void OnResize(EventArgs e) { base.OnResize(e); Fit(); }
    }
    public sealed partial class ToolsModule
    {
        private PortableWindow aiWindow;
        private AiOperations ai;
        private CustomTaskPane aiPane;
        private AiPane aiControl;
        private System.Threading.Mutex aiProfileLock;
        private void OpenAi()
        {
            if (aiWindow != null && !aiWindow.IsDisconnected)
            {
                if (aiPane != null) aiPane.Visible = true;
                aiWindow.Activate();
                Announce("AI助手已打开；可在任务栏或Excel侧边栏继续对话。"); return;
            }
            if (aiWindow != null) CloseAi();
            string directory = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            if (!File.Exists(Path.Combine(directory, "Chatbox", "bootstrap.cjs")))
                throw new InvalidOperationException("缺少Chatbox 1.20.3组件，请完整解压新版本发布包。");
            aiProfileLock = new System.Threading.Mutex(false, @"Local\GridPilot-Chatbox-" + System.Security.Principal.WindowsIdentity.GetCurrent().User.Value);
            bool locked;
            try { locked = aiProfileLock.WaitOne(0); } catch (System.Threading.AbandonedMutexException) { locked = true; }
            if (!locked) { aiProfileLock.Dispose(); aiProfileLock = null; throw new InvalidOperationException("另一Excel进程已打开AI助手。请先关闭那里的Chatbox，再在此进程打开；同一Excel进程可直接切换工作簿。"); }
            ai = new AiOperations(() => ExcelDnaUtil.Application);
            bool sidebar = settings.AiDisplayMode == 0;
            try
            {
                if (sidebar)
                {
                    aiControl = new AiPane();
                    aiPane = CustomTaskPaneFactory.CreateCustomTaskPane(aiControl, "GridPilot AI · Chatbox 1.20.3");
                    aiPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                    using (var graphics = aiControl.CreateGraphics()) aiPane.Width = (int)(480 * graphics.DpiX / 96.0);
                    aiPane.Visible = true;
                }
                aiWindow = new PortableWindow(directory, message =>
                {
                    string action = (string)message["action"];
                    object result;
                    switch (action)
                    {
                        case "aiReady":
                            if (sidebar) aiControl.Attach(new IntPtr(long.Parse((string)message["handle"])));
                            result = new { mode = sidebar ? "sidebar" : "window", version = "1.20.3" }; break;
                        case "aiContext": result = ai.Context(); break;
                        case "aiCapture": result = ai.Capture((string)message["scope"]); break;
                        case "aiRead": result = ai.Read(message); break;
                        case "aiPropose": result = ai.Propose(message); break;
                        case "aiApply": result = ai.Apply((string)message["proposalId"]); break;
                        case "aiUndo": result = ai.Apply(null, true); break;
                        case "aiNavigate": result = ai.Navigate((string)message["snapshotId"], (string)message["sheet"], (string)message["address"]); break;
                        default: throw new InvalidOperationException("未知AI操作。");
                    }
                    return Task.FromResult(result);
                }, "Chatbox");
                portableWindows.Add(aiWindow);
                aiWindow.Closed += CloseAi;
            }
            catch { CloseAi(); throw; }
        }
        private void CloseAi()
        {
            if (aiWindow != null) { aiWindow.Closed -= CloseAi; portableWindows.Remove(aiWindow); aiWindow.Dispose(); aiWindow = null; }
            ai?.Dispose(); ai = null;
            if (aiPane != null) { try { aiPane.Delete(); } catch { } aiPane = null; }
            aiControl?.Dispose(); aiControl = null;
            if (aiProfileLock != null) { aiProfileLock.ReleaseMutex(); aiProfileLock.Dispose(); aiProfileLock = null; }
        }
    }
}
