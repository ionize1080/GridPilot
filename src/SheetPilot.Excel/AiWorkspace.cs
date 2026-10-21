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
        private IntPtr child,owner;
        private int originalStyle,originalExtendedStyle;
        [StructLayout(LayoutKind.Sequential)] private struct WindowRect { public int left,top,right,bottom; }
        private WindowRect floatingBounds;
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd,out WindowRect rect);
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr hwnd,int attribute,ref int value,int size);
        private static void SetFrame(IntPtr hwnd, bool docked) { int policy=docked?1:0; DwmSetWindowAttribute(hwnd,2,ref policy,4); }
        private readonly Timer tracking=new Timer{Interval=120};
        private System.Drawing.Rectangle lastBounds;
        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr window,int index);
        [DllImport("user32.dll")] private static extern int SetWindowLong(IntPtr window,int index,int value);
        [DllImport("user32.dll",EntryPoint="SetWindowLongPtrW")] private static extern IntPtr SetWindowLongPtr(IntPtr window,int index,IntPtr value);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr window,IntPtr after,int x,int y,int width,int height,uint flags);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsZoomed(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window,int command);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window,uint flags);
        private static void Owner(IntPtr window,IntPtr owner){if(IntPtr.Size==8)SetWindowLongPtr(window,-8,owner);else SetWindowLong(window,-8,owner.ToInt32());}
        internal IntPtr HostWindow => IsHandleCreated?GetAncestor(Handle,2):IntPtr.Zero;
        public AiPane()
        {
            AutoScaleMode=AutoScaleMode.Dpi;BackColor=System.Drawing.Color.White;
            tracking.Tick+=(s,e)=>Fit();
        }
        internal void Attach(IntPtr window)
        {
            if(child==window){Fit();return;}
            if(IsIconic(window)||IsZoomed(window))ShowWindow(window,9);
            child=window;originalStyle=GetWindowLong(child,-16);originalExtendedStyle=GetWindowLong(child,-20);
            GetWindowRect(child,out floatingBounds);
            // Keep Chromium as an owned top-level HWND. SetParent/WS_CHILD lets
            // Excel's accelerator loop consume text before Chromium receives it.
            owner=GetAncestor(Handle,2);Owner(child,owner);
            SetWindowLong(child,-16,(originalStyle & ~0x00C40000 & ~0x40000000)|unchecked((int)0x80000000));
            SetWindowLong(child,-20,originalExtendedStyle & ~0x00000301);SetFrame(child,true);
            lastBounds=System.Drawing.Rectangle.Empty;Fit();tracking.Start();
        }
        internal void Detach()
        {
            tracking.Stop();if(child==IntPtr.Zero)return;
            Owner(child,IntPtr.Zero);SetWindowLong(child,-16,originalStyle);SetWindowLong(child,-20,originalExtendedStyle);SetFrame(child,false);
            var window=child;child=IntPtr.Zero;
            SetWindowPos(window,IntPtr.Zero,floatingBounds.left,floatingBounds.top,Math.Max(320,floatingBounds.right-floatingBounds.left),Math.Max(240,floatingBounds.bottom-floatingBounds.top),0x34);ShowWindow(window,4);
        }
        private void Fit()
        {
            if(child==IntPtr.Zero||!IsHandleCreated)return;
            if(!Visible||!IsWindowVisible(Handle)||IsIconic(owner)){ShowWindow(child,0);return;}
            var bounds=RectangleToScreen(ClientRectangle);
            if(bounds.Width<20||bounds.Height<20){ShowWindow(child,0);return;}
            if(bounds!=lastBounds){SetWindowPos(child,IntPtr.Zero,bounds.X,bounds.Y,bounds.Width,bounds.Height,0x34);lastBounds=bounds;}
            ShowWindow(child,4);
        }
        protected override void OnResize(EventArgs e){base.OnResize(e);Fit();}
        protected override void OnVisibleChanged(EventArgs e){base.OnVisibleChanged(e);Fit();}
        protected override void Dispose(bool disposing){if(disposing){Detach();tracking.Dispose();}base.Dispose(disposing);}
    }
    public sealed partial class ToolsModule
    {
        private PortableWindow aiWindow;
        private readonly System.Threading.SemaphoreSlim aiRequests=new System.Threading.SemaphoreSlim(1,1);
        private AiOperations ai;
        private CustomTaskPane aiPane;
        private AiPane aiControl;
        private IntPtr aiHandle;
        private bool aiDocked;
        private void ApplyAiSettings()
        {
            if(!settings.AiEnabled){CloseAi();return;}
            if(aiWindow==null || aiHandle==IntPtr.Zero)return;
            SetAiMode(settings.AiDisplayMode==0);
        }
        private void SetAiMode(bool sidebar)
        {
            if(sidebar&&aiPane!=null&&(aiControl==null||aiControl.IsDisposed||aiControl.HostWindow!=FocusOverlay.ActiveWindowHandle(ExcelDnaUtil.Application)))
            {
                aiControl?.Detach();aiDocked=false;
                try{aiPane.Delete();}catch(Exception ex){Log.Error("Release old AI pane",ex);}
                aiPane=null;aiControl?.Dispose();aiControl=null;
            }
            if(sidebar && (aiPane==null||aiControl==null||aiControl.IsDisposed))
            {
                aiControl=new AiPane();aiPane=CustomTaskPaneFactory.CreateCustomTaskPane(aiControl,"GridPilot AI");
                aiDocked=false;
                aiPane.DockPosition=MsoCTPDockPosition.msoCTPDockPositionRight;
                using(var g=aiControl.CreateGraphics())aiPane.Width=(int)(440*g.DpiX/96.0);
            }
            if(sidebar)
            {
                aiPane.Visible=true;if(!aiDocked){aiControl.Attach(aiHandle);aiDocked=true;}
            }
            else if(aiPane!=null){if(aiDocked){aiControl.Detach();aiDocked=false;}aiPane.Visible=false;}
            aiWindow?.SetAiMode(sidebar);
        }
        private System.Threading.Mutex aiProfileLock;
        private void OpenAi()
        {
            if (aiWindow != null && !aiWindow.IsDisconnected)
            {
                if (aiDocked) SetAiMode(true);
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
                aiWindow = new PortableWindow(directory, async message =>
                {
                    await aiRequests.WaitAsync();
                    try {
                    string action = (string)message["action"];
                    object result;
                    switch (action)
                    {
                        case "aiReady":
                            long handle=long.Parse((string)message["handle"]);aiHandle=IntPtr.Size==4?new IntPtr(unchecked((int)handle)):new IntPtr(handle);
                            SetAiMode(settings.AiDisplayMode==0);
                            result = new { mode = settings.AiDisplayMode==0 ? "sidebar" : "window", version = "1.20.3" }; break;
                        case "aiMode":
                            bool dock=(string)message["mode"]=="sidebar";SetAiMode(dock);settings.AiDisplayMode=dock?0:1;store.Save(settings);
                            result=new {mode=settings.AiDisplayMode==0?"sidebar":"window",message="已切换工作方式"};break;
                        case "aiApplyBatch": result=ai.ApplyBatch(message["proposalIds"]?.ToObject<string[]>());break;
                        case "aiApproval": result = ai.Approval((bool?)message["enabled"] == true,(string)message["scope"]??"session"); break;
                        case "aiConversation": result=ai.ConversationChanged();break;
                        case "aiDiscard": result=ai.Discard(message["proposalIds"]?.ToObject<string[]>());break;
                        case "aiCheckpoint": result = ai.CheckpointState(); break;
                        case "aiRestore": result = ai.RestoreCheckpoint((string)message["position"] == "before"); break;
                        case "aiContext": var state=JObject.FromObject(ai.Context());state["checkpoint"]=JToken.FromObject(ai.CheckpointState());result=state;break;
                        case "aiOpenRecovery": var recovery=(string)JObject.FromObject(ai.CheckpointState())["folder"];if(recovery==null||!Directory.Exists(recovery))throw new InvalidOperationException("尚无完整备份。");System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("explorer.exe", "\"" + recovery + "\""){UseShellExecute=false});result=new{message="已打开本地完整备份"};break;
                        case "aiCapture": result = await ai.CaptureAsync((string)message["scope"]); break;
                        case "aiInspect": result=ai.Metadata((string)message["snapshotId"]);break;
                        case "aiRead": result = ai.Read(message); break;
                        case "aiPropose": result = ai.Propose(message); break;
                        case "aiApply": result = ai.Apply((string)message["proposalId"]); break;
                        case "aiUndo": result = ai.Apply(null, true); break;
                        case "aiNavigate": result = ai.Navigate((string)message["snapshotId"], (string)message["sheet"], (string)message["address"]); break;
                        default: throw new InvalidOperationException("未知AI操作。");
                    }
                    return result;
                    } finally { aiRequests.Release(); }
                }, "Chatbox");
                aiWindow.CancelRequested += () => { if(ai!=null)ai.CancelCapture=true; };
                ai.Progress = (phase,percent,current,total) => aiWindow?.SendProgress(phase,percent,current,total);
                ai.YieldCapture = async () => { await Task.Delay(1); };
                portableWindows.Add(aiWindow);
                aiWindow.Closed += CloseAi;
            }
            catch { CloseAi(); throw; }
        }
        private void CloseAi()
        {
            if (aiWindow != null) { aiWindow.Closed -= CloseAi; portableWindows.Remove(aiWindow); aiWindow.Dispose(); aiWindow = null; }
            ai?.Dispose(); ai = null;
            aiHandle=IntPtr.Zero;
            aiDocked=false;
            if (aiPane != null) { try { aiPane.Delete(); } catch { } aiPane = null; }
            aiControl?.Dispose(); aiControl = null;
            if (aiProfileLock != null) { aiProfileLock.ReleaseMutex(); aiProfileLock.Dispose(); aiProfileLock = null; }
        }
    }
}
