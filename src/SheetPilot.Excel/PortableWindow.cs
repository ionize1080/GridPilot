using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
namespace SheetPilot.Excel
{
    // Per-window named pipe restricted to current Windows user. No TCP transport.
    internal sealed class PortableWindow : IDisposable
    {
        private readonly Process process;
        private readonly NamedPipeServerStream pipe;
        private StreamWriter writer;
        private readonly Func<JObject,Task<object>> handle;
        private readonly object gate=new object();
        private volatile bool disposed,disconnected;
        private int lastProgressTick;
        private readonly SemaphoreSlim sending=new SemaphoreSlim(1,1);
        public bool IsDisconnected=>disposed || disconnected;
        public void SetAiMode(bool sidebar) => Send(new { type = "mode", mode = sidebar ? "sidebar" : "window" });
        public void Activate() => Send(new { type = "activate" });
        public event Action Disconnected;
        private void Disconnect(){lock(gate){if(disconnected)return;disconnected=true;}Disconnected?.Invoke();CancelRequested?.Invoke();}
        public event Action Closed;
        public event Action CancelRequested;
        public PortableWindow(string directory,Func<JObject,Task<object>> handler,string appFolder="App")
        {
            handle=handler;string executable=Path.Combine(directory,"Browser","electron.exe");
            if(!File.Exists(executable))throw new InvalidOperationException("缺少便携界面组件。请完整解压发布包（包含 Browser 和 App 文件夹）。");
            string name="GridPilot-"+Guid.NewGuid().ToString("N");
            var acl=new PipeSecurity();acl.SetAccessRuleProtection(true,false);acl.AddAccessRule(new PipeAccessRule(WindowsIdentity.GetCurrent().User,PipeAccessRights.FullControl,AccessControlType.Allow));
            pipe=new NamedPipeServerStream(name,PipeDirection.InOut,1,PipeTransmissionMode.Byte,PipeOptions.Asynchronous,65536,65536,acl);
            var start=new ProcessStartInfo(executable,"\""+Path.Combine(directory,appFolder)+"\" --gridpilot-pipe="+name) {UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=directory};
            start.EnvironmentVariables.Remove("ELECTRON_RUN_AS_NODE");start.EnvironmentVariables.Remove("NODE_OPTIONS");
            process=new Process {StartInfo=start,EnableRaisingEvents=true};
            process.Exited+=(s,e)=>{Disconnect();if(!disposed)ExcelAsyncUtil.QueueAsMacro(()=>{Closed?.Invoke();Dispose();});};
            try{process.Start();_ = ReadMessages();}catch{pipe.Dispose();process.Dispose();throw;}
        }
        private async Task ReadMessages()
        {
            try
            {
                await pipe.WaitForConnectionAsync().ConfigureAwait(false);
                writer=new StreamWriter(pipe,new UTF8Encoding(false),65536,true){AutoFlush=true};
                using(var reader=new StreamReader(pipe,Encoding.UTF8,false,65536,true))
                {string line;while(!disposed && (line=await reader.ReadLineAsync().ConfigureAwait(false))!=null)Receive(line);}
            }
            catch(Exception ex) when(ex is IOException || ex is ObjectDisposedException){ }
            finally{Disconnect();}
        }
        private void Receive(string line)
        {
            if(IsDisconnected || string.IsNullOrEmpty(line) || line.Length>32*1024*1024)return;
            JObject message;try{message=JObject.Parse(line);}catch(JsonException){return;}
            if((string)message["action"]=="cancel"){CancelRequested?.Invoke();Send(new {id=(int?)message["id"],data=new {message="正在取消…"}});return;}
            ExcelAsyncUtil.QueueAsMacro(()=>
            {
                if(IsDisconnected)return;
                Task<object> response;
                var previous=SynchronizationContext.Current;
                try{SynchronizationContext.SetSynchronizationContext(new MacroContext());response=handle(message);}
                catch(Exception ex){Send(new {id=(int?)message["id"],error=ex.Message});Log.Error("Portable UI request",ex);return;}
                finally{SynchronizationContext.SetSynchronizationContext(previous);}
                _ = response.ContinueWith(task=>
                {
                    if(task.IsCanceled)Send(new {id=(int?)message["id"],error="已取消。"});
                    else if(task.IsFaulted){var ex=task.Exception.GetBaseException();Send(new {id=(int?)message["id"],error=ex.Message});Log.Error("Portable UI request",ex);}
                    else Send(new {id=(int?)message["id"],data=task.Result});
                },TaskScheduler.Default);
            });
        }
        private sealed class MacroContext : SynchronizationContext
        {
            public override void Post(SendOrPostCallback callback,object state)=>ExcelAsyncUtil.QueueAsMacro(()=>
            {var previous=Current;try{SetSynchronizationContext(this);callback(state);}finally{SetSynchronizationContext(previous);}});
        }
        public void SendProgress(string phase,int percent,long current=0,long total=0)
        {
            int now=Environment.TickCount;lock(gate){if(unchecked(now-lastProgressTick)<80 && percent<100)return;lastProgressTick=now;}
            Send(new {type="progress",phase,percent=Math.Max(0,Math.Min(100,percent)),current,total});
        }
        private void Send(object message){if(IsDisconnected)return;_ = SendAsync(JsonConvert.SerializeObject(message));}
        private async Task SendAsync(string json)
        {
            await sending.WaitAsync().ConfigureAwait(false);
            try{if(!IsDisconnected && writer!=null)await writer.WriteLineAsync(json).ConfigureAwait(false);}
            catch(Exception ex) when(ex is IOException || ex is ObjectDisposedException){Disconnect();}
            finally{sending.Release();}
        }
        public void Dispose()
        {
            if(disposed)return;disposed=true;Disconnect();pipe.Dispose();
            // Process cleanup must not block the host Excel message pump.
            _ = Task.Run(()=>{try{if(!process.HasExited && !process.WaitForExit(500))process.Kill();}catch(InvalidOperationException){}finally{process.Dispose();}});
        }
    }
}
