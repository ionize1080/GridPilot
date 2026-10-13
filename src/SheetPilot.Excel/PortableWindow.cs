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
    // Per-window named pipe restricted to current Windows user. No TCP or cell data on disk.
    internal sealed class PortableWindow : IDisposable
    {
        private readonly Process process;
        private readonly NamedPipeServerStream pipe;
        private StreamWriter writer;
        private readonly Func<JObject,Task<object>> handle;
        private readonly object gate=new object();
        private bool disposed;
        public event Action Closed;
        public event Action CancelRequested;
        public PortableWindow(string directory,Func<JObject,Task<object>> handler)
        {
            handle=handler;string executable=Path.Combine(directory,"Browser","electron.exe");
            if(!File.Exists(executable))throw new InvalidOperationException("缺少便携界面组件。请完整解压发布包（包含 Browser 和 App 文件夹）。");
            string name="GridPilot-"+Guid.NewGuid().ToString("N");
            var acl=new PipeSecurity();acl.SetAccessRuleProtection(true,false);acl.AddAccessRule(new PipeAccessRule(WindowsIdentity.GetCurrent().User,PipeAccessRights.FullControl,AccessControlType.Allow));
            pipe=new NamedPipeServerStream(name,PipeDirection.InOut,1,PipeTransmissionMode.Byte,PipeOptions.Asynchronous,65536,65536,acl);
            var start=new ProcessStartInfo(executable,"\""+Path.Combine(directory,"App")+"\" --gridpilot-pipe="+name) {UseShellExecute=false,CreateNoWindow=true,WorkingDirectory=directory};
            start.EnvironmentVariables.Remove("ELECTRON_RUN_AS_NODE");start.EnvironmentVariables.Remove("NODE_OPTIONS");
            process=new Process {StartInfo=start,EnableRaisingEvents=true};
            process.Exited+=(s,e)=>{if(!disposed)ExcelAsyncUtil.QueueAsMacro(()=>{Closed?.Invoke();Dispose();});};
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
            finally{CancelRequested?.Invoke();}
        }
        private void Receive(string line)
        {
            if(disposed || string.IsNullOrEmpty(line) || line.Length>4*1024*1024)return;
            JObject message;try{message=JObject.Parse(line);}catch(JsonException){return;}
            if((string)message["action"]=="cancel")CancelRequested?.Invoke();
            ExcelAsyncUtil.QueueAsMacro(()=>
            {
                if(disposed)return;
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
        private void Send(object message)
        {lock(gate){if(disposed)return;try{writer?.WriteLine(JsonConvert.SerializeObject(message));}catch(IOException){}}}
        public void Dispose(){if(disposed)return;disposed=true;lock(gate){pipe.Dispose();}process.Dispose();}
    }
}
