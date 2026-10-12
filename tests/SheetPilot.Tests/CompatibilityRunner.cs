using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static partial class Program
{
    private static int CompatibilityRunner()
    {
        string root=AppDomain.CurrentDomain.BaseDirectory;string logs=Path.Combine(root,"results-"+DateTime.Now.ToString("yyyyMMdd-HHmmss"));Directory.CreateDirectory(logs);
        var form=new Form {Text="GridPilot compatibility test",Width=740,Height=280,StartPosition=FormStartPosition.CenterScreen};
        var label=new Label {Dock=DockStyle.Fill,Padding=new Padding(20),Text="Starting tests. Please do not use Excel until the tests finish.",AutoSize=false};form.Controls.Add(label);
        int failures=0;
        form.Shown+=async (s,e)=>
        {
            string addin=Path.Combine(root,Environment.Is64BitProcess?"GridPilot-x64.xll":"GridPilot-x86.xll");
            string[] names={"core","color-picker","release-features","excel-regression","copy-excel"};
            string[] arguments={"","--color-picker-host","--release-features --excel","--excel \""+addin+"\"","--copy-excel \""+addin+"\""};
            File.WriteAllText(Path.Combine(logs,"environment.txt"),"OS="+Environment.OSVersion+"\r\nProcess="+(Environment.Is64BitProcess?"x64":"x86")+"\r\nCLR="+Environment.Version+"\r\nUTC="+DateTime.UtcNow.ToString("o"));
            for(int i=0;i<names.Length;i++)
            {
                int test=i;label.Text="Running "+(i+1)+" / "+names.Length+": "+names[i]+"\r\nPlease keep Excel available.\r\nReports: "+logs;
                int code=await Task.Run(()=>
                {
                    using(var p=new Process {StartInfo=new ProcessStartInfo(Path.Combine(root,"SheetPilot.Tests.exe"),arguments[test]) {WorkingDirectory=root,UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true}})
                    {
                        p.Start();var stdout=p.StandardOutput.ReadToEndAsync();var stderr=p.StandardError.ReadToEndAsync();
                        if(!p.WaitForExit(300000)){p.Kill();File.WriteAllText(Path.Combine(logs,names[test]+".txt"),"TIMEOUT after 300s");return 2;}
                        Task.WaitAll(stdout,stderr);File.WriteAllText(Path.Combine(logs,names[test]+".txt"),stdout.Result+stderr.Result,Encoding.UTF8);return p.ExitCode;
                    }
                });
                if(code!=0)failures++;
            }
            File.WriteAllText(Path.Combine(logs,"summary.txt"),failures==0?"ALL SUITES PASSED":"FAILED SUITES: "+failures);
            label.Text=(failures==0?"ALL SUITES PASSED":"FAILED SUITES: "+failures)+"\r\nReports: "+logs+"\r\nYou may close this window.";
        };
        Application.Run(form);return failures==0?0:1;
    }
}
