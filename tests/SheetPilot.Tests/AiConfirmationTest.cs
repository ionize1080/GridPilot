using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;

internal static partial class Program
{
    [DllImport("user32.dll")] private static extern IntPtr GetDlgItem(IntPtr hwnd,int id);
    [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr hwnd,int index);
    [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr hwnd,uint message,IntPtr wParam,IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hwnd,AiEnumWindow callback,IntPtr arg);
    private static void AiConfirmationTest(uint pid,bool yes)
    {
        var clock=Stopwatch.StartNew();IntPtr dialog=IntPtr.Zero;
        while(dialog==IntPtr.Zero&&clock.ElapsedMilliseconds<20000){AiEnumTop((hwnd,arg)=>{AiWindowProcess(hwnd,out uint owner);var title=new StringBuilder(200);AiText(hwnd,title,200);if(owner==pid&&title.ToString()=="确认自动批准"){dialog=hwnd;return false;}return true;},IntPtr.Zero);if(dialog==IntPtr.Zero)Thread.Sleep(100);}
        Check(dialog!=IntPtr.Zero,"native confirmation appears in owned Excel process");
        var text=new StringBuilder();EnumChildWindows(dialog,(hwnd,arg)=>{var part=new StringBuilder(2000);AiText(hwnd,part,2000);text.Append(part);return true;},IntPtr.Zero);
        Check(text.ToString().Contains("AI 可能出错")&&text.ToString().Contains("500 MB")&&text.ToString().Contains("其他工作簿不会获得授权"),"dialog explains risk, retention and workbook scope");
        Check((GetWindowLong(GetDlgItem(dialog,7),-16)&15)==1,"No is the default confirmation button");
        SendMessage(dialog,0x111,new IntPtr(yes?6:7),IntPtr.Zero);
        Console.WriteLine("PASS native confirmation "+(yes?"accepted":"cancelled"));
    }
}
