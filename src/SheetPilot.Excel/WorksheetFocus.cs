using System;
using System.Runtime.InteropServices;
using System.Text;
namespace SheetPilot.Excel
{
    public static class WorksheetFocus
    {
        public static void Restore(IntPtr owner)
        {
            // Office 2010 has one XLMAIN with multiple EXCEL7 children. Only the visible grid
            // of the active workbook may receive focus; no global hooks or synthetic shortcuts.
            if(owner==IntPtr.Zero || !IsWindow(owner))return;
            IntPtr grid=FindGrid(owner);
            if(grid==IntPtr.Zero)return;
            SetForegroundWindow(owner);
            System.Windows.Input.Keyboard.ClearFocus();
            SetFocus(grid);
        }
        internal static IntPtr FindGrid(IntPtr owner)
        {
            // Use the active MDI child when present; enumeration order/area can select another workbook.
            IntPtr mdi=IntPtr.Zero;
            EnumChildWindows(owner,(child,param)=>{var name=new StringBuilder(64);GetClassName(child,name,64);if(name.ToString()=="MDIClient"){mdi=child;return false;}return true;},IntPtr.Zero);
            IntPtr root=mdi==IntPtr.Zero?owner:SendMessage(mdi,0x229,IntPtr.Zero,IntPtr.Zero);if(root==IntPtr.Zero)root=owner;
            IntPtr grid=IntPtr.Zero;
            EnumChildWindows(root,(child,param)=>{var name=new StringBuilder(64);GetClassName(child,name,64);if(name.ToString()=="EXCEL7" && IsWindowVisible(child)){grid=child;return false;}return true;},IntPtr.Zero);
            return grid;
        }
        [DllImport("user32.dll")] private static extern IntPtr SendMessage(IntPtr window,uint message,IntPtr wparam,IntPtr lparam);
        private delegate bool EnumWindowsProc(IntPtr hwnd,IntPtr param);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent,EnumWindowsProc callback,IntPtr param);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd,StringBuilder text,int count);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hwnd);
    }
}
