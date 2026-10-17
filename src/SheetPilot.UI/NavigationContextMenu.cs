using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace SheetPilot.UI
{
    // Office supplies the message loop. A native menu avoids WPF Popup capture /
    // activation dependencies in a cold ElementHost (notably Excel 2010).
    internal sealed class NavigationContextMenu
    {
        internal sealed class Entry
        {
            internal string Text;
            internal Action Execute;
            internal bool Enabled = true;
        }

        public bool IsOpen { get; private set; }
        public void Close() { if (IsOpen) EndMenu(); }

        public void Show(FrameworkElement target, bool keyboard, IList<Entry> entries)
        {
            if (IsOpen || !(PresentationSource.FromVisual(target) is HwndSource source)) return;
            IntPtr owner = GetAncestor(source.Handle, 2); // GA_ROOT; never reparent Office windows.
            if (owner == IntPtr.Zero) return;
            IntPtr menu = CreatePopupMenu();
            if (menu == IntPtr.Zero) throw new Win32Exception();
            uint command = 0;
            try
            {
                for (int i = 0; i < entries.Count; i++)
                {
                    var entry = entries[i];
                    uint flags = entry == null ? 0x800u : entry.Enabled ? 0u : 1u;
                    if (!AppendMenu(menu, flags, new UIntPtr((uint)i + 1), entry?.Text)) throw new Win32Exception();
                }
                POINT point;
                if (keyboard)
                {
                    var location = target.PointToScreen(new Point(0, target.ActualHeight));
                    point = new POINT { X = (int)location.X, Y = (int)location.Y };
                }
                else if (!GetCursorPos(out point)) return;
                IsOpen = true;
                SetForegroundWindow(owner);
                // Native tracking owns capture, Escape, click-away dismissal and keyboard navigation.
                // RETURNCMD | NONOTIFY | RIGHTBUTTON: do not send our command IDs to Excel.
                command = TrackPopupMenuEx(menu, 0x100 | 0x80 | 0x2, point.X, point.Y, owner, IntPtr.Zero);
                PostMessage(owner, 0, IntPtr.Zero, IntPtr.Zero);
            }
            finally { IsOpen = false; DestroyMenu(menu); }
            if (command > 0 && command <= entries.Count && entries[(int)command - 1]?.Enabled == true)
                entries[(int)command - 1].Execute?.Invoke();
        }

        [StructLayout(LayoutKind.Sequential)] private struct POINT { public int X, Y; }
        [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr CreatePopupMenu();
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool AppendMenu(IntPtr menu, uint flags, UIntPtr id, string text);
        [DllImport("user32.dll")] private static extern uint TrackPopupMenuEx(IntPtr menu, uint flags, int x, int y, IntPtr owner, IntPtr parameters);
        [DllImport("user32.dll")] private static extern bool DestroyMenu(IntPtr menu);
        [DllImport("user32.dll")] private static extern bool EndMenu();
        [DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT point);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    }
}
