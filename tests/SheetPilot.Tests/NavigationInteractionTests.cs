using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using SheetPilot.Core;
using SheetPilot.UI;
using Forms = System.Windows.Forms;

internal static partial class Program
{
    private static void NavigationInteractionTests()
    {
        Check(Application.Current == null && !PickerWindows().Any(), "navigation: cold native host without WPF windows or Electron");
        using(var model=new NavigationModel())
        using(var view=new NavigationView(model))
        using(var form=new Forms.Form{Text="GridPilot navigation regression",Width=420,Height=700,ShowInTaskbar=false})
        using(var host=new System.Windows.Forms.Integration.ElementHost{Dock=Forms.DockStyle.Fill,Child=view})
        {
            var sheets=new[]{new SheetItem("one","Sheet1",1,SheetVisibility.Visible,true,"工作表","book","Book"),new SheetItem("two","Sheet2",2,SheetVisibility.Visible,false,"工作表","book","Book")};
            var edits=new List<SheetEdit>();model.EditRequested+=edits.Add;
            model.Update(new WorkbookSnapshot("book","Book",sheets));
            form.Controls.Add(host);form.Show();PumpUntil(()=>view.IsLoaded&&model.Results.Count==2);view.UpdateLayout();
            var list=(ListBox)view.FindName("SheetList");
            var row=(ListBoxItem)list.ItemContainerGenerator.ContainerFromIndex(0);
            var label=PickerDescendants(row).OfType<TextBlock>().First(b=>b.Name=="SheetNameLabel");
            var box=PickerDescendants(row).OfType<TextBox>().Single(b=>b.Name=="InlineName");
            Action begin=()=>{InvokeNavigation(view,"Rename",sheets[0]);Check(model.IsEditing&&box.IsVisible,"navigation: rename opens inline editor");};
            Action<UIElement,MouseButton> click=(target,button)=>target.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice,Environment.TickCount,button){RoutedEvent=Mouse.PreviewMouseDownEvent});
            Action<Key> key=k=>box.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice,PresentationSource.FromVisual(view),Environment.TickCount,k){RoutedEvent=Keyboard.PreviewKeyDownEvent});
            try
            {
                ExerciseNavigationMenu(view,list,model,false,()=>
                {
                    Check(!model.IsEditing,"navigation: sheet menu does not start rename");
                    model.Update(new WorkbookSnapshot("book","Book",sheets.Reverse()));
                    Check(model.Results[0].Key=="one","navigation: refresh does not replace rows during menu tracking");
                    view.FocusSelected();
                });
                begin();Check(!box.CanUndo,"navigation: new rename session cannot undo into a previous draft");box.Text="空白提交";click(list,MouseButton.Left);
                Check(!model.IsEditing&&edits.Last().Name=="空白提交","navigation: list whitespace commits rename");
                begin();box.Text="页脚提交";click(view,MouseButton.Left);
                Check(!model.IsEditing&&edits.Last().Name=="页脚提交","navigation: pane background commits rename");
                begin();box.Text="同一行提交";click(label,MouseButton.Left);
                Check(!model.IsEditing&&edits.Last().Name=="同一行提交","navigation: non-editor area in same row commits");
                begin();box.Text="保留编辑";box.SelectAll();click(box,MouseButton.Right);
                Check(model.IsEditing&&box.SelectionLength==4,"navigation: editor right click preserves draft and selection");
                int count=edits.Count;
                ExerciseNavigationMenu(view,box,model,true,()=>Check(model.IsEditing&&box.Text=="保留编辑","navigation: edit menu keeps draft alive"));
                PumpNavigation();
                Check(model.IsEditing&&edits.Count==count,"navigation: dismissing edit menu does not submit");
                Clipboard.SetText("粘贴的新表名");box.SelectAll();box.Paste();
                Check(model.IsEditing&&box.Text=="粘贴的新表名","navigation: paste replaces selected sheet name");
                key(Key.Enter);Check(!model.IsEditing&&edits.Last().Name=="粘贴的新表名","navigation: Enter submits pasted draft");
                begin();box.Text="取消此名称";count=edits.Count;key(Key.Escape);
                Check(!model.IsEditing&&edits.Count==count,"navigation: Escape discards draft");
                begin();box.Text="invalid/name";click(list,MouseButton.Left);
                Check(model.IsEditing&&model.HasMessage,"navigation: invalid name stays editable on blank click");
                box.Text="有效名称";key(Key.Enter);
                begin();box.Text="失焦提交";view.FocusSearch();PumpNavigation();
                Check(!model.IsEditing&&edits.Last().Name=="失焦提交","navigation: focus elsewhere still submits");
                begin();box.Text="卸载取消";count=edits.Count;host.Child=null;PumpNavigation();
                Check(!model.IsEditing&&edits.Count==count,"navigation: unloading cancels draft");
            }
            finally{host.Child=null;form.Close();}
        }
        Console.WriteLine("PASS "+checks+" navigation interaction assertions");
    }

    private static void InvokeNavigation(NavigationView view,string method,params object[] args) =>
        typeof(NavigationView).GetMethod(method,BindingFlags.Instance|BindingFlags.NonPublic).Invoke(view,args);

    private static void PumpNavigation()
    {
        var frame=new DispatcherFrame();
        Dispatcher.CurrentDispatcher.BeginInvoke(new Action(()=>frame.Continue=false),DispatcherPriority.ApplicationIdle);
        Dispatcher.PushFrame(frame);
    }

    private static void ExerciseNavigationMenu(NavigationView view,FrameworkElement target,NavigationModel model,bool inline,Action inspect)
    {
        // Invoke the same ContextMenuOpening route in a cold native host. A native timer
        // runs inside TrackPopupMenu's nested loop, proving the menu survives that loop.
        bool observed=false;Exception failure=null;
        using(var timer=new Forms.Timer{Interval=350})
        {
            timer.Tick+=(s,e)=>
            {
                timer.Stop();
                try
                {
                    Check(model.IsMenuOpen,"navigation: native menu remains open after 350ms");
                    var info=new NavigationGuiInfo{Size=Marshal.SizeOf(typeof(NavigationGuiInfo))};
                    Check(GetGUIThreadInfo(0,ref info)&&(info.Flags&4)!=0,"navigation: Windows owns the popup menu loop");
                    inspect();observed=true;
                }
                catch(Exception ex){failure=ex;}
                finally{EndNavigationMenu();}
            };
            timer.Start();
            // ContextMenuEventArgs has no public constructor; use WPF's keyboard-menu constructor.
            var args=(ContextMenuEventArgs)Activator.CreateInstance(typeof(ContextMenuEventArgs),BindingFlags.Instance|BindingFlags.NonPublic,null,new object[]{target,true},null);
            InvokeNavigation(view,inline?"InlineMenu":"OpenSheetMenu",target,args);
            timer.Stop();
        }
        if(failure!=null)throw failure;
        Check(observed&&!model.IsMenuOpen,"navigation: menu closes cleanly without warming another window");
    }
    [StructLayout(LayoutKind.Sequential)] private struct NavigationGuiInfo
    {
        public int Size;public uint Flags;public IntPtr Active,Focus,Capture,MenuOwner,MoveSize,Caret;
        public int Left,Top,Right,Bottom;
    }
    [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread,ref NavigationGuiInfo info);
    [DllImport("user32.dll",EntryPoint="EndMenu")] private static extern bool EndNavigationMenu();
}
