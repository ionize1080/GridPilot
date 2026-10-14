using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public partial class NavigationView : UserControl, System.IDisposable
    {
        public NavigationView(NavigationModel model)
        {
            InitializeComponent(); DataContext = model; ViewTheme.Apply(Resources, model.Settings);
            model.Settings.PropertyChanged += SettingsChanged;
        }
        private void SettingsChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName.EndsWith("Color", System.StringComparison.Ordinal)) ViewTheme.Apply(Resources, Model.Settings);
        }
        public void Dispose()
        {
            Model.Settings.PropertyChanged -= SettingsChanged;
            if (IsKeyboardFocusWithin) Keyboard.ClearFocus();
        }
        private void PaletteClick(object sender, RoutedEventArgs e) => Model.Settings.ApplyPalette((string)((Button)sender).Tag);
        private void OnToolTipOpening(object sender, ToolTipEventArgs e) { if (!Model.Settings.ShowTooltips) e.Handled = true; }
        private NavigationModel Model => (NavigationModel)DataContext;
        public void FocusSearch() { SearchBox.Focus(); SearchBox.SelectAll(); }
        private void RefreshClick(object sender, RoutedEventArgs e) => Model.Refresh();
        private void SettingsClick(object sender, RoutedEventArgs e) => Model.IsSettingsOpen = true;
        private void BackClick(object sender, RoutedEventArgs e) { Model.IsSettingsOpen = false; FocusSearch(); }
        private void ResetClick(object sender, RoutedEventArgs e) => Model.Settings.Reset();
        private void ClearClick(object sender, RoutedEventArgs e) { Model.Query = ""; FocusSearch(); }
        private void SheetClick(object sender, MouseButtonEventArgs e)
        {
            if(dragging)return;
            var row = ItemsControl.ContainerFromElement(SheetList, e.OriginalSource as DependencyObject) as ListBoxItem;
            if (row?.DataContext is SheetItem item) { Model.Selected = item; Model.Navigate(item); }
        }
        private Point dragOrigin;
        private SheetItem dragItem;
        private ListBoxItem dropCue;
        private bool dragging;
        private ListBoxItem RowAt(object source)=>ItemsControl.ContainerFromElement(SheetList,source as DependencyObject) as ListBoxItem;
        private void SelectContextRow(object sender,MouseButtonEventArgs e){var row=RowAt(e.OriginalSource);if(row?.DataContext is SheetItem item){Model.Selected=item;row.Focus();}}
        private void OpenSheetMenu(object sender,ContextMenuEventArgs e)
        {
            var item=Model.Selected;if(item==null){e.Handled=true;return;}
            var menu=SheetList.ContextMenu;menu.Items.Clear();
            System.Action<string,string> add=(title,action)=>{var entry=new MenuItem{Header=title};entry.Click+=(s,a)=>{if(action=="rename")Rename(item);else if(action=="copy"){try{Clipboard.SetText(item.Name);}catch(System.Exception){Model.Report("剪贴板忙，请重试。");}}else Model.Edit(new SheetEdit{Source=item,Action=action});};menu.Items.Add(entry);};
            add("重命名…    F2","rename");add("在后面新增工作表","add");menu.Items.Add(new Separator());
            add("上移    Alt+↑","up");add("下移    Alt+↓","down");add("移到最前","first");add("移到最后","last");menu.Items.Add(new Separator());
            add(item.CanNavigate?"隐藏工作表":"取消隐藏",item.CanNavigate?"hide":"show");add("复制工作表名称","copy");
        }
        private void Rename(SheetItem item)
        {
            var dialog=new Window{Title="重命名工作表",Width=380,Height=190,ResizeMode=ResizeMode.NoResize,WindowStartupLocation=WindowStartupLocation.CenterOwner,Owner=Window.GetWindow(this)};
            var panel=new StackPanel{Margin=new Thickness(20)};var text=new TextBox{Text=item.Name,MaxLength=31,Margin=new Thickness(0,10,0,12)};
            panel.Children.Add(new TextBlock{Text="工作表名称（最多 31 个字符）"});panel.Children.Add(text);
            var ok=new Button{Content="确定",IsDefault=true,HorizontalAlignment=HorizontalAlignment.Right,Padding=new Thickness(20,6,20,6)};
            ok.Click+=(s,e)=>{string error=RenamePlanner.ValidateName(text.Text);if(error!=null){MessageBox.Show(dialog,error);return;}dialog.DialogResult=true;};panel.Children.Add(ok);dialog.Content=panel;
            dialog.Loaded+=(s,e)=>{text.Focus();text.SelectAll();};dialog.PreviewKeyDown+=(s,e)=>{if(e.Key==Key.Escape)dialog.Close();};
            if(dialog.ShowDialog()==true)Model.Edit(new SheetEdit{Source=item,Action="rename",Name=text.Text});
        }
        private void DragStart(object sender,MouseButtonEventArgs e){dragOrigin=e.GetPosition(SheetList);dragItem=RowAt(e.OriginalSource)?.DataContext as SheetItem;}
        private void DragMove(object sender,MouseEventArgs e)
        {
            if(dragItem==null||e.LeftButton!=MouseButtonState.Pressed||dragging)return;
            var point=e.GetPosition(SheetList);if(System.Math.Abs(point.X-dragOrigin.X)<SystemParameters.MinimumHorizontalDragDistance&&System.Math.Abs(point.Y-dragOrigin.Y)<SystemParameters.MinimumVerticalDragDistance)return;
            dragging=true;try{DragDrop.DoDragDrop(SheetList,new DataObject("GridPilot.Sheet",dragItem),DragDropEffects.Move);}finally{dragging=false;dragItem=null;ClearDropCue();}
        }
        private void ClearDropCue(){if(dropCue!=null)dropCue.Tag=null;dropCue=null;}
        private void SheetDragLeave(object sender,DragEventArgs e)=>ClearDropCue();
        private void SheetDragOver(object sender,DragEventArgs e)
        {
            ClearDropCue();var source=e.Data.GetData("GridPilot.Sheet") as SheetItem;var row=RowAt(e.OriginalSource);var target=row?.DataContext as SheetItem;
            bool valid=source!=null&&target!=null&&source.WorkbookKey==target.WorkbookKey&&source.Key!=target.Key;
            e.Effects=valid?DragDropEffects.Move:DragDropEffects.None;e.Handled=true;
            if(valid){dropCue=row;row.Tag=e.GetPosition(row).Y>row.ActualHeight/2?"DropAfter":"DropBefore";}
        }
        private void SheetDrop(object sender,DragEventArgs e)
        {
            var source=e.Data.GetData("GridPilot.Sheet") as SheetItem;var row=RowAt(e.OriginalSource);var target=row?.DataContext as SheetItem;
            if(source!=null&&target!=null&&source.WorkbookKey==target.WorkbookKey&&source.Key!=target.Key)Model.Edit(new SheetEdit{Source=source,Target=target,Action="move",After=e.GetPosition(row).Y>row.ActualHeight/2});
            ClearDropCue();e.Handled=true;
        }
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (Model.IsSettingsOpen)
            {
                if (e.Key == Key.Escape && !(Keyboard.FocusedElement is TextBox)) { Model.IsSettingsOpen = false; FocusSearch(); e.Handled = true; }
                return;
            }
            if(e.Key==Key.F2&&SheetList.IsKeyboardFocusWithin&&Model.Selected!=null){Rename(Model.Selected);e.Handled=true;return;}
            if((Keyboard.Modifiers&ModifierKeys.Alt)!=0&&(e.Key==Key.Up||e.Key==Key.Down)&&Model.Selected!=null){Model.Edit(new SheetEdit{Source=Model.Selected,Action=e.Key==Key.Up?"up":"down"});e.Handled=true;return;}
            if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) != 0) { FocusSearch(); e.Handled = true; }
            else if (e.Key == Key.F5) { Model.Refresh(); e.Handled = true; }
            else if (e.Key == Key.Escape) { Model.Query = ""; e.Handled = true; }
            else if (e.Key == Key.Enter && (SearchBox.IsKeyboardFocusWithin || SheetList.IsKeyboardFocusWithin)) { Model.Navigate(Model.Selected); e.Handled = true; }
            else if (e.Key == Key.Down && SearchBox.IsKeyboardFocusWithin && SheetList.Items.Count > 0)
            { SheetList.Focus(); SheetList.SelectedIndex = 0; ((ListBoxItem)SheetList.ItemContainerGenerator.ContainerFromIndex(0))?.Focus(); e.Handled = true; }
        }
    }
}
