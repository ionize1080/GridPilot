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
            model.Settings.PropertyChanged += SettingsChanged;UpdateQuickSort();
        }
        private void SettingsChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName.EndsWith("Color", System.StringComparison.Ordinal)) ViewTheme.Apply(Resources, Model.Settings);
            if(e.PropertyName.StartsWith("QuickSort",System.StringComparison.Ordinal))UpdateQuickSort();
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
        public void FocusSelected(){if(editing!=null||Model.Selected==null)return;var row=SheetList.ItemContainerGenerator.ContainerFromItem(Model.Selected) as ListBoxItem;row?.Focus();}
        private void RefreshClick(object sender, RoutedEventArgs e) => Model.Refresh();
        private void SettingsClick(object sender, RoutedEventArgs e) => Model.IsSettingsOpen = true;
        private void BackClick(object sender, RoutedEventArgs e) { Model.IsSettingsOpen = false; FocusSearch(); }
        private void ResetClick(object sender, RoutedEventArgs e) => Model.Settings.Reset();
        private void ClearClick(object sender, RoutedEventArgs e) { Model.Query = ""; FocusSearch(); }
        private void SheetClick(object sender, MouseButtonEventArgs e)
        {
            if(dragging||editing!=null)return;
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
            System.Action<string,string> add=(title,action)=>{var entry=new MenuItem{Header=title};entry.Click+=(s,a)=>{if(action=="rename")Dispatcher.BeginInvoke(new System.Action(()=>Rename(item)),System.Windows.Threading.DispatcherPriority.Input);else if(action=="copy"){try{Clipboard.SetText(item.Name);}catch(System.Exception){Model.Report("剪贴板忙，请重试。");}}else Model.Edit(new SheetEdit{Source=item,Action=action});};menu.Items.Add(entry);};
            add("原地重命名    F2","rename");add("工作表自定义排序…","sort");add("在后面新增工作表","add");menu.Items.Add(new Separator());
            add("上移    Alt+↑","up");add("下移    Alt+↓","down");add("移到最前","first");add("移到最后","last");menu.Items.Add(new Separator());
            add(item.CanNavigate?"隐藏工作表":"取消隐藏",item.CanNavigate?"hide":"show");add("复制工作表名称","copy");
        }
        private TextBox editing;
        private TextBlock editingLabel;
        private SheetItem editingItem;
        private static T Child<T>(DependencyObject root,string name) where T:FrameworkElement
        {
            if(root is T element&&element.Name==name)return element;
            for(int i=0;i<System.Windows.Media.VisualTreeHelper.GetChildrenCount(root);i++){var found=Child<T>(System.Windows.Media.VisualTreeHelper.GetChild(root,i),name);if(found!=null)return found;}return null;
        }
        private void Rename(SheetItem item)
        {
            EndRename(false);SheetList.ScrollIntoView(item);SheetList.UpdateLayout();
            var row=SheetList.ItemContainerGenerator.ContainerFromItem(item) as ListBoxItem;if(row==null)return;
            editing=Child<TextBox>(row,"InlineName");editingLabel=Child<TextBlock>(row,"SheetNameLabel");if(editing==null||editingLabel==null)return;
            editingItem=item;Model.IsEditing=true;editing.Text=item.Name;editingLabel.Visibility=Visibility.Collapsed;editing.Visibility=Visibility.Visible;editing.Focus();editing.SelectAll();dragItem=null;
        }
        private void EndRename(bool commit)
        {
            if(editing==null)return;
            string text=editing.Text;var item=editingItem;
            if(commit){string error=RenamePlanner.ValidateName(text);if(error!=null){Model.Report(error);return;}}
            var box=editing;editing=null;editingItem=null;Model.IsEditing=false;editingLabel.Visibility=Visibility.Visible;box.Visibility=Visibility.Collapsed;
            if(commit&&text!=item.Name)Model.Edit(new SheetEdit{Source=item,Action="rename",Name=text});
        }
        private void InlineKey(object sender,KeyEventArgs e){if(e.Key==Key.Enter||e.Key==Key.Escape){EndRename(e.Key==Key.Enter);e.Handled=true;}}
        private void InlineBlur(object sender,KeyboardFocusChangedEventArgs e){EndRename(true);}
        private void UpdateQuickSort()
        {
            if(QuickSortButton==null)return;
            QuickSortButton.Content=Model.Settings.QuickSortDescending?"↓":"↑";
            QuickSortButton.ToolTip=(Model.Settings.QuickSortDescending?"一键降序":"一键升序")+" · "+new[]{"名称 A–Z","自然数字顺序","最近使用的自定义规则"}[Model.Settings.QuickSortMode]+"（同一工作簿的全部标签）";
        }
        private void QuickSortClick(object sender,RoutedEventArgs e)
        {EndRename(true);var item=Model.Selected;if(item==null)return;Model.Edit(new SheetEdit{Source=item,Action=Model.Settings.QuickSortDescending?"sortDesc":"sortAsc"});Model.Settings.QuickSortDescending=!Model.Settings.QuickSortDescending;}
        private void SortClick(object sender,RoutedEventArgs e)
        {EndRename(true);var item=Model.Selected;if(item==null){Model.Report("请先选择一个工作表。");return;}Model.Edit(new SheetEdit{Source=item,Action="sort"});}
        private void DragStart(object sender,MouseButtonEventArgs e)
        {
            dragOrigin=e.GetPosition(SheetList);dragItem=RowAt(e.OriginalSource)?.DataContext as SheetItem;
            if(editing!=null&&dragItem!=null&&dragItem.Key!=editingItem.Key)
            {EndRename(true);if(editing!=null){e.Handled=true;return;}}
            if(e.ClickCount==2&&editing==null&&dragItem!=null){var item=dragItem;Model.Selected=item;Rename(item);e.Handled=true;}
        }
        private void DragMove(object sender,MouseEventArgs e)
        {
            if(editing!=null||dragItem==null||e.LeftButton!=MouseButtonState.Pressed||dragging)return;
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
            if(editing!=null)return;
            if (Model.IsSettingsOpen)
            {
                if (e.Key == Key.Escape && !(Keyboard.FocusedElement is TextBox)) { Model.IsSettingsOpen = false; FocusSearch(); e.Handled = true; }
                return;
            }
            if(e.Key==Key.F2&&SheetList.IsKeyboardFocusWithin&&Model.Selected!=null){Rename(Model.Selected);e.Handled=true;return;}
            if(e.Key==Key.Space&&SheetList.IsKeyboardFocusWithin){Model.Navigate(Model.Selected);e.Handled=true;return;}
            if(e.Key==Key.Insert&&Keyboard.Modifiers==ModifierKeys.None&&SheetList.IsKeyboardFocusWithin&&Model.Selected!=null){Model.Edit(new SheetEdit{Source=Model.Selected,Action="add"});e.Handled=true;return;}
            var key=e.Key==Key.System?e.SystemKey:e.Key;
            if((Keyboard.Modifiers&ModifierKeys.Alt)!=0&&(key==Key.Up||key==Key.Down)&&SheetList.IsKeyboardFocusWithin&&Model.Selected!=null){Model.Edit(new SheetEdit{Source=Model.Selected,Action=key==Key.Up?"up":"down"});e.Handled=true;return;}
            if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) != 0) { FocusSearch(); e.Handled = true; }
            else if (e.Key == Key.F5) { Model.Refresh(); e.Handled = true; }
            else if (e.Key == Key.Escape) { Model.Query = ""; e.Handled = true; }
            else if (e.Key == Key.Enter && (SearchBox.IsKeyboardFocusWithin || SheetList.IsKeyboardFocusWithin)) { Model.Navigate(Model.Selected); e.Handled = true; }
            else if (e.Key == Key.Down && SearchBox.IsKeyboardFocusWithin && SheetList.Items.Count > 0)
            { SheetList.Focus(); SheetList.SelectedIndex = 0; ((ListBoxItem)SheetList.ItemContainerGenerator.ContainerFromIndex(0))?.Focus(); e.Handled = true; }
        }
    }
}
