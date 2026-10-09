using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public partial class NavigationView : UserControl
    {
        public NavigationView(NavigationModel model) { InitializeComponent(); DataContext = model; }
        private NavigationModel Model => (NavigationModel)DataContext;
        public void FocusSearch() { SearchBox.Focus(); SearchBox.SelectAll(); }
        private void RefreshClick(object sender, RoutedEventArgs e) => Model.Refresh();
        private void SettingsClick(object sender, RoutedEventArgs e) => Model.IsSettingsOpen = true;
        private void BackClick(object sender, RoutedEventArgs e) { Model.IsSettingsOpen = false; FocusSearch(); }
        private void ResetClick(object sender, RoutedEventArgs e) => Model.Settings.Reset();
        private void ClearClick(object sender, RoutedEventArgs e) { Model.Query = ""; FocusSearch(); }
        private void SheetClick(object sender, MouseButtonEventArgs e)
        {
            var row = ItemsControl.ContainerFromElement(SheetList, e.OriginalSource as DependencyObject) as ListBoxItem;
            if (row?.DataContext is SheetItem item) { Model.Selected = item; Model.Navigate(item); }
        }
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (Model.IsSettingsOpen)
            {
                if (e.Key == Key.Escape) { Model.IsSettingsOpen = false; FocusSearch(); e.Handled = true; }
                return;
            }
            if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) != 0) { FocusSearch(); e.Handled = true; }
            else if (e.Key == Key.F5) { Model.Refresh(); e.Handled = true; }
            else if (e.Key == Key.Escape) { Model.Query = ""; e.Handled = true; }
            else if (e.Key == Key.Enter && (SearchBox.IsKeyboardFocusWithin || SheetList.IsKeyboardFocusWithin)) { Model.Navigate(Model.Selected); e.Handled = true; }
            else if (e.Key == Key.Down && SearchBox.IsKeyboardFocusWithin && SheetList.Items.Count > 0)
            { SheetList.Focus(); SheetList.SelectedIndex = 0; ((ListBoxItem)SheetList.ItemContainerGenerator.ContainerFromIndex(0))?.Focus(); e.Handled = true; }
        }
    }
}
