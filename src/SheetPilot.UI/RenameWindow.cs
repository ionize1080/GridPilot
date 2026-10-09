using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public sealed class RenameWindow : Window
    {
        private readonly ToolPreferences preferences;
        private readonly Func<bool, WorkbookSnapshot> capture;
        private readonly Func<IReadOnlyList<RenameEntry>, string> apply;
        private readonly Func<string> save;
        private readonly DataGrid grid;
        private readonly TextBlock status;
        private readonly Button applyButton;
        private readonly DispatcherTimer timer;
        private readonly HashSet<string> excluded = new HashSet<string>();
        private IReadOnlyList<SheetItem> sources = new SheetItem[0];
        private IReadOnlyList<RenameEntry> plan = new RenameEntry[0];
        private bool rebuilding, capturedScope;
        public RenameWindow(ToolPreferences preferences, NavigationSettings colors, Func<bool, WorkbookSnapshot> capture, Func<IReadOnlyList<RenameEntry>, string> apply, Func<string> save)
        {
            this.preferences = preferences; this.capture = capture; this.apply = apply; this.save = save;
            ToolUi.Theme(this, colors, "批量重命名", 1100, 800);
            var root = new DockPanel { Margin = new Thickness(18) }; Content = root;
            var bottom = new StackPanel { Margin = new Thickness(0, 12, 0, 0) }; DockPanel.SetDock(bottom, Dock.Bottom); root.Children.Add(bottom);
            status = ToolUi.Note("准备预览…"); bottom.Children.Add(status);
            var actions = new StackPanel { Orientation = Orientation.Horizontal }; bottom.Children.Add(actions);
            applyButton = ToolUi.Button(actions, "应用预览", ApplyPreview);
            ToolUi.Button(actions, "刷新页面", RefreshSources); ToolUi.Button(actions, "保存为默认规则", () => { Commit(); status.Text = save(); }); ToolUi.Button(actions, "关闭", Close);
            var columns = new Grid(); columns.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(290) }); columns.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(16) }); columns.ColumnDefinitions.Add(new ColumnDefinition()); root.Children.Add(columns);
            var inputs = ToolUi.RenameControls(preferences.Rename, true); inputs.Margin = new Thickness(0, 0, 12, 8); columns.Children.Add(ToolUi.Scroll(inputs));
            var preview = new DockPanel(); Grid.SetColumn(preview, 2); columns.Children.Add(preview);
            var title = ToolUi.Heading("应用前预览"); DockPanel.SetDock(title, Dock.Top); preview.Children.Add(title);
            var note = ToolUi.Note("勾选要处理的页面；重名和非法名称会阻止应用。结构受保护或只读工作簿在应用前统一检查。"); DockPanel.SetDock(note, Dock.Top); preview.Children.Add(note);
            var selection = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 3, 0, 10) }; DockPanel.SetDock(selection, Dock.Top); preview.Children.Add(selection);
            ToolUi.Button(selection, "全选", () => { excluded.Clear(); BuildPreview(); }); ToolUi.Button(selection, "全不选", () => { foreach (var source in sources) excluded.Add(RenamePlanner.Key(source)); BuildPreview(); });
            grid = new DataGrid { AutoGenerateColumns = false, CanUserAddRows = false, CanUserDeleteRows = false, CanUserSortColumns = false, MinColumnWidth = 110, SelectionMode = DataGridSelectionMode.Single, HeadersVisibility = DataGridHeadersVisibility.Column, GridLinesVisibility = DataGridGridLinesVisibility.Horizontal, RowHeight = 36, BorderThickness = new Thickness(0), EnableRowVirtualization = true };
            grid.SetResourceReference(Control.BackgroundProperty, "SurfaceBrush"); grid.SetResourceReference(Control.ForegroundProperty, "TextBrush"); grid.SetResourceReference(DataGrid.HorizontalGridLinesBrushProperty, "BorderBrush");
            grid.SetResourceReference(DataGrid.RowBackgroundProperty, "SurfaceBrush");
            var headerStyle = new Style(typeof(DataGridColumnHeader)); headerStyle.Setters.Add(new Setter(Control.BackgroundProperty, FindResource("CanvasBrush"))); headerStyle.Setters.Add(new Setter(Control.ForegroundProperty, FindResource("MutedBrush"))); headerStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(7))); grid.ColumnHeaderStyle = headerStyle;
            var check = new FrameworkElementFactory(typeof(CheckBox)); check.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center); check.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center); check.SetBinding(ToggleButton.IsCheckedProperty, new Binding("Selected") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
            grid.Columns.Add(new DataGridTemplateColumn { Header = "选", Width = 34, MinWidth = 34, CellTemplate = new DataTemplate { VisualTree = check } });
            grid.Columns.Add(new DataGridTextColumn { Header = "工作簿", Binding = new Binding("Book"), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            grid.Columns.Add(new DataGridTextColumn { Header = "原名称", Binding = new Binding("Before"), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            grid.Columns.Add(new DataGridTextColumn { Header = "新名称", Binding = new Binding("After"), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            grid.Columns.Add(new DataGridTextColumn { Header = "状态", Binding = new Binding("Status"), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) }); preview.Children.Add(grid);
            grid.SizeChanged += (s,e) => { foreach (var column in grid.Columns.Skip(1)) column.Width = new DataGridLength(Math.Max(110, (e.NewSize.Width - 48) / 4)); };
            timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(180) }; timer.Tick += (s,e) => { timer.Stop(); if (capturedScope != preferences.Rename.AllWorkbooks) RefreshSources(); else BuildPreview(); };
            inputs.AddHandler(TextBox.TextChangedEvent, new TextChangedEventHandler((s,e) => Schedule())); inputs.AddHandler(ToggleButton.CheckedEvent, new RoutedEventHandler((s,e) => Schedule())); inputs.AddHandler(ToggleButton.UncheckedEvent, new RoutedEventHandler((s,e) => Schedule()));
            inputs.AddHandler(Selector.SelectionChangedEvent, new SelectionChangedEventHandler((s,e) => Schedule())); inputs.AddHandler(RangeBase.ValueChangedEvent, new RoutedPropertyChangedEventHandler<double>((s,e) => Schedule())); inputs.LostKeyboardFocus += (s,e) => Schedule();
            Loaded += (s,e) => RefreshSources(); Closed += (s,e) => timer.Stop();
        }
        private void Schedule() { if (rebuilding || !IsLoaded) return; applyButton.IsEnabled = false; timer.Stop(); timer.Start(); }
        private void Commit() { grid.CommitEdit(); Keyboard.ClearFocus(); preferences.Normalize(); }
        public void RefreshSources()
        {
            try { Commit(); capturedScope = preferences.Rename.AllWorkbooks; sources = capture(capturedScope).Sheets; BuildPreview(); }
            catch (Exception ex) { plan = new RenameEntry[0]; applyButton.IsEnabled = false; status.Text = ex.Message; }
        }
        public void BuildPreview()
        {
            if (rebuilding) return; rebuilding = true;
            try
            {
                plan = RenamePlanner.Build(sources, preferences.Rename, excluded);
                grid.ItemsSource = plan.Select(entry => new Row(entry, !excluded.Contains(RenamePlanner.Key(entry.Source)), value => { if (value) excluded.Remove(RenamePlanner.Key(entry.Source)); else excluded.Add(RenamePlanner.Key(entry.Source)); Schedule(); })).ToArray();
                int count = plan.Count(p => p.Changed), errors = plan.Count(p => p.Changed && p.Error != null);
                status.Text = "共 " + plan.Count + " 页 · 待修改 " + count + " · 错误 " + errors;
                applyButton.IsEnabled = count > 0 && errors == 0;
            }
            catch (Exception ex) when (ex is ArgumentException || ex is System.Text.RegularExpressions.RegexMatchTimeoutException)
            { plan = new RenameEntry[0]; grid.ItemsSource = null; status.Text = "无法预览：" + ex.Message; applyButton.IsEnabled = false; }
            finally { rebuilding = false; }
        }
        private void ApplyPreview()
        {
            Commit(); timer.Stop(); BuildPreview(); if (!applyButton.IsEnabled) return;
            try { string result = apply(plan); sources = capture(capturedScope).Sheets; BuildPreview(); status.Text = result; }
            catch (Exception ex) { status.Text = ex.Message; applyButton.IsEnabled = false; }
        }
        private sealed class Row
        {
            private bool selected; private readonly Action<bool> changed;
            public bool Selected { get => selected; set { if (selected == value) return; selected = value; changed(value); } }
            public string Book { get; } public string Before { get; } public string After { get; } public string Status { get; }
            public Row(RenameEntry entry, bool selected, Action<bool> changed) { this.selected = selected; this.changed = changed; Book = entry.Source.WorkbookName; Before = entry.Source.Name; After = entry.NewName; Status = entry.Error ?? (!entry.Selected ? "跳过" : entry.Changed ? "待应用" : "不变"); }
        }
    }
}
