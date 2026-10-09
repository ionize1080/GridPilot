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
            var root = new DockPanel { Margin = new Thickness(24) }; Content = root;
            var heading = new StackPanel { Margin = new Thickness(0,0,0,20) }; DockPanel.SetDock(heading,Dock.Top); root.Children.Add(heading);
            heading.Children.Add(new TextBlock { Text="批量重命名", FontSize=22, FontWeight=FontWeights.SemiBold });
            heading.Children.Add(ToolUi.Note("设置规则，核对新名称，再应用到所选页面。"));
            var bottom = new DockPanel { Margin = new Thickness(0,18,0,0), LastChildFill=true }; DockPanel.SetDock(bottom, Dock.Bottom); root.Children.Add(bottom);
            var actions = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment=HorizontalAlignment.Right }; DockPanel.SetDock(actions,Dock.Right); bottom.Children.Add(actions);
            ToolUi.Button(actions, "关闭", Close);
            applyButton = ToolUi.Button(actions, "应用更改", ApplyPreview); applyButton.SetResourceReference(Control.BackgroundProperty,"AccentBrush"); applyButton.SetResourceReference(Control.ForegroundProperty,"AccentTextBrush"); applyButton.Padding=new Thickness(20,9,20,9);
            status = ToolUi.Note("准备预览…"); status.VerticalAlignment=VerticalAlignment.Center; bottom.Children.Add(status);
            var columns = new Grid(); columns.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(286) }); columns.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) }); columns.ColumnDefinitions.Add(new ColumnDefinition()); root.Children.Add(columns);
            var left = new DockPanel(); var leftActions=new StackPanel { Margin=new Thickness(0,12,0,0) }; DockPanel.SetDock(leftActions,Dock.Bottom); left.Children.Add(leftActions);
            ToolUi.Button(leftActions,"保存为默认规则",()=>{Commit();status.Text=save();});
            var inputs = ToolUi.RenameControls(preferences.Rename, true); inputs.Margin = new Thickness(0, 0, 10, 8); if(inputs.Children[0] is TextBlock firstHeading) { firstHeading.Margin=new Thickness(0,0,0,8); firstHeading.FontSize=15; } left.Children.Add(ToolUi.Scroll(inputs)); columns.Children.Add(Card(left));
            var preview = new DockPanel(); var previewCard=Card(preview); Grid.SetColumn(previewCard, 2); columns.Children.Add(previewCard);
            var title = new TextBlock { Text="应用前预览",FontSize=15,FontWeight=FontWeights.SemiBold,Margin=new Thickness(0,0,0,6) }; DockPanel.SetDock(title, Dock.Top); preview.Children.Add(title);
            var note = ToolUi.Note("重名或非法名称会阻止应用；仅修改勾选且符合规则的页面。"); DockPanel.SetDock(note, Dock.Top); preview.Children.Add(note);
            var selection = new DockPanel { Margin = new Thickness(0, 8, 0, 14) }; DockPanel.SetDock(selection, Dock.Top); preview.Children.Add(selection);
            var refresh=new StackPanel {Orientation=Orientation.Horizontal}; DockPanel.SetDock(refresh,Dock.Right); selection.Children.Add(refresh); ToolUi.Button(refresh,"刷新页面",RefreshSources);
            var selectActions=new StackPanel {Orientation=Orientation.Horizontal};selection.Children.Add(selectActions);
            ToolUi.Button(selectActions, "全选", () => { excluded.Clear(); BuildPreview(); }); ToolUi.Button(selectActions, "全不选", () => { foreach (var source in sources) excluded.Add(RenamePlanner.Key(source)); BuildPreview(); });
            grid = new DataGrid { AutoGenerateColumns = false, CanUserAddRows = false, CanUserDeleteRows = false, CanUserSortColumns = false, CanUserReorderColumns=false, MinColumnWidth = 90, SelectionMode = DataGridSelectionMode.Single, HeadersVisibility = DataGridHeadersVisibility.Column, GridLinesVisibility = DataGridGridLinesVisibility.Horizontal, RowHeight = 44, ColumnHeaderHeight=40, BorderThickness = new Thickness(0), EnableRowVirtualization = true, CanUserResizeRows=false };
            grid.SetResourceReference(Control.BackgroundProperty, "SurfaceBrush"); grid.SetResourceReference(Control.ForegroundProperty, "TextBrush"); grid.SetResourceReference(DataGrid.HorizontalGridLinesBrushProperty, "BorderBrush");
            grid.SetResourceReference(DataGrid.RowBackgroundProperty, "SurfaceBrush"); grid.SetResourceReference(DataGrid.AlternatingRowBackgroundProperty,"CanvasBrush");
            var cellStyle=new Style(typeof(DataGridCell)); cellStyle.Setters.Add(new Setter(Control.BorderThicknessProperty,new Thickness(0))); cellStyle.Setters.Add(new Setter(Control.PaddingProperty,new Thickness(0))); cellStyle.Setters.Add(new Setter(Control.VerticalContentAlignmentProperty,VerticalAlignment.Center));
            var selectedTrigger=new Trigger {Property=DataGridCell.IsSelectedProperty,Value=true}; selectedTrigger.Setters.Add(new Setter(Control.BackgroundProperty,FindResource("SelectedBrush"))); selectedTrigger.Setters.Add(new Setter(Control.ForegroundProperty,FindResource("TextBrush"))); cellStyle.Triggers.Add(selectedTrigger); grid.CellStyle=cellStyle;
            var headerStyle = new Style(typeof(DataGridColumnHeader)); headerStyle.Setters.Add(new Setter(Control.BackgroundProperty, FindResource("CanvasBrush"))); headerStyle.Setters.Add(new Setter(Control.ForegroundProperty, FindResource("MutedBrush"))); headerStyle.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(12,0,12,0))); headerStyle.Setters.Add(new Setter(Control.VerticalContentAlignmentProperty,VerticalAlignment.Center)); headerStyle.Setters.Add(new Setter(Control.BorderThicknessProperty,new Thickness(0))); grid.ColumnHeaderStyle = headerStyle;
            var check = new FrameworkElementFactory(typeof(CheckBox)); check.SetValue(FrameworkElement.MarginProperty,new Thickness(7,0,0,0)); check.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Center); check.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center); check.SetBinding(ToggleButton.IsCheckedProperty, new Binding("Selected") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
            grid.Columns.Add(new DataGridTemplateColumn { Header = "", Width = 42, MinWidth = 42, CellTemplate = new DataTemplate { VisualTree = check } });
            grid.Columns.Add(new DataGridTextColumn { Header = "工作簿", Binding = new Binding("Book"), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            grid.Columns.Add(new DataGridTextColumn { Header = "原名称", Binding = new Binding("Before"), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            grid.Columns.Add(new DataGridTextColumn { Header = "新名称", Binding = new Binding("After"), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            grid.Columns.Add(new DataGridTextColumn { Header = "状态", Binding = new Binding("Status"), IsReadOnly = true, Width = new DataGridLength(1, DataGridLengthUnitType.Star) }); preview.Children.Add(grid);
            foreach(var column in grid.Columns.OfType<DataGridTextColumn>())
            {
                var textStyle=new Style(typeof(TextBlock)); textStyle.Setters.Add(new Setter(FrameworkElement.VerticalAlignmentProperty,VerticalAlignment.Center)); textStyle.Setters.Add(new Setter(FrameworkElement.MarginProperty,new Thickness(12,0,12,0))); textStyle.Setters.Add(new Setter(TextBlock.TextTrimmingProperty,TextTrimming.CharacterEllipsis)); textStyle.Setters.Add(new Setter(FrameworkElement.ToolTipProperty,new Binding(((Binding)column.Binding).Path.Path)));
                if(column.Header.ToString()=="新名称") {textStyle.Setters.Add(new Setter(TextBlock.FontWeightProperty,FontWeights.SemiBold));textStyle.Setters.Add(new Setter(TextBlock.ForegroundProperty,FindResource("AccentBrush")));}
                column.ElementStyle=textStyle;
            }
            grid.SizeChanged += (s,e) => { double available=Math.Max(360,e.NewSize.Width-72); for(int i=1;i<grid.Columns.Count;i++) grid.Columns[i].Width=new DataGridLength(Math.Max(90,available*(i==4?.18:i==1?.26:.28))); };
            timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(180) }; timer.Tick += (s,e) => { timer.Stop(); if (capturedScope != preferences.Rename.AllWorkbooks) RefreshSources(); else BuildPreview(); };
            inputs.AddHandler(TextBox.TextChangedEvent, new TextChangedEventHandler((s,e) => Schedule())); inputs.AddHandler(ToggleButton.CheckedEvent, new RoutedEventHandler((s,e) => Schedule())); inputs.AddHandler(ToggleButton.UncheckedEvent, new RoutedEventHandler((s,e) => Schedule()));
            inputs.AddHandler(Selector.SelectionChangedEvent, new SelectionChangedEventHandler((s,e) => Schedule())); inputs.AddHandler(RangeBase.ValueChangedEvent, new RoutedPropertyChangedEventHandler<double>((s,e) => Schedule())); inputs.LostKeyboardFocus += (s,e) => Schedule();
            Loaded += (s,e) => RefreshSources(); Closed += (s,e) => timer.Stop();
        }
        private Border Card(UIElement child)
        {
            var card=new Border {Child=child,Padding=new Thickness(16),CornerRadius=new CornerRadius(10),BorderThickness=new Thickness(1)};
            card.SetResourceReference(Border.BackgroundProperty,"SurfaceBrush");card.SetResourceReference(Border.BorderBrushProperty,"BorderBrush");return card;
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
