using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    internal static class ToolUi
    {
        internal static void Theme(Window window, NavigationSettings colors, string title, double width, double height)
        {
            window.Title = "GridPilot · " + title; window.FontFamily = new FontFamily("Microsoft YaHei UI, Segoe UI"); window.FontSize = 12;
            window.Width = Math.Min(width, SystemParameters.WorkArea.Width * .94); window.Height = Math.Min(height, SystemParameters.WorkArea.Height * .9);
            window.MinWidth = Math.Min(600, window.Width); window.MinHeight = Math.Min(440, window.Height);
            window.WindowStartupLocation = WindowStartupLocation.CenterOwner; window.ShowInTaskbar = false;
            window.Resources.MergedDictionaries.Add(new ResourceDictionary { Source = new Uri("/SheetPilot.UI;component/Styles.xaml", UriKind.Relative) });
            ViewTheme.Apply(window.Resources, colors);
            window.SetResourceReference(Control.BackgroundProperty, "CanvasBrush"); window.SetResourceReference(Control.ForegroundProperty, "TextBrush");
        }
        internal static TextBlock Heading(string text) { var block = new TextBlock { Text = text, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 16, 0, 8) }; block.SetResourceReference(TextBlock.ForegroundProperty, "AccentBrush"); return block; }
        internal static TextBlock Note(string text) { var block = new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap, FontSize = 11, Margin = new Thickness(0, 6, 0, 6) }; block.SetResourceReference(TextBlock.ForegroundProperty, "MutedBrush"); return block; }
        internal static void Bind(FrameworkElement element, DependencyProperty property, object source, string path) => element.SetBinding(property, new Binding(path) { Source = source, Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged });
        internal static CheckBox Check(Panel parent, object source, string path, string caption)
        { var control = new CheckBox { Content = caption, Margin = new Thickness(0, 5, 0, 5) }; Bind(control, ToggleButton.IsCheckedProperty, source, path); parent.Children.Add(control); return control; }
        internal static void Text(Panel parent, object source, string path, string caption)
        { parent.Children.Add(Note(caption)); var control = new TextBox { Padding = new Thickness(8, 6, 8, 6), MaxLength = 512 }; Bind(control, TextBox.TextProperty, source, path); parent.Children.Add(control); }
        internal static void Choice(Panel parent, object source, string path, string caption, params string[] items)
        { parent.Children.Add(Note(caption)); var control = new ComboBox { ItemsSource = items, MinHeight = 30 }; Bind(control, Selector.SelectedIndexProperty, source, path); parent.Children.Add(control); }
        internal static void Number(Panel parent, object source, string path, string caption, int min, int max)
        { var editor = new NumberEditor { Caption = caption, Minimum = min, Maximum = max }; Bind(editor, NumberEditor.ValueProperty, source, path); parent.Children.Add(editor); }
        internal static Button Button(Panel parent, string caption, Action action)
        { var button = new Button { Content = caption, Margin = new Thickness(0, 0, 8, 0) }; button.Click += (s,e) => action(); parent.Children.Add(button); return button; }
        internal static ScrollViewer Scroll(UIElement child) => new ScrollViewer { Content = child, VerticalScrollBarVisibility = ScrollBarVisibility.Auto, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
        internal static StackPanel RenameControls(RenameOptions o, bool includeScope)
        {
            var panel = new StackPanel();
            if (includeScope) { panel.Children.Add(Heading("处理范围")); Check(panel, o, "AllWorkbooks", "所有已打开工作簿（同一 Excel）"); }
            Check(panel, o, "IncludeHidden", "包含隐藏页面"); Check(panel, o, "IncludeCharts", "包含图表页");
            panel.Children.Add(Heading("查找与替换")); Text(panel, o, "Find", "查找（留空时仅处理前后缀与编号）"); Text(panel, o, "Replacement", "替换为（正则支持 $1、${name}）");
            Check(panel, o, "Regex", "正则表达式"); Check(panel, o, "MatchCase", "区分大小写"); Check(panel, o, "FirstOnly", "只替换首次匹配"); Check(panel, o, "MatchesOnly", "仅处理查找匹配的页面");
            panel.Children.Add(Heading("名称整理")); Text(panel, o, "Prefix", "前缀"); Text(panel, o, "Suffix", "后缀"); Check(panel, o, "Trim", "去除原名称首尾空格");
            Choice(panel, o, "LetterCase", "字母大小写", "保持原样", "大写", "小写"); Check(panel, o, "Truncate", "超过31字符时截断（仍校验重名）");
            panel.Children.Add(Heading("连续编号")); Check(panel, o, "Numbering", "添加编号"); Number(panel, o, "Start", "起始值", 0, 1000000); Number(panel, o, "Step", "步长", 1, 1000); Number(panel, o, "Digits", "最少位数", 1, 10);
            Text(panel, o, "Separator", "编号分隔符"); Check(panel, o, "NumberAtEnd", "编号放在名称末尾"); Check(panel, o, "RestartPerBook", "每个工作簿重新编号"); return panel;
        }
    }
}
