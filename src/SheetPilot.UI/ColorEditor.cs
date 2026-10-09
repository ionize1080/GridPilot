using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public sealed class ColorEditor : StackPanel
    {
        public static readonly DependencyProperty ValueProperty = DependencyProperty.Register(nameof(Value), typeof(string), typeof(ColorEditor), new FrameworkPropertyMetadata("#FFFFFF", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, Changed));
        public string Value { get => (string)GetValue(ValueProperty); set => SetValue(ValueProperty, value); }
        private readonly TextBlock label = new TextBlock { VerticalAlignment = VerticalAlignment.Center };
        private readonly TextBox input = new TextBox { Width = 94, Padding = new Thickness(6, 5, 6, 5) };
        private readonly TextBlock error = new TextBlock { FontSize = 10, TextWrapping = TextWrapping.Wrap, Visibility = Visibility.Collapsed };
        public string Caption { get => label.Text; set { label.Text = value; System.Windows.Automation.AutomationProperties.SetName(input, value); } }
        public ColorEditor()
        {
            Margin = new Thickness(0, 7, 0, 0);
            var row = new DockPanel(); DockPanel.SetDock(input, Dock.Right); row.Children.Add(input);
            var swatch = new Border { Width = 20, Height = 20, CornerRadius = new CornerRadius(5), BorderThickness = new Thickness(1), Margin = new Thickness(6, 0, 8, 0) };
            swatch.SetResourceReference(Border.BorderBrushProperty, "BorderBrush");
            swatch.SetBinding(Border.BackgroundProperty, new Binding(nameof(Value)) { Source = this });
            DockPanel.SetDock(swatch, Dock.Right); row.Children.Add(swatch); row.Children.Add(label); Children.Add(row); Children.Add(error);
            error.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
            input.LostKeyboardFocus += (s,e) => CommitInput();
            input.PreviewKeyDown += (s,e) => { if (e.Key == Key.Enter) { CommitInput(); e.Handled = true; } else if (e.Key == Key.Escape) { Restore(); e.Handled = true; } };
            Restore();
        }
        private static void Changed(DependencyObject d, DependencyPropertyChangedEventArgs e) => ((ColorEditor)d).Restore();
        private void Restore() { input.Text = Value; error.Visibility = Visibility.Collapsed; }
        public bool CommitInput()
        {
            if (!PreferenceInput.TryColor(input.Text, out string value)) { error.Text = "请输入十六进制颜色，如 #23825A。"; error.Visibility = Visibility.Visible; return false; }
            SetCurrentValue(ValueProperty, value); Restore(); return true;
        }
    }
}
