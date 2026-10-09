using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    // The text is a draft until Enter or focus loss. Empty/intermediate edits never change settings.
    public sealed class NumberEditor : StackPanel
    {
        public static readonly DependencyProperty ValueProperty = DependencyProperty.Register(nameof(Value), typeof(int), typeof(NumberEditor), new FrameworkPropertyMetadata(0, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, Changed));
        public static readonly DependencyProperty MinimumProperty = DependencyProperty.Register(nameof(Minimum), typeof(int), typeof(NumberEditor), new PropertyMetadata(0));
        public static readonly DependencyProperty MaximumProperty = DependencyProperty.Register(nameof(Maximum), typeof(int), typeof(NumberEditor), new PropertyMetadata(100));
        public static readonly DependencyProperty StepProperty = DependencyProperty.Register(nameof(Step), typeof(int), typeof(NumberEditor), new PropertyMetadata(1));
        public int Value { get => (int)GetValue(ValueProperty); set => SetValue(ValueProperty, value); }
        public int Minimum { get => (int)GetValue(MinimumProperty); set => SetValue(MinimumProperty, value); }
        public int Maximum { get => (int)GetValue(MaximumProperty); set => SetValue(MaximumProperty, value); }
        public int Step { get => (int)GetValue(StepProperty); set => SetValue(StepProperty, value); }
        private readonly TextBlock label = new TextBlock { VerticalAlignment = VerticalAlignment.Center };
        private readonly TextBox input = new TextBox { Width = 66, Padding = new Thickness(6, 4, 6, 4), HorizontalContentAlignment = HorizontalAlignment.Right };
        private readonly TextBlock error = new TextBlock { FontSize = 10, TextWrapping = TextWrapping.Wrap, Visibility = Visibility.Collapsed };
        public string Caption { get => label.Text; set { label.Text = value; System.Windows.Automation.AutomationProperties.SetName(input, value); } }
        public NumberEditor()
        {
            Margin = new Thickness(0, 8, 0, 0);
            var row = new DockPanel(); DockPanel.SetDock(input, Dock.Right); row.Children.Add(input); row.Children.Add(label); Children.Add(row);
            var slider = new Slider { Margin = new Thickness(0, 3, 0, 0), IsSnapToTickEnabled = true };
            slider.SetBinding(Slider.MinimumProperty, new Binding(nameof(Minimum)) { Source = this });
            slider.SetBinding(Slider.MaximumProperty, new Binding(nameof(Maximum)) { Source = this });
            slider.SetBinding(Slider.TickFrequencyProperty, new Binding(nameof(Step)) { Source = this });
            slider.SetBinding(Slider.ValueProperty, new Binding(nameof(Value)) { Source = this, Mode = BindingMode.TwoWay });
            Children.Add(slider); Children.Add(error);
            error.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
            input.LostKeyboardFocus += (s,e) => CommitInput();
            input.PreviewKeyDown += (s,e) => { if (e.Key == Key.Enter) { CommitInput(); e.Handled = true; } else if (e.Key == Key.Escape) { Restore(); e.Handled = true; } };
            Restore();
        }
        private static void Changed(DependencyObject d, DependencyPropertyChangedEventArgs e) => ((NumberEditor)d).Restore();
        private void Restore() { input.Text = Value.ToString(CultureInfo.InvariantCulture); error.Visibility = Visibility.Collapsed; }
        public bool CommitInput()
        {
            if (!PreferenceInput.TryInteger(input.Text, Minimum, Maximum, out int value))
            { error.Text = "请输入整数（" + Minimum + "–" + Maximum + "），原值未改变。"; error.Visibility = Visibility.Visible; return false; }
            SetCurrentValue(ValueProperty, value); Restore(); return true;
        }
    }
}
