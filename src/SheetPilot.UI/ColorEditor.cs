using System;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Controls.Primitives;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public sealed class ColorEditor : StackPanel
    {
        public static readonly DependencyProperty ValueProperty = DependencyProperty.Register(nameof(Value), typeof(string), typeof(ColorEditor), new FrameworkPropertyMetadata("#FFFFFF", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, Changed));
        public string Value { get => (string)GetValue(ValueProperty); set => SetValue(ValueProperty, value); }
        private ColorPickerWindow nativePicker;
        private readonly TextBlock label = new TextBlock { VerticalAlignment = VerticalAlignment.Center };
        private readonly TextBox input = new TextBox { Width = 94, Padding = new Thickness(6, 5, 6, 5) };
        private readonly TextBlock error = new TextBlock { FontSize = 10, TextWrapping = TextWrapping.Wrap, Visibility = Visibility.Collapsed };
        public string Caption { get => label.Text; set { label.Text = value; System.Windows.Automation.AutomationProperties.SetName(input, value); } }
        public ColorEditor()
        {
            Margin = new Thickness(0, 7, 0, 0);
            var row = new DockPanel(); DockPanel.SetDock(input, Dock.Right); row.Children.Add(input);
            var swatch = new Border { Width = 24, Height = 20, CornerRadius = new CornerRadius(4), BorderThickness = new Thickness(1) };
            swatch.SetResourceReference(Border.BorderBrushProperty, "BorderBrush");
            swatch.SetBinding(Border.BackgroundProperty, new Binding(nameof(Value)) { Source = this });
            var picker = new Button { Content=swatch, Padding=new Thickness(4), Margin=new Thickness(6,0,8,0), ToolTip="打开色盘" };
            System.Windows.Automation.AutomationProperties.SetName(picker,"打开色盘");
            var palette = new ColorPalette(); palette.SetBinding(ColorPalette.ValueProperty,new Binding(nameof(Value)) { Source=this,Mode=BindingMode.TwoWay });
            var surface = new Border { Child=palette,Padding=new Thickness(16),CornerRadius=new CornerRadius(10),BorderThickness=new Thickness(1) };
            surface.SetResourceReference(Border.BackgroundProperty,"SurfaceBrush"); surface.SetResourceReference(Border.BorderBrushProperty,"BorderBrush");
            var popup = new Popup { Child=surface,PlacementTarget=picker,Placement=PlacementMode.Bottom,AllowsTransparency=true,StaysOpen=false };
            picker.Click += (s,e) =>
            {
                if (Window.GetWindow(this) != null) popup.IsOpen = !popup.IsOpen;
                else OpenNativePicker(picker);
            };
            popup.PreviewKeyDown += (s,e) => { if(e.Key==Key.Escape) { popup.IsOpen=false; e.Handled=true; } };
            Unloaded += (s,e) => { popup.IsOpen = false; nativePicker?.Close(); };
            DockPanel.SetDock(picker, Dock.Right); row.Children.Add(picker); row.Children.Add(label); Children.Add(row); Children.Add(error); Children.Add(popup);
            error.SetResourceReference(TextBlock.ForegroundProperty, "TextBrush");
            input.LostKeyboardFocus += (s,e) => CommitInput();
            input.PreviewKeyDown += (s,e) => { if (e.Key == Key.Enter) { CommitInput(); e.Handled = true; } else if (e.Key == Key.Escape) { Restore(); e.Handled = true; } };
            Restore();
        }
        private void OpenNativePicker(Button picker)
        {
            if (nativePicker != null || !IsLoaded || !IsVisible) return;
            var source = PresentationSource.FromVisual(this) as HwndSource;
            if (source == null || source.IsDisposed) return;
            // An Office task pane has an ElementHost/HwndSource, not a WPF Window.
            // A capture-based Popup relies on the native host's modeless input loop.
            // Use an explicitly owned modal window so cold starts do not depend on
            // a PreferencesWindow having initialized WPF window/input handling first.
            try
            {
                nativePicker = new ColorPickerWindow(this, source.Handle);
                nativePicker.ShowDialog();
            }
            catch (System.ComponentModel.Win32Exception ex) { ReportPickerError(ex); }
            catch (InvalidOperationException ex) { ReportPickerError(ex); }
            finally
            {
                nativePicker?.Close();
                nativePicker = null;
                if (IsLoaded && IsVisible) picker.Focus();
            }
        }
        private void ReportPickerError(Exception ex)
        {
            System.Diagnostics.Trace.TraceError("GridPilot color picker: {0}", ex);
            error.Text = "无法打开选色窗口，请关闭其他对话框后重试。";
            error.Visibility = Visibility.Visible;
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
