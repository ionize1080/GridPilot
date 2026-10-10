using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace SheetPilot.UI
{
    // Native Office/ElementHost fallback; ordinary WPF windows retain the inline popup.
    internal sealed class ColorPickerWindow : Window
    {
        internal ColorPickerWindow(ColorEditor editor, IntPtr hostHandle)
        {
            IntPtr owner = GetAncestor(hostHandle, 2); // GA_ROOT: Excel frame, not the child ElementHost.
            if (owner == IntPtr.Zero) throw new InvalidOperationException("The color editor has no native owner.");
            new WindowInteropHelper(this).Owner = owner;
            Title = "GridPilot · 选择颜色";
            ShowInTaskbar = false;
            ResizeMode = ResizeMode.NoResize;
            SizeToContent = SizeToContent.WidthAndHeight;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            FontFamily = TextElement.GetFontFamily(editor);
            FontSize = TextElement.GetFontSize(editor);

            // A top-level window is outside the editor's resource tree. Share the
            // existing dictionaries (root first), including live navigation colors.
            var resources = new Stack<ResourceDictionary>();
            for (DependencyObject current = editor; current != null;)
            {
                if (current is FrameworkElement element) resources.Push(element.Resources);
                current = LogicalTreeHelper.GetParent(current) ??
                    (current is Visual ? VisualTreeHelper.GetParent(current) : null);
            }
            foreach (var dictionary in resources) Resources.MergedDictionaries.Add(dictionary);
            SetResourceReference(BackgroundProperty, "SurfaceBrush");
            SetResourceReference(ForegroundProperty, "TextBrush");

            var content = new StackPanel { Margin = new Thickness(16) };
            var palette = new ColorPalette();
            palette.SetBinding(ColorPalette.ValueProperty, new Binding(nameof(ColorEditor.Value))
            {
                Source = editor, Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            });
            content.Children.Add(palette);
            var done = new Button
            {
                Content = "完成", IsDefault = true, HorizontalAlignment = HorizontalAlignment.Right,
                MinWidth = 72, Margin = new Thickness(0, 16, 0, 0)
            };
            done.Click += (s, e) => Close();
            content.Children.Add(done);
            Content = content;
            ContentRendered += (s, e) => palette.MoveFocus(new TraversalRequest(FocusNavigationDirection.First));
            PreviewKeyDown += (s, e) =>
            {
                if (e.Key == Key.Escape) { e.Handled = true; Close(); }
            };
            // Like the popup, colors apply immediately. Closing also commits a
            // valid HEX draft through the palette's existing LostKeyboardFocus handler.
            Closing += (s, e) => { FocusManager.SetFocusedElement(this, this); Keyboard.ClearFocus(); };
            Closed += (s, e) =>
            {
                BindingOperations.ClearBinding(palette, ColorPalette.ValueProperty);
                Content = null;
                Resources.MergedDictionaries.Clear();
            };
        }

        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
    }
}
