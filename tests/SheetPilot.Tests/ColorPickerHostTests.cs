using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using SheetPilot.Core;
using SheetPilot.UI;
using Forms = System.Windows.Forms;

internal static partial class Program
{
    // Run in its own process: no WPF Window/Application may warm up the native-host path.
    private static void ColorPickerHostTests()
    {
        Check(Application.Current == null && !PickerWindows().Any(), "picker: cold process without a WPF application/window");
        var settings = new NavigationSettings();
        using (var model = new NavigationModel(settings))
        using (var view = new NavigationView(model))
        using (var form = new Forms.Form { Text = "GridPilot native-host regression", Width = 480, Height = 850, ShowInTaskbar = false })
        using (var host = new System.Windows.Forms.Integration.ElementHost { Dock = Forms.DockStyle.Fill, Child = view })
        {
            model.IsSettingsOpen = true;
            form.Controls.Add(host);
            try
            {
                form.Show();
                PumpUntil(() => view.IsLoaded);
                view.UpdateLayout();
                var editors = PickerDescendants(view).OfType<ColorEditor>().ToArray();
                Check(editors.Length == 5, "picker: all five navigation color editors are present");
                foreach (var editor in editors)
                {
                    editor.BringIntoView();
                    view.UpdateLayout();
                    ExerciseNativePicker(editor, form.Handle, dialog =>
                    {
                        var palette = PickerDescendants(dialog).OfType<ColorPalette>().Single();
                        Check(palette.Value == editor.Value, "picker: starts at current navigation color");
                        var preset = PickerDescendants(palette).OfType<Button>().First(b => (string)b.ToolTip == "#2563EB");
                        preset.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        Check(editor.Value == "#2563EB", "picker: preset updates bound editor immediately");
                        var property = System.Windows.Data.BindingOperations.GetBinding(editor, ColorEditor.ValueProperty).Path.Path.Split('.').Last();
                        Check((string)typeof(NavigationSettings).GetProperty(property).GetValue(settings) == "#2563EB",
                            "picker: preset reaches navigation settings");
                        Check(ReferenceEquals(dialog.FindResource("SurfaceBrush"), view.FindResource("SurfaceBrush")),
                            "picker: native dialog shares live navigation theme");
                        var hex = PickerDescendants(palette).OfType<TextBox>().Single();
                        hex.Focus(); hex.Text = "#123456";
                        PickerDescendants(dialog).OfType<Button>().Single(b => b.IsDefault).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    });
                    Check(editor.Value == "#123456", "picker: closing commits a valid HEX draft");
                }
                ExerciseNativePicker(editors[0], form.Handle, dialog =>
                {
                    dialog.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice, PresentationSource.FromVisual(dialog),
                        Environment.TickCount, Key.Escape) { RoutedEvent = Keyboard.PreviewKeyDownEvent });
                });
                Check(editors[0].Value == "#123456", "picker: Escape closes without reverting live choices");
                {
                    var prefs = new PreferencesWindow(new ToolPreferences(), settings, () => "Saved", () => {});
                    try
                    {
                        prefs.Show(); PumpUntil(() => prefs.IsLoaded); prefs.UpdateLayout();
                        var generalEditor = PickerDescendants(prefs).OfType<ColorEditor>().First();
                        PickerButton(generalEditor).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                        var popup = generalEditor.Children.OfType<Popup>().Single();
                        PumpUntil(() => popup.IsOpen && popup.Child.IsVisible);
                        Check(!PickerWindows().Any(w => w.GetType().Name == "ColorPickerWindow"),
                            "picker: general settings retains its inline popup");
                        PickerDescendants(popup.Child).OfType<ColorPalette>().Single().SetCurrentValue(ColorPalette.ValueProperty, "#ABCDEF");
                        Check(settings.BackgroundColor == "#ABCDEF", "picker: general popup still writes colors");
                        popup.IsOpen = false;
                    }
                    finally { prefs.Close(); }
                }
                editors[0].BringIntoView(); view.UpdateLayout();
                ExerciseNativePicker(editors[0], form.Handle, dialog => dialog.Close());
                ExerciseNativePicker(editors[0], form.Handle, dialog => host.Child = null);
                Check(!PickerWindows().Any(), "picker: unloading navigation closes its picker without orphan windows");
            }
            finally { host.Child = null; form.Close(); }
        }
        Console.WriteLine("PASS " + checks + " cold native-host assertions");
    }

    private static void ExerciseNativePicker(ColorEditor editor, IntPtr owner, Action<Window> exercise)
    {
        Exception failure = null;
        bool inspected = false;
        var watchdog = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
        int attempts = 0;
        watchdog.Tick += (s, e) =>
        {
            var dialog = PickerWindows().FirstOrDefault(w => w.GetType().Name == "ColorPickerWindow");
            if (dialog == null) { if (++attempts >= 30) watchdog.Stop(); return; }
            watchdog.Stop();
            try
            {
                inspected = true;
                Check(dialog.IsVisible && dialog.ActualWidth > 250 && dialog.ActualHeight > 250,
                    "picker: cold native-host click creates a visible color surface");
                Check(GetPickerWindow(new WindowInteropHelper(dialog).Handle, 4) == owner,
                    "picker: native owner is the top-level frame");
                Check(!IsPickerWindowEnabled(owner), "picker: owner is disabled during modal selection");
                Check(dialog.IsKeyboardFocusWithin, "picker: palette receives keyboard focus");
                exercise(dialog);
            }
            catch (Exception ex) { failure = ex; }
            finally { if (dialog.IsVisible) dialog.Close(); }
        };
        watchdog.Start();
        try { PickerButton(editor).RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent)); }
        finally { watchdog.Stop(); }
        if (failure != null) throw failure;
        Check(inspected, "picker: opens without warming up general settings");
        Check(IsPickerWindowEnabled(owner), "picker: owner re-enabled after close");
    }

    private static Button PickerButton(ColorEditor editor) =>
        ((DockPanel)editor.Children[0]).Children.OfType<Button>().Single();
    private static IEnumerable<Window> PickerWindows() => PresentationSource.CurrentSources.OfType<HwndSource>()
        .Select(s => s.RootVisual).OfType<Window>();
    private static IEnumerable<DependencyObject> PickerDescendants(DependencyObject parent)
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            yield return child;
            foreach (var descendant in PickerDescendants(child)) yield return descendant;
        }
    }
    [DllImport("user32.dll", EntryPoint = "GetWindow")] private static extern IntPtr GetPickerWindow(IntPtr hwnd, uint command);
    [DllImport("user32.dll", EntryPoint = "IsWindowEnabled")] private static extern bool IsPickerWindowEnabled(IntPtr hwnd);
}
