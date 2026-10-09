using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public sealed class ColorPalette : StackPanel
    {
        public static readonly DependencyProperty ValueProperty = DependencyProperty.Register(nameof(Value), typeof(string), typeof(ColorPalette),
            new FrameworkPropertyMetadata("#42AD91", FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, (d,e) => ((ColorPalette)d).Sync()));
        public string Value { get => (string)GetValue(ValueProperty); set => SetValue(ValueProperty, value); }
        private readonly ColorPlane plane;
        private readonly ColorPlane hue;
        private readonly TextBox hex;
        private bool syncing;
        private double h, s, v;
        public ColorPalette()
        {
            Width = 260;
            Children.Add(new TextBlock { Text = "选择颜色", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0,0,0,12) });
            plane = new ColorPlane(false) { Height = 168, Focusable = true };
            hue = new ColorPlane(true) { Height = 22, Margin = new Thickness(0,12,0,12), Focusable = true };
            System.Windows.Automation.AutomationProperties.SetName(plane, "饱和度与明度色盘");
            System.Windows.Automation.AutomationProperties.SetName(hue, "色相条");
            plane.Picked += (x,y) => { s=x; v=1-y; Publish(); };
            hue.Picked += (x,y) => { h=x*360; Publish(); };
            Children.Add(plane); Children.Add(hue);
            var presets = new System.Windows.Controls.Primitives.UniformGrid { Rows=1, Columns=10, Margin=new Thickness(0,0,0,12) };
            foreach (string color in new[] { "#23825A", "#42AD91", "#2563EB", "#7C3AED", "#DB2777", "#DC2626", "#F59E0B", "#F8FAFC", "#64748B", "#111827" })
            {
                var dot = new Border { Background=(Brush)new BrushConverter().ConvertFromString(color), Width=18, Height=18, CornerRadius=new CornerRadius(9) };
                var button = new Button { Content=dot, Padding=new Thickness(2), BorderThickness=new Thickness(0), ToolTip=color };
                System.Windows.Automation.AutomationProperties.SetName(button, "选择颜色 " + color);
                button.Click += (a,b) => SetCurrentValue(ValueProperty, color); presets.Children.Add(button);
            }
            Children.Add(presets);
            var row = new DockPanel(); var label = new TextBlock { Text="HEX", Width=40, VerticalAlignment=VerticalAlignment.Center }; row.Children.Add(label);
            hex = new TextBox { Padding=new Thickness(8,6,8,6), MaxLength=9 };
            System.Windows.Automation.AutomationProperties.SetName(hex,"颜色十六进制值");
            hex.LostKeyboardFocus += (a,b) => Commit(); hex.KeyDown += (a,b) => { if(b.Key==Key.Enter) { Commit(); b.Handled=true; } };
            row.Children.Add(hex); Children.Add(row); Sync();
        }
        private void Commit() { if(PreferenceInput.TryColor(hex.Text,out string value)) SetCurrentValue(ValueProperty,value); else hex.Text=Value; }
        private void Sync()
        {
            if (plane==null || syncing) return;
            Color color = (Color)ColorConverter.ConvertFromString(Value);
            double r=color.R/255.0,g=color.G/255.0,b=color.B/255.0,max=Math.Max(r,Math.Max(g,b)),min=Math.Min(r,Math.Min(g,b)),delta=max-min;
            v=max; s=max==0?0:delta/max;
            if(delta>0) h=(((max==r?(g-b)/delta:max==g?(b-r)/delta+2:(r-g)/delta+4)*60)+360)%360;
            Paint();
        }
        private void Publish() { syncing=true; try { SetCurrentValue(ValueProperty, FromHsv(h,s,v).ToString().Remove(1,2)); Paint(); } finally { syncing=false; } }
        private void Paint() { plane.Hue=h; plane.X=s; plane.Y=1-v; hue.X=h/360; plane.InvalidateVisual(); hue.InvalidateVisual(); hex.Text=Value; }
        public static Color FromHsv(double hue,double saturation,double value)
        {
            double c=value*saturation,x=c*(1-Math.Abs((hue/60)%2-1)),m=value-c,r=0,g=0,b=0;
            hue=(hue%360+360)%360;
            if(hue<60){r=c;g=x;}else if(hue<120){r=x;g=c;}else if(hue<180){g=c;b=x;}else if(hue<240){g=x;b=c;}else if(hue<300){r=x;b=c;}else{r=c;b=x;}
            return Color.FromRgb((byte)Math.Round((r+m)*255),(byte)Math.Round((g+m)*255),(byte)Math.Round((b+m)*255));
        }
        private sealed class ColorPlane : FrameworkElement
        {
            private readonly bool strip;
            public double Hue,X,Y;
            public event Action<double,double> Picked;
            public ColorPlane(bool strip)
            {
                this.strip=strip; Cursor=Cursors.Cross;
                MouseLeftButtonDown += (a,b) => { Focus(); CaptureMouse(); Pick(b.GetPosition(this)); b.Handled=true; };
                MouseMove += (a,b) => { if(IsMouseCaptured) Pick(b.GetPosition(this)); };
                MouseLeftButtonUp += (a,b) => { if(IsMouseCaptured) { Pick(b.GetPosition(this)); ReleaseMouseCapture(); } };
                KeyDown += (a,b) => { double step=(Keyboard.Modifiers & ModifierKeys.Shift)!=0?.1:.01; if(b.Key==Key.Left)X-=step;else if(b.Key==Key.Right)X+=step;else if(!strip && b.Key==Key.Up)Y-=step;else if(!strip && b.Key==Key.Down)Y+=step;else return; Picked?.Invoke(Clamp(X),Clamp(Y)); b.Handled=true; };
            }
            private static double Clamp(double value)=>Math.Max(0,Math.Min(1,value));
            private void Pick(Point point) { X=Clamp(point.X/ActualWidth); Y=Clamp(point.Y/ActualHeight); Picked?.Invoke(X,Y); }
            protected override void OnRender(DrawingContext dc)
            {
                var rect=new Rect(0,0,ActualWidth,ActualHeight);
                if(strip)
                {
                    var gradient=new LinearGradientBrush { StartPoint=new Point(0,0),EndPoint=new Point(1,0) };
                    for(int i=0;i<=6;i++)gradient.GradientStops.Add(new GradientStop(FromHsv(i*60,1,1),i/6.0));
                    dc.DrawRoundedRectangle(gradient,null,rect,5,5);
                    dc.DrawRoundedRectangle(null,new Pen(Brushes.White,3),new Rect(Math.Max(2,Math.Min(ActualWidth-8,X*ActualWidth-3)),2,6,ActualHeight-4),3,3);
                }
                else
                {
                    dc.DrawRectangle(new SolidColorBrush(FromHsv(Hue,1,1)),null,rect);
                    dc.DrawRectangle(new LinearGradientBrush(Colors.White,Colors.Transparent,new Point(0,0),new Point(1,0)),null,rect);
                    dc.DrawRectangle(new LinearGradientBrush(Colors.Transparent,Colors.Black,new Point(0,0),new Point(0,1)),null,rect);
                    var point=new Point(Math.Max(6,Math.Min(ActualWidth-6,X*ActualWidth)),Math.Max(6,Math.Min(ActualHeight-6,Y*ActualHeight)));
                    dc.DrawEllipse(null,new Pen(Brushes.Black,3),point,5,5); dc.DrawEllipse(null,new Pen(Brushes.White,2),point,5,5);
                }
            }
        }
    }
}
