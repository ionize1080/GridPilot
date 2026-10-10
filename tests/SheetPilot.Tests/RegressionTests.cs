using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Windows.Media;
using System.Windows.Threading;
using SheetPilot.Core;
using SheetPilot.UI;

internal static partial class Program
{
    private static void RegressionTests()
    {
        using(var model=new NavigationModel(new NavigationSettings {ScopeIndex=1}))
        using(var view=new NavigationView(model))
        {
            Func<bool,WorkbookSnapshot> snapshot=second=>new WorkbookSnapshot(second?"b":"a","Books",new[]{
                new SheetItem("same","Summary",1,SheetVisibility.Visible,!second,"工作表","a","A.xlsx"),
                new SheetItem("same","Summary",1,SheetVisibility.Visible,second,"工作表","b","B.xlsx")});
            model.Update(snapshot(false)); PumpUntil(()=>model.Results.Count==2);
            view.Measure(new System.Windows.Size(350,850));view.Arrange(new System.Windows.Rect(0,0,350,850));view.UpdateLayout();
            for(int i=0;i<20;i++) {bool second=i%2==0;model.Update(snapshot(second));Check(model.Selected.WorkbookKey==(second?"b":"a") && model.Results.Count(s=>s.IsActive)==1,"navigation: duplicate sheet identities follow active workbook immediately");}
            Check(ReferenceEquals(((System.Windows.Controls.ListBox)view.FindName("SheetList")).SelectedItem,model.Selected),"navigation: actual ListBox binding retains the correct active item");
            model.Query="Other"; PumpUntil(()=>model.Results.Count==0);model.Update(snapshot(true));Check(model.Selected==null,"navigation: filtered active sheet cannot leave stale selection");
        }
        var prefs=new PreferencesWindow(new ToolPreferences(),new NavigationSettings(),()=>"Saved",()=>{});
        var content=(System.Windows.FrameworkElement)prefs.Content;
        content.Measure(new System.Windows.Size(940,740));content.Arrange(new System.Windows.Rect(0,0,940,740));content.UpdateLayout();
        foreach(var id in new[]{"focus","conversion","copy","rename","about","navigation","general"})
        {
            ((System.Windows.Controls.Primitives.ToggleButton)prefs.FindName("Category_"+id)).RaiseEvent(new System.Windows.RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
            content.UpdateLayout();
            Check(prefs.ActiveSection==id && (id=="general" ? prefs.ScrollOffset<1 : prefs.ScrollOffset>0),"settings: category "+id+" scrolls to matching section");
        }
        var scroll=(System.Windows.Controls.ScrollViewer)prefs.FindName("SettingsScroll");scroll.ScrollToVerticalOffset(900);content.UpdateLayout();
        Check(prefs.ActiveSection!="general","settings: continuous scroll updates the sidebar selection");
        prefs.Close();
        Check(ColorPalette.FromHsv(0,1,1)==Colors.Red && ColorPalette.FromHsv(120,1,1)==Colors.Lime && ColorPalette.FromHsv(240,1,1)==Colors.Blue,"palette: hue primaries");
        Check(ColorPalette.FromHsv(30,0,1)==Colors.White && ColorPalette.FromHsv(30,1,0)==Colors.Black,"palette: saturation and brightness limits");
    }
    private static void InteractiveTests(string addIn,string directory)
    {
        Directory.CreateDirectory(directory);
        dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));
        int pid;GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out pid);
        try
        {
            app.Visible=true; app.DisplayAlerts=false;
            dynamic a=app.Workbooks.Add(-4167);a.Sheets.Item[1].Name="总览";
            foreach(string name in new[]{"销售2025","预算2025","区域数据","图表与说明"})a.Worksheets.Add().Name=name;
            a.SaveAs(Path.Combine(directory,"GridPilot-QA-A.xlsx"),51);
            dynamic b=app.Workbooks.Add(-4167);b.Sheets.Item[1].Name="总览";
            foreach(string name in new[]{"销售2025","预算2025","区域数据"})b.Worksheets.Add().Name=name;
            b.SaveAs(Path.Combine(directory,"GridPilot-QA-B.xlsx"),51);
            dynamic sample=app.Workbooks.Open(Path.Combine(directory,"Financial-Sample.xlsx"),0,true);
            Console.WriteLine("Sample: "+sample.Name+"; used rows="+sample.Sheets.Item[1].UsedRange.Rows.Count+"; columns="+sample.Sheets.Item[1].UsedRange.Columns.Count);
            app.Goto(a.Sheets["总览"].Range["C3"],false);
            for(int i=1;i<=100;i++) {a.Sheets["总览"].Cells[i,2].Value2="项目 "+i;a.Sheets["总览"].Cells[i,3].Value2=i*1380;}
            a.Sheets["总览"].Columns["B:D"].ColumnWidth=18;
            app.RegisterXLL(addIn);app.Run("GridPilot_ShowNavigation");
            File.WriteAllText(Path.Combine(directory,"ready.txt"),pid.ToString());
            var timer=new DispatcherTimer {Interval=TimeSpan.FromMilliseconds(250)};
            var frame=new DispatcherFrame();
            timer.Tick+=(s,e)=>
            {
                string command=Path.Combine(directory,"command.txt");
                if(!File.Exists(command))return;string text=File.ReadAllText(command).Trim();File.Delete(command);
                try
                {
                    if(text=="close")frame.Continue=false;
                    else if(text=="switch-close") { app.Goto(b.Sheets["总览"].Range["C3"],false); b.Close(false); Console.WriteLine("In-host lifecycle: activated and closed owned workbook B without saving"); }
                    else if(text=="diagnostics") {app.Run("GridPilot_Diagnostics");File.Copy(Path.Combine(Path.GetTempPath(),"GridPilot-"+pid+".log"),Path.Combine(directory,"host-diagnostics.log"),true);}
                    else if(text=="benchmark")
                    {
                        dynamic page=sample.Sheets.Item[1];app.Goto(page.Range["F10"],false);
                        var elapsed=Stopwatch.StartNew();for(int i=0;i<120;i++)page.Cells[10+i%24,6+i%8].Select();
                        Console.WriteLine("In-host focus benchmark: 120 selection operations, total="+elapsed.Elapsed.TotalMilliseconds.ToString("F1")+"ms (includes external COM and Excel work)");
                        app.Run("GridPilot_Diagnostics");File.Copy(Path.Combine(Path.GetTempPath(),"GridPilot-"+pid+".log"),Path.Combine(directory,"host-diagnostics.log"),true);
                    }
                    else if(text=="state")File.WriteAllText(Path.Combine(directory,"state.txt"),app.ActiveWorkbook.Name+" | "+app.ActiveSheet.Name+" | "+app.ActiveCell.Address);
                }
                catch(Exception ex){Console.WriteLine(ex);}
            };
            timer.Start(); Dispatcher.PushFrame(frame); timer.Stop();
        }
        finally {app.DisplayAlerts=false;app.Quit();System.Runtime.InteropServices.Marshal.FinalReleaseComObject((object)app);}
    }
}
