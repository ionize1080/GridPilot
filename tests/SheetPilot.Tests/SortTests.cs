using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using SheetPilot.Core;
using SheetPilot.Excel;
using SheetPilot.UI;

internal static partial class Program
{
    private static void SortTests()
    {
        foreach(var pair in new[]{new[]{"零","0"},new[]{"十一","11"},new[]{"一百零二","102"},new[]{"二〇二四","2024"},new[]{"一万零一","10001"},new[]{"十二亿三千四百五十六万七千八百九十","1234567890"},new[]{"壹佰贰拾叁","123"},new[]{"一万亿","1000000000000"},new[]{"九载","900000000000000000000000000000000000000000000"}})
            Check(CustomSort.ChineseNumber(pair[0])==BigInteger.Parse(pair[1]),"Chinese numeral "+pair[0]);
        string[] values={"业务十部","国际十一部","国际二部","其他","国际一百零一部","业务一部","国际二部"};
        var options=new SortOptions{Rules="(国际)[中文数字]\n(业务)[中文数字]"};
        CustomSort.Order(values,options,out int matched);Check(matched==6,"preview reports matched and unmatched items");
        Check(CustomSort.Order(values,options).SequenceEqual(new[]{2,6,1,4,5,0,3}),"rule priority, arbitrary Chinese numbers, stable ties, unmatched last");
        options.Descending=true;Check(CustomSort.Order(values,options).SequenceEqual(new[]{4,1,2,6,0,5,3}),"descending within rule groups with stable ties");
        Check(CustomSort.Order(new[]{"小十","大二","大一百","微一","中一"},new SortOptions{Rules="[大中小微][中文数字]"}).SequenceEqual(new[]{1,2,4,0,3}),"multiple typed sequence keys");
        Check(CustomSort.Order(new[]{"小型十部","特大型二部","大型一部"},new SortOptions{Sequences="级别=特大型,大型,中型,小型",Rules="[级别][中文数字]"}).SequenceEqual(new[]{1,2,0}),"custom multi-character sequences");
        Check(CustomSort.Order(new[]{"A+二","A.一"},new SortOptions{Sequences="符号=A.,A+",Rules="[符号][中文数字]"}).SequenceEqual(new[]{1,0}),"sequence text escaped literally");
        Check(CustomSort.Order(new[]{"表10","表2","表1"},new SortOptions{Natural=true}).SequenceEqual(new[]{2,1,0}),"natural Arabic digit ordering");
        Reject(()=>CustomSort.Order(values,new SortOptions{Rules="("}),"invalid regex rejected");
        Reject(()=>CustomSort.Order(values,new SortOptions{Sequences="级别=大,大"}),"duplicate sequence rejected");
        Reject(()=>CustomSort.Order(new[]{"a"},new SortOptions{Rules="(a|[中文数字])"}),"optional missing sort token rejected");
        Reject(()=>CustomSort.Order(new[]{new string('a',10000)+"!"},new SortOptions{Rules="(a+)+$"}),"regex timeout bounded");
        var large=Enumerable.Range(1,20000).Reverse().Select(i=>"表"+i).ToArray();var watch=Stopwatch.StartNew();
        Check(CustomSort.Order(large,new SortOptions{Natural=true})[0]==19999,"20k natural sort");Console.WriteLine("Natural 20k ms="+watch.ElapsedMilliseconds);
        var prefs=new ToolPreferences();Check(!prefs.IsEnabled("GridPilot_Sort")&&prefs.IsEnabled("GridPilot_Focus")&&prefs.IsEnabled("SheetPilot_Show")&&prefs.IsEnabled("GridPilot_Settings"),"optional features default off; essentials available");
        string folder=Path.Combine(Path.GetTempPath(),"GridPilot-sort-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        var store=new ToolPreferencesStore(Path.Combine(folder,"prefs.json"));prefs.EnabledFeatures=new[]{"GridPilot_Sort","GridPilot_Undo"};prefs.Sort=options;prefs.SortContextMenu=false;
        Check(store.Save(prefs)&&store.Load().IsEnabled("GridPilot_Sort")&&!store.Load().SortContextMenu&&store.Load().Sort.Rules==options.Rules,"feature + menu + rules persistence");
        File.WriteAllText(Path.Combine(folder,"prefs.json"),"{\"SchemaVersion\":1,\"FocusOpacity\":25}");prefs=store.Load();Check(prefs.FocusOpacity==25&&!prefs.IsEnabled("GridPilot_Sort")&&prefs.Sort!=null,"v1.4 settings migrate retaining preferences");
        var icons=new[]{"SheetPilot_Show","SheetPilot_Refresh","GridPilot_Focus","GridPilot_Sort","GridPilot_Reconcile","GridPilot_Database","GridPilot_SimpleMask","GridPilot_Mask","GridPilot_Segment","GridPilot_CopyInternal","GridPilot_CopyValues","GridPilot_Rename","GridPilot_Text","GridPilot_Number","GridPilot_Undo","GridPilot_Settings"};
        var hashes=icons.Select(id=>{using(var image=RibbonIcons.Draw(id,48))using(var stream=new MemoryStream()){image.Save(stream,System.Drawing.Imaging.ImageFormat.Png);return Convert.ToBase64String(stream.ToArray());}}).ToArray();Check(hashes.Distinct().Count()==icons.Length,"all 16 icons are distinct");
        string docs=Path.GetFullPath("docs/v1.5.1");Directory.CreateDirectory(docs);
        using(var atlas=new System.Drawing.Bitmap(16*72,3*76))using(var g=System.Drawing.Graphics.FromImage(atlas))
        {g.Clear(System.Drawing.Color.White);for(int theme=0;theme<3;theme++)for(int i=0;i<icons.Length;i++)using(var icon=RibbonIcons.Draw(icons[i],48,theme)){g.DrawImage(icon,i*72+12,theme*76+8);}atlas.Save(Path.Combine(docs,"icons.png"));}
        var window=new SortWindow(new NavigationSettings(),options,"合成部门清单 · A1:D8 · 8 行 × 4 列",new[]{"部门","金额","编号","备注"},0,(c,h)=>values,(c,o)=>"已排序");
        window.Show();window.UpdateLayout();
        var example=PickerDescendants(window).OfType<System.Windows.Controls.ComboBox>().Single(c=>c.Name=="SortExamples");
        var useExample=PickerDescendants(window).OfType<System.Windows.Controls.Button>().Single(b=>Equals(b.Content,"套用示例"));
        useExample.RaiseEvent(new System.Windows.RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        Check(CustomSort.Order(new[]{"客户十一处","客户二处","客户一百零一处"},window.Options).SequenceEqual(new[]{1,0,2}),"customer example matches embedded Chinese numbers");
        example.SelectedIndex=4;useExample.RaiseEvent(new System.Windows.RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        Check(window.Options.Rules=="[阶段]"&&window.Options.Sequences=="阶段=准备,实施,验收"&&CustomSort.Order(new[]{"验收","准备","实施"},window.Options).SequenceEqual(new[]{1,2,0}),"custom example populates both inputs coherently");
        example.SelectedIndex=2;useExample.RaiseEvent(new System.Windows.RoutedEventArgs(System.Windows.Controls.Primitives.ButtonBase.ClickEvent));
        Check(window.Options.Sequences==""&&!window.Options.Natural,"built-in example clears unused custom definitions");
        window.UpdateLayout();Render(window,Path.Combine(docs,"sort.png"),(int)window.ActualWidth,(int)window.ActualHeight,96);window.Close();
        using(var model=new NavigationModel())using(var nav=new NavigationView(model))
        {
            var item=new SheetItem("one","原名称",1,SheetVisibility.Visible,true,"工作表","book","合成工作簿");
            model.Update(new WorkbookSnapshot("book","合成工作簿",new[]{item}));PumpUntil(()=>model.Results.Count==1);
            var host=new System.Windows.Window{Content=nav,Width=400,Height=600};host.Show();host.UpdateLayout();
            SheetEdit edited=null;model.EditRequested+=request=>edited=request;
            var flags=System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic;
            var rename=typeof(NavigationView).GetMethod("Rename",flags);var end=typeof(NavigationView).GetMethod("EndRename",flags);
            rename.Invoke(nav,new object[]{item});var box=(System.Windows.Controls.TextBox)typeof(NavigationView).GetField("editing",flags).GetValue(nav);
            Check(box.IsVisible&&box.Text=="原名称","inline rename edits the selected row");box.Text="新名称";end.Invoke(nav,new object[]{true});Check(edited?.Name=="新名称"&&edited.Action=="rename","inline rename commits one sheet edit");
            edited=null;rename.Invoke(nav,new object[]{item});box=(System.Windows.Controls.TextBox)typeof(NavigationView).GetField("editing",flags).GetValue(nav);box.Text="取消内容";end.Invoke(nav,new object[]{false});Check(edited==null,"inline rename cancellation does not edit");
            host.Close();
        }
    }
    private static void SortExcelTests()
    {
        dynamic app=null;var com=new MaskComLifetime();
        try
        {
            app=com.Hold(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true)));app.Visible=false;app.DisplayAlerts=false;
            dynamic books=com.Hold(app.Workbooks),book=com.Hold(books.Add(-4167)),sheet=com.Hold(book.Sheets[1]);
            var gateway=new ExcelGateway(()=>(object)app);dynamic data=com.Hold(sheet.Range["A1:C5"]);
            data.Value2=new object[,]{{"部门","金额","计算"},{"国际十",10,null},{"业务一",20,null},{"国际二",30,null},{"国际一百零一",40,null}};
            sheet.Range["C2:C5"].FormulaR1C1="=RC[-1]*2";sheet.Range["A2:C2"].Interior.Color=12345;sheet.Range["A1:C5"].AutoFilter();sheet.Range["A2"].Select();
            var options=new SortOptions{Rules="(国际)[中文数字]\n(业务)[中文数字]"};var snapshot=gateway.CaptureSort();gateway.ApplySort(snapshot,0,options);
            Check((string)sheet.Range["A2"].Value2=="国际二"&&(double)sheet.Range["B2"].Value2==30&&(double)sheet.Range["C2"].Value2==60,"native range sort moves intact records and formulas");
            Check((double)sheet.Range["A3"].Interior.Color==12345,"native sort carries formatting");
            Check((object)sheet.Range["D1"].Value2==null&&(bool)sheet.AutoFilterMode,"temporary key cleaned and filter retained");
            gateway.UndoLast();Check((string)sheet.Range["A2"].Value2=="国际十"&&(double)sheet.Range["C2"].Value2==20,"sort undo restores original order and formulas");
            snapshot=gateway.CaptureSort();sheet.Range["B2"].Value2=15;Reject(()=>gateway.ApplySort(snapshot,0,options),"stale preview rejected");sheet.Range["B2"].Value2=10;
            sheet.AutoFilterMode=false;dynamic table=com.Hold(sheet.ListObjects.Add(1,data,Type.Missing,1));sheet.Range["A2"].Select();
            snapshot=gateway.CaptureSort();gateway.ApplySort(snapshot,0,options);Check((string)sheet.Range["A2"].Value2=="国际二"&&(int)table.ListColumns.Count==3,"native Excel table sort and temporary column removal");gateway.UndoLast();
            table.ShowTotals=true;snapshot=gateway.CaptureSort();gateway.ApplySort(snapshot,0,options);Check((int)table.Range.Rows.Count==6&&(int)table.ListColumns.Count==3,"table totals excluded and retained");gateway.UndoLast();
            table.Unlist();sheet.Range["A1:C5"].AutoFilter();sheet.Range["A1:C5"].AutoFilter(2,">20");snapshot=gateway.CaptureSort();gateway.ApplySort(snapshot,0,options);Check((bool)sheet.FilterMode,"active filter retained after custom sort");gateway.UndoLast();sheet.ShowAllData();
            sheet.Protect();Reject(()=>gateway.CaptureSort(),"protected range rejected");sheet.Unprotect();
            dynamic merged=com.Hold(book.Worksheets.Add());merged.Name="合并测试";
            merged.Range["A1:C1"].Merge();merged.Range["A1"].Value2="合并标题";
            merged.Range["A2:C4"].Value2=new object[,]{{"客户十一处",null,11},{"客户二处",null,2},{"客户一处",null,1}};
            for(int r=2;r<=4;r++)merged.Range["A"+r+":B"+r].Merge();merged.Range["A2"].Select();
            var mergedOptions=new SortOptions{Rules=".*?[中文数字]"};snapshot=gateway.CaptureSort();gateway.ApplySort(snapshot,0,mergedOptions);
            Check((string)merged.Range["A2"].Value2=="客户一处"&&(double)merged.Range["C2"].Value2==1&&(bool)merged.Range["A2:B2"].MergeCells,"horizontal merges sort with intact rows");
            Check((string)merged.Range["A1"].Value2=="合并标题"&&(int)merged.Range["A1"].MergeArea.Columns.Count==3,"merged header excluded from sorting");
            gateway.UndoLast();Check((string)merged.Range["A2"].Value2=="客户十一处"&&(bool)merged.Range["A2:B2"].MergeCells,"merged row undo preserves merge layout");
            merged.Range["A3:B3"].UnMerge();merged.Range["B3"].Value2="未合并备注";merged.Range["D2:D4"].FormulaR1C1="=RC[-1]*2";
            snapshot=gateway.CaptureSort();gateway.ApplySort(snapshot,0,mergedOptions);
            Check((double)merged.Range["D2"].Value2==2&&(bool)merged.Range["A2:B2"].MergeCells&&(string)merged.Range["B3"].Value2=="未合并备注"&&!((bool)merged.Range["A3"].MergeCells),"mixed merged and unmerged rows retain notes formulas and layout");
            gateway.UndoLast();Check((double)merged.Range["D2"].Value2==22,"mixed merge undo restores formulas");
            merged.Range["A2:B4"].UnMerge();merged.Range["A2:A3"].Merge();snapshot=gateway.CaptureSort();Reject(()=>gateway.ApplySort(snapshot,0,mergedOptions),"vertical merges rejected before moving rows");
            Check((string)merged.Range["A2"].Value2=="客户十一处"&&(object)merged.Range["E2"].Value2==null,"rejected vertical merge leaves values and helper unchanged");
            merged.Delete();sheet.Activate();
            sheet.Name="国际十";dynamic second=com.Hold(book.Sheets.Add(Type.Missing,sheet));second.Name="国际二";dynamic third=com.Hold(book.Sheets.Add(Type.Missing,second));third.Name="业务一";third.Visible=0;
            var sheets=gateway.Capture().Sheets;gateway.SortSheets(sheets,options);Console.WriteLine("Sheet sort actual="+(string)book.Sheets[1].Name+" hidden="+(int)third.Visible+" names="+string.Join(",",gateway.Capture().Sheets.Select(x=>x.Name)));Check((string)book.Sheets[1].Name=="国际二"&&(int)third.Visible==0,"sheet order includes hidden pages without changing visibility");gateway.UndoLast();Check((string)book.Sheets[1].Name=="国际十","sheet sort undo");
            gateway.ReleaseUndoEvents();book.Close(false);
        }
        finally{if(app!=null){try{app.Quit();}catch{}}com.Dispose();}
    }
    private static void FullSampleMaskTest(string path)
    {
        byte[] before=File.ReadAllBytes(path);dynamic app=null;string output=Path.Combine(Path.GetTempPath(),"GridPilot-v15-sample-"+Guid.NewGuid().ToString("N")+".xlsx");
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            dynamic book=app.Workbooks.Open(Path.GetFullPath(path),0,true);var gateway=new ExcelGateway(()=>(object)app);
            var snapshot=gateway.CaptureMaskWorkbook();var rules=MaskInference.Suggest(snapshot,System.Threading.CancellationToken.None);
            Reject(()=>MaskPlanner.Build(snapshot,rules,System.Threading.CancellationToken.None,true),"balance unsupported volatile formulas fail closed");
            // Freeze only a disposable copy for an additional full-size export regression.
            string staticCopy=Path.Combine(Path.GetTempPath(),"GridPilot-v15-static-"+Guid.NewGuid().ToString("N")+".xlsx");book.SaveCopyAs(staticCopy);book.Close(false);
            book=app.Workbooks.Open(staticCopy,0,false);
            for(int i=1;i<=(int)book.Worksheets.Count;i++){dynamic used=book.Worksheets[i].UsedRange;used.Value2=used.Value2;}
            snapshot=gateway.CaptureMaskWorkbook();rules=MaskInference.Suggest(snapshot,System.Threading.CancellationToken.None);var result=MaskPlanner.Build(snapshot,rules,System.Threading.CancellationToken.None,true);
            gateway.ExportMask(snapshot,result,output,System.Threading.CancellationToken.None,true);
            dynamic exported=app.Workbooks.Open(output,0,true);Check((int)exported.Worksheets.Count==snapshot.Sheets.Count,"full balance mask export sheet count");
            Check(result.Changed>0&&new FileInfo(output).Length>0,"full balance mask export changed data");
            Console.WriteLine("Full sample cells="+snapshot.Cells.Count+" changed="+result.Changed+" output_bytes="+new FileInfo(output).Length);
            exported.Close(false);book.Close(false);Check(before.SequenceEqual(File.ReadAllBytes(path)),"full balance source remains byte identical");gateway.ReleaseUndoEvents();
        }
        finally{if(app!=null){try{app.Quit();}catch{}Marshal.ReleaseComObject(app);}}
    }
    private static void SortUiLaunch(string addin)
    {
        dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=true;app.DisplayAlerts=false;
        GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out int pid);
        dynamic book=app.Workbooks.Add(-4167),sheet=book.Sheets[1];sheet.Name="国际十一部";
        sheet.Range["A1:C5"].Value2=new object[,]{{"部门","金额","编号"},{"国际十部",100,"001"},{"业务一部",200,"002"},{"国际二部",300,"003"},{"国际一百零一部",400,"004"}};sheet.Range["A:C"].ColumnWidth=23;
        book.Sheets.Add(Type.Missing,sheet).Name="国际二部";sheet.Activate();sheet.Range["A2"].Select();
        string before=Newtonsoft.Json.JsonConvert.SerializeObject((object)sheet.Range["A1:C5"].Value2);
        app.RegisterXLL(Path.GetFullPath(addin));app.Run("GridPilot_ShowNavigation");app.Run("GridPilot_Focus");app.Run("GridPilot_Sort");
        Console.WriteLine("UI fixture ready");var watch=Stopwatch.StartNew();string stop=Path.Combine(Path.GetDirectoryName(Path.GetFullPath(addin)),"stop-ui");
        var exitWatch=new Stopwatch();
        try{PumpUntil(()=>File.Exists(stop),900000);Check(before==Newtonsoft.Json.JsonConvert.SerializeObject((object)sheet.Range["A1:C5"].Value2),"typing in sort editor leaves worksheet data unchanged");}
        finally{try{app.DisplayAlerts=false;book.Close(false);exitWatch.Start();app.Quit();}catch(COMException){}Marshal.ReleaseComObject((object)sheet);Marshal.ReleaseComObject((object)book);Marshal.ReleaseComObject((object)app);sheet=null;book=null;app=null;}
        for(int i=0;i<2;i++){GC.Collect();var finalizers=Task.Run(()=>GC.WaitForPendingFinalizers());PumpUntil(()=>finalizers.IsCompleted,5000);}
        try{using(var process=Process.GetProcessById(pid))PumpUntil(()=>process.HasExited,12000);}catch(ArgumentException){}
        Console.WriteLine("UI fixture full exit ms="+exitWatch.ElapsedMilliseconds);
    }
    private static void ShutdownTiming(string addin,string command=null)
    {
        dynamic app=null;int pid=0;
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=true;app.DisplayAlerts=false;GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out pid);
            dynamic books=app.Workbooks,book=books.Add(-4167);dynamic sheet=book.Sheets[1];sheet.Range["A1:A3"].Value2=new object[,]{{"部门"},{"二"},{"一"}};Marshal.ReleaseComObject(sheet);book.Saved=true;Marshal.ReleaseComObject(book);Marshal.ReleaseComObject(books);
            Check((bool)app.RegisterXLL(Path.GetFullPath(addin)),"shutdown test register XLL");app.Run("GridPilot_ShowNavigation");app.Run("GridPilot_Focus");if(command!=null)app.Run(command);
            var settle=Stopwatch.StartNew();PumpUntil(()=>settle.ElapsedMilliseconds>2200);
            var watch=Stopwatch.StartNew();app.Quit();Console.WriteLine("Quit call ms="+watch.ElapsedMilliseconds+" PID="+pid);Marshal.ReleaseComObject(app);app=null;
            for(int i=0;i<2;i++){GC.Collect();var task=Task.Run(()=>GC.WaitForPendingFinalizers());PumpUntil(()=>task.IsCompleted,5000);}
            try{using(var process=Process.GetProcessById(pid))PumpUntil(()=>process.HasExited,35000);}catch(ArgumentException){}
            Console.WriteLine("Full exit ms="+watch.ElapsedMilliseconds+" PID="+pid);
        }
        finally{if(app!=null){try{app.Quit();}catch{}Marshal.ReleaseComObject(app);}}
    }
}
