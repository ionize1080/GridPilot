using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using Newtonsoft.Json.Linq;
using SheetPilot.Core;
using SheetPilot.Excel;

internal static partial class Program
{
    private static void V11Tests(bool excel)
    {
        var source=new MaskSnapshot();source.Sheets.Add(new DataSheet{Name="Synthetic",Rows=20101,Columns=2});
        for(int i=1;i<=20101;i++)source.Cells.Add(new DataCell{Sheet=0,Row=i,Column=1,Value="old"+i});
        var output=source.Clone();foreach(var c in output.Cells)c.Value="new"+c.Row;
        var result=new MaskResult{Output=output,Changed=20101};
        var page=JObject.FromObject(MaskPreview.Page(source,result,"",19900,100,CancellationToken.None));
        Check((int)page["total"]==20000&&(bool)page["capped"]&&page["rows"].Count()==100,"20k masking preview cap and final page");
        page=JObject.FromObject(MaskPreview.Page(source,result,"old20101",0,100,CancellationToken.None));Check((int)page["total"]==1,"preview search finds changes beyond first 20k");
        var clip=JObject.FromObject(MaskPreview.Clipboard(result,0,1,20000));Check((int)clip["rows"]==20000&&!((string)clip["text"]).Contains("old"),"clipboard copies 20k masked rows only");
        Reject(()=>MaskPreview.Clipboard(result,0,1,20001),"clipboard never silently truncates >20k");
        output.Cells[0].Value="=1+2";output.Cells[1].Value="line\tbreak\nsecond";output.Cells[2].Value="<unsafe>&";
        clip=JObject.FromObject(MaskPreview.Clipboard(result,0,1,3));
        Check(((string)clip["text"]).StartsWith("'=1+2")&&((string)clip["html"]).Contains("&lt;unsafe&gt;&amp;"),"clipboard prevents formula execution and escapes HTML");
        Check(((string)clip["text"]).Contains("\"line\tbreak\nsecond\""),"clipboard preserves embedded tabs and linebreaks");
        string dir=Path.Combine(Environment.CurrentDirectory,"docs","v1.1");Directory.CreateDirectory(dir);
        string[] icons={"Show","Refresh","Database","Mask","Segment","CopyInternal","CopyValues","Rename","Focus","Text","Number","Undo","Settings"};
        using(var atlas=new Bitmap(780,240))using(var g=Graphics.FromImage(atlas))using(var font=new Font("Segoe UI",8))
        {
            g.Clear(Color.White);for(int theme=0;theme<3;theme++)for(int i=0;i<icons.Length;i++)using(var icon=RibbonIcons.Draw(icons[i],32,theme)){g.DrawImageUnscaled(icon,i*60+14,theme*80+12);g.DrawString(icons[i],font,Brushes.Black,i*60,theme*80+49);}
            atlas.Save(Path.Combine(dir,"icons.png"));
        }
        foreach(int size in new[]{16,20,24,32,48,64,96})foreach(var id in icons)using(var icon=RibbonIcons.Draw(id,size))Check(icon.Width==size&&icon.Height==size,"icon raster size "+id+" "+size);
        if(excel)V11ExcelTests();
    }
    private static void V11Launch(string addin,string command="GridPilot_Database")
    {
        dynamic app=null;Process window=null;
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.Workbooks.Add(-4167);
            Check((bool)app.RegisterXLL(addin),"packaged XLL registration");var before=Process.GetProcessesByName("electron").Select(p=>p.Id).ToArray();
            app.ActiveSheet.Range["A1"].Value2="新加坡油气有限公司";app.ActiveSheet.Range["A1"].Select();app.Run(command);
            PumpUntil(()=>{window=Process.GetProcessesByName("electron").FirstOrDefault(p=>!before.Contains(p.Id)&&p.MainWindowHandle!=IntPtr.Zero&&p.MainWindowTitle=="GridPilot");return window!=null;},20000);
            Check(window!=null,"packaged "+command+" opens production Electron window");
            window.CloseMainWindow();PumpUntil(()=>window.HasExited,10000);Check(window.HasExited,"database window closes and named pipe disconnects cleanly");
        }
        finally{if(window!=null){try{if(!window.HasExited)window.CloseMainWindow();}catch{}window.Dispose();}if(app!=null){try{app.Quit();}catch{}System.Runtime.InteropServices.Marshal.ReleaseComObject(app);}}
    }
    private static void V11ExcelTests()
    {
        dynamic app=null;string dir=Path.Combine(Path.GetTempPath(),"GridPilot-v11-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            var gateway=new ExcelGateway(()=>(object)app);dynamic book=app.Workbooks.Add(-4167),first=book.Sheets[1];first.Name="Data";
            first.Range["A1"].Value2="Synthetic report";first.Range["A2:D5"].NumberFormat="@";
            first.Range["A2:D5"].Value2=new object[,]{{"月份","贷款号","余额","备注"},{"202402","001","123.45","synthetic"},{"202403","002","0.01","line\ntext"},{"202403","","10.00","missing"}};
            first.Range["E2"].Value2="金额公式";first.Range["E3"].Formula="=123.45";app.CalculateFull();
            string file=Path.Combine(dir,"synthetic.xlsx");book.SaveAs(file,51);
            var parsed=JObject.FromObject(ReportReader.Read(file,"Data",0,CancellationToken.None));
            Check(parsed["rows"].Count()==5&&(string)parsed["rows"][2]["values"][1]=="001","XLSX reader preserves title/header and identifier zeros");
            Check((string)parsed["rows"][2]["values"][4]=="123.45","XLSX reader uses saved formula cache without calculation");
            var snap=gateway.Capture();var a=snap.Sheets[0];gateway.EditSheet(new SheetEdit{Source=a,Action="rename",Name="Renamed"});
            Check((string)first.Name=="Renamed","navigation rename actual Excel sheet");
            gateway.EditSheet(new SheetEdit{Source=a,Action="add"});Check((int)book.Sheets.Count==2,"navigation add after selected sheet");
            snap=gateway.Capture();var second=snap.Sheets.First(s=>s.Key!=a.Key);
            gateway.EditSheet(new SheetEdit{Source=second,Target=a,Action="move",After=false});Check((string)book.Sheets[1].Name==second.Name,"drag move before target");
            gateway.EditSheet(new SheetEdit{Source=second,Target=a,Action="move",After=true});Check((string)book.Sheets[2].Name==second.Name,"drag move after target");
            gateway.EditSheet(new SheetEdit{Source=a,Action="hide"});Check((int)first.Visible==0,"navigation hide");
            Reject(()=>gateway.EditSheet(new SheetEdit{Source=second,Action="hide"}),"last visible sheet protected");gateway.EditSheet(new SheetEdit{Source=a,Action="show"});
            Reject(()=>gateway.EditSheet(new SheetEdit{Source=a,Action="rename",Name="bad/name"}),"illegal sheet name rejected");
            Reject(()=>gateway.EditSheet(new SheetEdit{Source=a,Action="rename",Name=second.Name}),"duplicate sheet name rejected");
            book.Protect(Type.Missing,true);Reject(()=>gateway.EditSheet(new SheetEdit{Source=a,Action="add"}),"protected workbook structure checked");book.Unprotect();
            gateway.EditSheet(new SheetEdit{Source=second,Action="first"});Check((string)book.Sheets[1].Name==second.Name,"move to first");gateway.EditSheet(new SheetEdit{Source=second,Action="last"});Check((string)book.Sheets[2].Name==second.Name,"move to last");
            book.Close(false);
            gateway.OpenReportSource(file,"Data",4);Check((int)app.Selection.Row==4&&(string)app.ActiveSheet.Name=="Data","source action opens exact original sheet and row");app.ActiveWorkbook.Close(false);
            gateway.ReleaseUndoEvents();
        }
        finally{if(app!=null){try{app.Quit();}catch{}System.Runtime.InteropServices.Marshal.ReleaseComObject(app);}Directory.Delete(dir,true);}
    }
}
