using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using SheetPilot.Excel;

internal static partial class Program
{
    private static void AiWorkspacePaperTests()
    {
        string folder=Path.Combine(Path.GetTempPath(),"gp-workspace-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        dynamic app=null,original=null,paper=null,opened=null;
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            original=app.Workbooks.Add(-4167);original.Worksheets[1].Range["A1"].Value2="Unsaved original";
            paper=app.Workbooks.Add(-4167);paper.Worksheets[1].Range["A1:B2"].Value2=new object[,]{{"Value","Total"},{7,12}};
            string file=Path.Combine(folder,"分析底稿.xlsx");paper.SaveAs(file,51);paper.Close(false);byte[] before=File.ReadAllBytes(file);original.Activate();
            using(var ai=new AiOperations(()=>(object)app,Path.Combine(folder,"history")))
            {
                int security=app.AutomationSecurity;bool events=app.EnableEvents;
                ai.Approval(true,"workbook");
                var result=JObject.FromObject(ai.OpenWorkspacePaper(file));opened=app.ActiveWorkbook;
                Check((bool)result["opened"]&&(int)app.Workbooks.Count==2,"workspace paper opens in the same Excel instance");
                Check((string)opened.Name=="分析底稿.xlsx"&&(double)opened.Worksheets[1].Range["B2"].Value2==12,"opened generated workbook retains computed values");
                Check((string)original.Worksheets[1].Range["A1"].Value2=="Unsaved original"&&!(bool)original.Saved,"source workbook and unsaved edits are preserved");
                Check((int)app.AutomationSecurity==security&&(bool)app.EnableEvents==events,"opening restores host security and event settings");
                Check(!(bool)JObject.FromObject(ai.CheckpointState())["automatic"],"existing workbook approval never transfers to a paper");
                opened.Close(false);
                Check(before.SequenceEqual(File.ReadAllBytes(file)),"opening never rewrites the saved paper");
                bool refused=false;try{ai.OpenWorkspacePaper(Path.Combine(folder,"missing.xlsx"));}catch(InvalidOperationException){refused=true;}Check(refused,"missing paper rejected");
                refused=false;try{ai.OpenWorkspacePaper(Path.Combine(folder,"unsafe.xlsm"));}catch(InvalidOperationException){refused=true;}Check(refused,"UI opener accepts generated XLSX only");
            }
        }
        finally
        {
            foreach(object book in new object[]{opened,paper,original})if(book!=null){try{((dynamic)book).Close(false);}catch{}try{Marshal.ReleaseComObject(book);}catch{}}
            if((object)app!=null){try{app.Quit();}finally{Marshal.ReleaseComObject((object)app);}}
        }
    }
}
