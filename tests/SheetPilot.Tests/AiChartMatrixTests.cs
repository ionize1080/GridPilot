using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using SheetPilot.Excel;

internal static partial class Program
{
    private static void AiChartMatrixTests()
    {
        dynamic app=null,book=null;
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            book=app.Workbooks.Add(-4167);dynamic sheet=book.Worksheets[1];sheet.Name="Charts";
            using(var ai=new AiOperations(()=>(object)app))
            {
                string snapshot=(string)JObject.FromObject(ai.Capture("workbook"))["snapshotId"];ai.Approval(true,"workbook");
                foreach(var type in AiOperations.Excel2010Charts.Where(k=>k.Key.StartsWith("xl")).GroupBy(k=>k.Value).Select(g=>g.First()))
                {
                    sheet.Cells.Clear();string[] headers=type.Value==88?new[]{"Date","High","Low","Close"}:type.Value==89?new[]{"Date","Open","High","Low","Close"}:type.Value==90?new[]{"Date","Volume","High","Low","Close"}:type.Value==91?new[]{"Date","Volume","Open","High","Low","Close"}:new[]{"Date","Value","Other"};
                    var data=new object[6,headers.Length];for(int c=0;c<headers.Length;c++)data[0,c]=headers[c];for(int r=1;r<6;r++)for(int c=0;c<headers.Length;c++) data[r,c]=headers[c]=="Date"?45000.0+r:headers[c]=="Volume"?1000.0+r*100:headers[c]=="Open"?10.0+r:headers[c]=="High"?15.0+r:headers[c]=="Low"?8.0+r:headers[c]=="Close"?12.0+r:20.0+r*c;
                    string address="A1:"+(char)('A'+headers.Length-1)+"6";sheet.Range[address].Value2=data;sheet.Range["A2:A6"].NumberFormat="yyyy-mm-dd";
                    var create=JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="chart.create",["options"]=new JObject{["sheet"]="Charts",["address"]=address,["type"]=type.Key,["name"]="TestChart"}}));
                    Check((bool)create["applied"]&&(int)sheet.ChartObjects("TestChart").Chart.ChartType==type.Value,"native "+type.Key);
                    ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="chart.delete",["options"]=new JObject{["sheet"]="Charts",["chart"]="TestChart"}});
                }
            }
        }
        finally{if(book!=null){try{book.Close(false);}catch{}Marshal.ReleaseComObject((object)book);}if(app!=null){try{app.Quit();}catch{}Marshal.ReleaseComObject((object)app);}}
    }
    private static void AiVbaAndFormatsTests()
    {
        dynamic app=null;
        string folder=Path.Combine(Path.GetTempPath(),"gp-formats-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            foreach(var format in new[]{Tuple.Create(51,".xlsx"),Tuple.Create(52,".xlsm"),Tuple.Create(50,".xlsb"),Tuple.Create(56,".xls")})
            {
                object book=app.Workbooks.Add(-4167);dynamic b=book;
                try
                {
                    b.Worksheets[1].Name="Data";b.Worksheets[1].Range["A1"].Value2="original";b.SaveAs(Path.Combine(folder,"format"+format.Item2),format.Item1);
                    using(var ai=new AiOperations(()=>(object)app))
                    {
                        string snapshot=(string)JObject.FromObject(ai.Capture("workbook"))["snapshotId"];
                        var p=JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="range.write",["options"]=new JObject{["sheet"]="Data",["address"]="A1",["data"]=new JArray{new JArray{"changed"}}}}));
                        System.Threading.Thread.Sleep(1100); // Cross save-timestamp boundaries in every format.
                        ai.ApplyBatch(new[]{(string)p["proposalId"]});Check((string)b.Worksheets[1].Range["A1"].Value2=="changed","manual edit and snapshots "+format.Item2);
                        var conflict=JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="range.write",["options"]=new JObject{["sheet"]="Data",["address"]="A2",["data"]=new JArray{new JArray{"must not apply"}}}}));
                        b.Worksheets[1].Range["A3"].Value2="external edit";
                        bool rejected=false;try{ai.ApplyBatch(new[]{(string)conflict["proposalId"]});}catch(InvalidOperationException ex){rejected=ex.Message.Contains("已变化");}
                        Check(rejected&&(object)b.Worksheets[1].Range["A2"].Value2==null,"real content conflict remains blocked "+format.Item2);
                        if(format.Item1!=51)
                        {
                            try { ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="vba.read",["options"]=new JObject()}); }
                            catch(InvalidOperationException ex){if(!ex.Message.Contains("VBA"))throw;Console.WriteLine("SKIP real VBA read/write: Excel trust access disabled (not changed)");continue;}
                            ai.Approval(true,"workbook");
                            var code="Public Function GPTest(ByVal x As Double) As Double\n GPTest=x*2\nEnd Function";
                            var v=JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="vba.write",["options"]=new JObject{["module"]="GPTests",["code"]=code}}));Check((bool)v["applied"],"benign VBA auto writes when trusted");
                            var read=JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="vba.read",["options"]=new JObject{["module"]="GPTests"}}));Check(((string)read["modules"][0]["code"]).Contains("GPTest"),"read native VBA module");
                            var risky=JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="vba.write",["options"]=new JObject{["module"]="GPReview",["code"]="Sub ReviewOnly()\n Shell \"cmd /c echo test\"\nEnd Sub"}}));
                            Check(!(bool)risky["applied"]&&(bool)risky["requiresManualApproval"],"real VBA risk cannot auto apply");
                            bool blocked=false;try{ai.ApplyBatch(new[]{(string)risky["proposalId"]},false,false);}catch(InvalidOperationException){blocked=true;}Check(blocked,"host rejects non-manual risky VBA approval");
                            ai.ApplyBatch(new[]{(string)risky["proposalId"]});
                            var reviewed=JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="vba.read",["options"]=new JObject{["module"]="GPReview"}}));Check(((string)reviewed["modules"][0]["code"]).Contains("Shell"),"manual approval writes reviewed VBA without running it");
                            ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="vba.delete",["options"]=new JObject{["module"]="GPReview"}});
                            var remaining=JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=snapshot,["operation"]="vba.read",["options"]=new JObject()}));Check(!remaining["modules"].Any(m=>(string)m["module"]=="GPReview"),"VBA module delete");
                        }
                    }
                }
                finally{b.Close(false);Marshal.ReleaseComObject(book);}
            }
        }
        finally{if(app!=null){try{app.Quit();}catch{}Marshal.ReleaseComObject((object)app);}}
    }
}
