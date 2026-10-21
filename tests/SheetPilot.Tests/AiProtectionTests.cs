using System;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;
using SheetPilot.Excel;

internal static partial class Program
{
    private static void AiProtectionTests(dynamic app)
    {
        dynamic book=app.Workbooks.Add(-4167),sheet=book.Worksheets[1];sheet.Name="测试";
        sheet.Range["A1:D5"].Value2=new object[,]{{"部门","季度","金额","数量"},{"甲","Q1",10.0,1.0},{"甲","Q2",20.0,2.0},{"乙","Q1",30.0,3.0},{"乙","Q2",40.0,4.0}};
        try
        {
            using(var ai=new AiOperations(()=>(object)app))
            {
                string id=(string)JObject.FromObject(ai.Capture("sheet"))["snapshotId"];
                Func<string,string,JObject,JObject> propose=(kind,address,extra)=>{extra["snapshotId"]=id;extra["sheet"]="测试";extra["address"]=address;extra["kind"]=kind;return JObject.FromObject(ai.Propose(extra));};
                Check(!(bool)JObject.FromObject(ai.CheckpointState())["automatic"],"automatic approval defaults off");
                ai.Approval(true);
                var first=propose("values","F2",new JObject{["data"]=new JArray{new JArray{10}}});
                Check((bool)first["applied"]&&(double)sheet.Range["F2"].Value2==10,"explicit opt-in automatically applies");
                string folder=(string)first["checkpoint"]["folder"];
                Check(Directory.GetFiles(folder,"before-*.xlsx").Length==1&&Directory.GetFiles(folder,"after-*.xlsx").Length==1,"complete before/after workbook files exist");
                propose("formula","F2",new JObject{["formula"]="=SUM(C2:C5)"});
                propose("values","G2:H2",new JObject{["data"]=new JArray{new JArray{"=literal",true}}});
                ai.RestoreCheckpoint(true);Check(sheet.Range["F2"].Value2==null&&sheet.Range["G2"].Value2==null,"auto checkpoint coalesces repeated edits to original state");
                Check(!(bool)JObject.FromObject(ai.CheckpointState())["automatic"],"restore disables auto approval");
                ai.RestoreCheckpoint(false);Check((double)sheet.Range["F2"].Value2==100&&(string)sheet.Range["G2"].Value2=="=literal"&&!(bool)sheet.Range["G2"].HasFormula,"after snapshot restores formula and literal text");
                for(int i=0;i<10;i++)ai.Capture("sheet");ai.RestoreCheckpoint(true);ai.RestoreCheckpoint(false);Check((double)sheet.Range["F2"].Value2==100,"restore survives context eviction");
                sheet.Range["F2"].Value2=999;bool conflict=false;try{ai.RestoreCheckpoint(true);}catch(InvalidOperationException){conflict=true;}
                Check(conflict&&(double)sheet.Range["F2"].Value2==999,"manual edit blocks entire restore");sheet.Range["F2"].Formula="=SUM(C2:C5)";
                id=(string)JObject.FromObject(ai.Capture("sheet"))["snapshotId"];
                var style=propose("format","J2:J3",new JObject{["style"]=new JObject{["bold"]=true,["italic"]=true,["font_size"]=14,["font_color"]="#165C49",["fill_color"]="#D9EFE5",["number_format"]="0.00",["horizontal_alignment"]="center",["wrap_text"]=true}});
                ai.Apply((string)style["proposalId"]);Check((bool)sheet.Range["J2"].Font.Italic&&(double)sheet.Range["J2"].Font.Size==14&&(bool)sheet.Range["J2"].WrapText,"all format properties apply");ai.RestoreCheckpoint(true);ai.RestoreCheckpoint(false);Check((bool)sheet.Range["J2"].Font.Bold,"format redo");
                foreach(string kind in new[]{"column","bar","line","pie","scatter","area"})
                {
                    var chart=propose("chart","B1:D5",new JObject{["chartType"]=kind,["title"]="测试 "+kind});ai.Apply((string)chart["proposalId"]);
                    Check((int)sheet.ChartObjects().Count==1,"chart type "+kind);ai.RestoreCheckpoint(true);ai.RestoreCheckpoint(false);Check((int)sheet.ChartObjects().Count==1,"chart redo "+kind);ai.RestoreCheckpoint(true);
                }
                foreach(string fn in new[]{"sum","count","average","min","max"})
                {
                    string name="透视_"+fn;
                    var pivot=propose("pivot","A1:D5",new JObject{["destinationSheet"]=name,["rows"]=new JArray("部门"),["columns"]=new JArray("季度"),["values"]=new JArray{new JObject{["field"]="金额",["function"]=fn}}});
                    ai.Apply((string)pivot["proposalId"]);dynamic ps=book.Worksheets[name];Check((int)ps.PivotTables().Count==1,"native pivot aggregation "+fn);
                    ai.RestoreCheckpoint(true);Check((int)book.Worksheets.Count==1,"pivot before removes generated sheet "+fn);
                    ai.RestoreCheckpoint(false);Check((int)book.Worksheets[name].PivotTables().Count==1,"pivot redo "+fn);
                    book.Worksheets[name].Range["A1"].Value2="manual";conflict=false;try{ai.RestoreCheckpoint(true);}catch(InvalidOperationException){conflict=true;}Check(conflict,"pivot manual edits block sheet removal "+fn);
                    book.Worksheets[name].Range["A1"].Clear();ai.RestoreCheckpoint(true);sheet.Activate();
                }
                sheet.Activate();id=(string)JObject.FromObject(ai.Capture("sheet"))["snapshotId"];
                ai.Approval(true);
                propose("values","C2",new JObject{["data"]=new JArray{new JArray{11}}});
                propose("chart","A1:D5",new JObject{["chartType"]="column",["title"]="source-change"});
                propose("values","C2",new JObject{["data"]=new JArray{new JArray{12}}});
                ai.RestoreCheckpoint(true);Check((double)sheet.Range["C2"].Value2==10&&(int)sheet.ChartObjects().Count==0,"restore chart and repeatedly edited source together");
                ai.RestoreCheckpoint(false);Check((double)sheet.Range["C2"].Value2==12&&(int)sheet.ChartObjects().Count==1,"redo restores source before chart");
                ai.RestoreCheckpoint(true);
                long reported=0;ai.Progress=(phase,percent,current,total)=>reported=Math.Max(reported,current);
                sheet.Range["L1:L20000"].Value2=1;
                ai.YieldCapture=()=>{ai.CancelCapture=true;return System.Threading.Tasks.Task.CompletedTask;};
                bool cancelled=false;try{ai.Capture("sheet");}catch(OperationCanceledException){cancelled=true;}
                Check(cancelled,"large capture can cancel before any model send");ai.YieldCapture=null;ai.Capture("sheet");Check(reported>=20000,"capture reports actual cell progress after cancellation recovery");
                var invalid=false;try{propose("pivot","A1:D5",new JObject{["destinationSheet"]="测试",["rows"]=new JArray("部门"),["values"]=new JArray{new JObject{["field"]="金额",["function"]="sum"}}});}catch(InvalidOperationException){invalid=true;}Check(invalid,"pivot refuses existing worksheet");
            }
        }
        finally{book.Close(false);}
    }
}
