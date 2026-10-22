using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using SheetPilot.Core;
using SheetPilot.Excel;

internal static partial class Program
{
    private static void AiV21Tests(bool excel)
    {
        var prefs=new ToolPreferences();Check(prefs.SnapshotBudgetMB==500,"500 MB default");prefs.SnapshotBudgetMB=1;prefs.Normalize();Check(prefs.SnapshotBudgetMB==100,"100 MB minimum");
        string temp=Path.Combine(Path.GetTempPath(),"gp-v21-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(temp);
        var store=new WorkbookSnapshotStore(Path.Combine(temp,"history"));string source=Path.Combine(temp,"random.xls");var bytes=new byte[60000];new Random(42).NextBytes(bytes);File.WriteAllBytes(source,bytes);
        string first=store.Save(source,"book","first"),second=store.Save(source,"book","second");long two=store.UsedBytes;store.BudgetBytes=two-1;store.Prune();Check(store.List().Count==1&&(string)store.List()[0]["id"]==second,"FIFO prunes manifest while keeping shared chunks");
        Check(File.ReadAllBytes(store.Materialize(second,temp)).SequenceEqual(bytes),"shared chunks survive pruning");
        byte[] unique=(byte[])bytes.Clone();unique[0]^=1;File.WriteAllBytes(source,unique);store.BudgetBytes=100000;string third=store.Save(source,"book","third");Check(store.List().Count==1&&(string)store.List()[0]["id"]==third&&store.UsedBytes<=100000,"unreferenced unique blobs removed");
        string orphan=Path.Combine(store.Root,"blobs",new string('f',64)+".gz");File.WriteAllBytes(orphan,new byte[10]);store.Prune();Check(!File.Exists(orphan),"orphan reclamation");
        store.BudgetBytes=100;bool rejected=false;try{store.Save(source,"book","oversize");}catch(IOException){rejected=true;}Check(rejected&&store.UsedBytes<=100,"oversize rejected and budget enforced");
        store.BudgetBytes=100000;store.Save(source,"book","again");store.BudgetBytes=0;store.Prune();Check(store.UsedBytes==0&&store.List().Count==0,"lower budget cleans all obsolete data");
        string settings=Path.Combine(temp,"prefs.json");File.WriteAllText(settings,"{\"SchemaVersion\":1}");var migrated=new ToolPreferencesStore(settings).Load();Check(migrated.SnapshotBudgetMB==500&&migrated.AiShowExplain&&migrated.AiShowManualFormula,"old settings migration defaults");
        if(!excel)return;
        dynamic app=null,one=null,twoBook=null;
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            one=app.Workbooks.Add(-4167);dynamic a=one.Worksheets[1];a.Name="Sales";a.Range["A1:B4"].Value2=new object[,]{{"Item","Amount"},{"A",10},{"B",20},{"C",30}};
            dynamic hidden=one.Worksheets.Add();hidden.Name="Hidden";hidden.Range["A1"].Value2=99;hidden.Visible=2;a.Activate();
            twoBook=app.Workbooks.Add(-4167);dynamic b=twoBook.Worksheets[1];b.Name="Sales";b.Range["A1:B3"].Value2=new object[,]{{"Item","Amount"},{"D",40},{"E",50}};one.Activate();
            using(var ai=new AiOperations(()=>(object)app,Path.Combine(temp,"excel-history")))
            {
                var root=JObject.FromObject(ai.Capture("auto"));string id=(string)root["snapshotId"];
                Check((bool)root["metadataOnly"]&&root["catalog"].Count()==2,"auto scans metadata only for both books");
                var firstBook=root["catalog"].First(v=>(string)v["book"]==(string)one.Name);Check(firstBook["sheets"].Any(v=>(string)v["sheet"]=="Hidden"),"directory includes very hidden sheet");
                Func<string,JObject,JObject> discovery=(op,o)=>JObject.FromObject(ai.ExtendedAsync(new JObject{["snapshotId"]=id,["operation"]=op,["options"]=o}).GetAwaiter().GetResult());
                var cap=discovery("context.capture",new JObject{["bookId"]=firstBook["bookId"]});string sid=(string)cap["snapshotId"];
                Check((double)cap["statistics"]["sum"]==159,"full capture includes hidden value");
                var other=root["catalog"].First(v=>(string)v["book"]==(string)twoBook.Name);
                var cap2=discovery("context.capture",new JObject{["bookId"]=other["bookId"],["sheet"]="Sales"});Check((double)cap2["statistics"]["sum"]==90,"same sheet names in separate books stay distinct");
                a.Range["B2"].Value2=1000;var frozen=JObject.FromObject(ai.Read(new JObject{["snapshotId"]=sid,["operation"]="statistics"}));Check((double)frozen["statistics"]["sum"]==159,"frozen read remains stable after editing");
                bool refused=false;try{ai.Read(new JObject{["snapshotId"]=id,["operation"]="statistics"});}catch(InvalidOperationException){refused=true;}Check(refused,"metadata cannot masquerade as zero data");
                Func<string,JObject,JObject> run=(op,o)=>JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=sid,["operation"]=op,["options"]=o}));
                refused=false;try{run("workbook.list",new JObject{["directory"]=temp});}catch(InvalidOperationException){refused=true;}Check(refused,"auto mode cannot enumerate disk workbooks");
                one.Activate();ai.Approval(true,"workbook");var staged=JObject.FromObject(ai.Extended(new JObject{["snapshotId"]=cap2["snapshotId"],["operation"]="range.write",["options"]=JObject.Parse("{sheet:'Sales',address:'D1',data:[[12]]}")}));Check(!(bool)staged["applied"]&&b.Range["D1"].Value2==null,"auto approval never crosses workbook identity");
                var features=run("capability.list",new JObject());Console.WriteLine("HOST "+features["excelVersion"]+" build "+features["build"]);
                if((bool)features["features"]["dynamicArrays"]){run("range.formula2",JObject.Parse("{sheet:'Sales',address:'G1',formula:'=SEQUENCE(3)'}"));Check((double)a.Range["G3"].Value2==3,"Formula2 spills three rows");var spill=run("range.spill",JObject.Parse("{sheet:'Sales',address:'G1'}"));Check((string)spill["address"]=="G1:G3","spill inspection");}
                if((bool)features["features"]["queries"]){run("query.set",JObject.Parse("{name:'Synthetic',formula:'let Source = #table({\"A\"}, {{1}}) in Source'}"));Check(run("query.list",new JObject())["items"].Count()==1,"query creation and read without refresh");run("query.delete",JObject.Parse("{name:'Synthetic'}"));Check(run("query.list",new JObject())["total"].Value<int>()==0,"query deletion");}
                if((bool)features["features"]["threadedComments"]){run("comment.add",JObject.Parse("{sheet:'Sales',address:'A1',text:'Synthetic test'}"));run("comment.reply",JObject.Parse("{sheet:'Sales',address:'A1',text:'Reply'}"));Check(run("comment.list",JObject.Parse("{sheet:'Sales'}"))["items"][0]["replyCount"].Value<int>()==1,"threaded comments and replies");run("comment.delete",JObject.Parse("{sheet:'Sales',address:'A1'}"));}
                run("connection.list",new JObject());
                if((bool)features["features"]["dataModel"])run("model.inspect",new JObject());
                run("table.create",JObject.Parse("{sheet:'Sales',name:'TestTable',address:'A1:B4'}"));run("table.resize",JObject.Parse("{sheet:'Sales',name:'TestTable',address:'A1:B5'}"));Check((int)a.ListObjects["TestTable"].Range.Rows.Count==5,"table resize preserves native table");
                run("sheet.create",JObject.Parse("{name:'Summary'}"));one.Activate();
                run("pivot.create",JObject.Parse("{sheet:'Sales',address:'A1:B4',destinationSheet:'Summary',destination:'A1',name:'V21Pivot',rows:['Item'],values:[{field:'Amount',function:'sum',caption:'Total'}]}"));one.Activate();
                run("slicer.add",JObject.Parse("{sheet:'Summary',pivot:'V21Pivot',field:'Item',name:'V21Slicer'}"));
                var caches=run("slicer.list",new JObject());string cache=(string)caches["items"][0]["name"];
                run("slicer.filter",new JObject{["cache"]=cache,["items"]=new JArray{"B"}});Check((double)one.Worksheets["Summary"].PivotTables("V21Pivot").GetPivotData("Total").Value2==20,"slicer filter changes pivot result");
                run("slicer.clear",new JObject{["cache"]=cache});Check((double)one.Worksheets["Summary"].PivotTables("V21Pivot").GetPivotData("Total").Value2==1050,"slicer clear restores full source total");
                a.Activate();
                if((bool)features["features"]["modernCharts"])foreach(string type in new[]{"xlWaterfall","xlTreemap","xlHistogram","xlSunburst","xlBoxwhisker","xlPareto","xlFunnel"}){run("chart.create",new JObject{["sheet"]="Sales",["address"]="A1:B4",["type"]=type,["name"]=type});Console.WriteLine("PASS modern chart "+type);}
                twoBook.Close(false);refused=false;try{ai.Read(new JObject{["snapshotId"]=cap2["snapshotId"],["operation"]="statistics"});}catch(InvalidOperationException){refused=true;}Check(refused,"closed workbook context rejected");
                var all=JObject.FromObject(ai.Capture("all"));Check((string)all["scope"]=="all"&&all["catalog"].Count()==1,"all mode refresh excludes closed book");
            }
        }
        finally{if((object)twoBook!=null){try{twoBook.Close(false);}catch{}try{Marshal.ReleaseComObject((object)twoBook);}catch{}}if((object)one!=null){try{one.Close(false);}catch{}try{Marshal.ReleaseComObject((object)one);}catch{}}if((object)app!=null){try{app.Quit();}catch{}try{Marshal.ReleaseComObject((object)app);}catch{}}}
    }
}
