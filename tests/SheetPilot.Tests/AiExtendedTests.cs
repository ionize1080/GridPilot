using System;
using System.IO;
using System.Linq;
using System.IO.Compression;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using SheetPilot.Excel;

internal static partial class Program
{
    private static void AiExtendedTests(bool excel)
    {
        Check(VbaRisk.Inspect("Sub A()\nRange(\"A1\").Value = 3\nEnd Sub").Length == 0, "ordinary VBA requires no extra approval");
        Check(VbaRisk.Inspect("' Shell \"cmd\"\nRem WinHttp\nSub A()\nEnd Sub").Length == 0, "VBA comments ignored");
        foreach (string code in new[] { "Shell \"cmd\"", "CreateObject(\"WScript.Shell\").Run \"cmd\"", "CreateObject(\"MSXML2.XMLHTTP\")", "CreateObject(\"WScript.\" & \"Shell\")", "CreateObject(target)", "CallByName obj, name, 1", "Declare Function X Lib \"urlmon\" Alias \"URLDownloadToFileA\" () As Long", "CreateObject _\n (\"WinHttp.WinHttpRequest.5.1\")" }) Check(VbaRisk.Inspect(code).Length > 0, "VBA execution/network detection");
        Check(VbaRisk.Inspect("Open \"C:\\data.txt\" For Output As #1\nPrint #1, \"a\"\nClose #1").Length == 0, "local file VBA not added to risk list");
        Check(VbaRisk.Inspect("Range(\"A1\").Value=\"Shell\"").Length==0,"ordinary shell word in string is not process execution");
        Check(AiOperations.Excel2010Charts.Values.Distinct().Count() == 73, "all 73 Excel 2010 chart types");
        string temp = Path.Combine(Path.GetTempPath(), "gp-snapshot-test-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(temp);
        var store = new WorkbookSnapshotStore(Path.Combine(temp, "store")); string source = Path.Combine(temp, "source.xlsx");
        using (var zip = ZipFile.Open(source, ZipArchiveMode.Create)) { using (var writer = new StreamWriter(zip.CreateEntry("xl/worksheets/sheet1.xml").Open())) writer.Write("<sheet>" + new string('x', 600000) + "</sheet>"); using (var writer = new StreamWriter(zip.CreateEntry("xl/vbaProject.bin").Open())) writer.Write("synthetic-vba"); }
        string first = store.Save(source, "test", "before"), second = store.Save(source, "test", "after"); int blobs = Directory.GetFiles(Path.Combine(store.Root, "blobs")).Length;
        Check(blobs < 6 && store.List().Count == 2, "identical versions share immutable compressed chunks");
        string restored = store.Materialize(first, Path.Combine(temp, "restored")); using (var a = ZipFile.OpenRead(source)) using (var b = ZipFile.OpenRead(restored)) foreach (var entry in a.Entries) using (var ar = new StreamReader(entry.Open())) using (var br = new StreamReader(b.GetEntry(entry.FullName).Open())) Check(ar.ReadToEnd() == br.ReadToEnd(), "snapshot restores exact ZIP part contents");
        string blob = Directory.GetFiles(Path.Combine(store.Root, "blobs"))[0]; File.WriteAllBytes(blob, new byte[] { 0, 1, 2 }); bool rejected = false; try { store.Materialize(second, temp); } catch { rejected = true; } Check(rejected, "corrupt snapshot cannot be restored");
        if (!excel) return;
        dynamic app = null, book = null;
        try
        {
            app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true)); app.Visible = false; app.DisplayAlerts = false; app.AutomationSecurity = 3;
            book = app.Workbooks.Add(-4167); dynamic sheet = book.Worksheets[1]; sheet.Name = "Sales";
            sheet.Range["A1:D5"].Value2 = new object[,] { { "Region", "Revenue", "Cost", "Year" }, { "East", 100.0, 40.0, 2025 }, { "West", 200.0, 70.0, 2026 }, { "East", 150.0, 60.0, 2026 }, { "North", 80.0, 30.0, 2025 } };
            using (var ai = new AiOperations(() => (object)app))
            {
                string snapshot = (string)JObject.FromObject(ai.Capture("workbook"))["snapshotId"];
                Func<string, JObject, JObject> call = (operation, options) => { var result = JObject.FromObject(ai.Extended(new JObject { ["snapshotId"] = snapshot, ["operation"] = operation, ["options"] = options })); Console.WriteLine("OP " + operation); return result; };
                Func<string, string, JObject> run = (op, json) => call(op, JObject.Parse(json));
                var pending = run("range.write", "{sheet:'Sales',address:'F1',data:[[99]]}"); Check(!(bool)pending["applied"] && sheet.Range["F1"].Value2 == null, "extended mutation stages by default"); ai.ApplyBatch(new[] { (string)pending["proposalId"] }); Check((double)sheet.Range["F1"].Value2 == 99, "manual extended apply writes");
                var stale=run("range.write", "{sheet:'Sales',address:'F5',data:[[8]]}");sheet.Range["F6"].Value2=7;bool staleRejected=false;try{ai.ApplyBatch(new[]{(string)stale["proposalId"]});}catch(InvalidOperationException){staleRejected=true;}Check(staleRejected&&sheet.Range["F5"].Value2==null,"manual advanced proposal refuses intervening workbook edits");ai.Discard(new[]{(string)stale["proposalId"]});
                var basic=JObject.FromObject(ai.Propose(new JObject{["snapshotId"]=snapshot,["kind"]="values",["sheet"]="Sales",["address"]="F8",["data"]=new JArray{new JArray{12}}}));
                var advanced=run("range.write","{sheet:'Sales',address:'F9',data:[[13]]}");ai.ApplyBatch(new[]{(string)basic["proposalId"],(string)advanced["proposalId"]});Check((double)sheet.Range["F8"].Value2==12&&(double)sheet.Range["F9"].Value2==13,"mixed basic and advanced batch uses one preflight");
                ai.Approval(true, "workbook");
                run("sheet.create", "{name:'Summary'}"); book.Activate();
                run("sheet.create", "{name:'Scratch'}"); book.Activate();
                run("plugin.rename", "{sheets:['Scratch'],settings:{Prefix:'Tmp_',Numbering:false}}");run("sheet.rename", "{sheet:'Tmp_Scratch',name:'Scratch'}");
                run("sheet.batch_rename", "{changes:[{from:'Summary',to:'Scratch'},{from:'Scratch',to:'Summary'}]}"); Check((string)book.Worksheets[2].Name == "Scratch", "batch rename swaps names");
                run("range.write", "{sheet:'Sales',address:'F2:F3',data:[['=literal'],['=1+1']]}"); Check(!(bool)sheet.Range["F3"].HasFormula, "literal writes do not execute formula text");
                run("range.write", "{sheet:'Sales',address:'G2:G3',data:[['=B2-C2'],['=B3-C3']],formulas:true}"); Check((double)sheet.Range["G3"].Value2 == 130, "advanced formula write");
                Check(run("range.read", "{sheet:'Sales',address:'G2:G3',limit:1}")["nextOffset"].Value<int>() == 1, "live range pagination");
                run("range.format", "{sheet:'Sales',address:'A1:D1',font:{Bold:true,Size:13},fill:{Color:123456},borders:[{index:9,properties:{LineStyle:1,Weight:2}}]}"); Check((bool)sheet.Range["A1"].Font.Bold, "rich range formatting");
                run("range.validation", "{sheet:'Sales',address:'H2:H5',type:3,formula1:'A,B,C'}"); Check((int)sheet.Range["H2"].Validation.Type == 3, "native validation");
                run("range.conditional_format", "{sheet:'Sales',address:'B2:B5',kind:'dataBar',barColor:'#117A55'}"); Check((int)sheet.Range["B2:B5"].FormatConditions.Count == 1, "native conditional data bar");
                run("range.copy", "{sheet:'Sales',address:'A1:D5',destination:'J1',pasteType:-4163}"); Check((string)sheet.Range["J2"].Value2 == "East", "copy values");
                run("range.sort", "{sheet:'Sales',address:'J1:M5',header:true,levels:[{column:2,mode:'values',descending:true}]}"); Check((double)sheet.Range["K2"].Value2 == 200, "existing multi-sort engine through MCP");
                run("range.replace", "{sheet:'Sales',address:'J2:J5',find:'West',replacement:'West2'}"); Check((string)sheet.Range["J2"].Value2 == "West2", "range replacement");
                run("range.merge", "{sheet:'Sales',address:'O1:P1'}"); run("range.merge", "{sheet:'Sales',address:'O1:P1',merge:false}"); Check(!(bool)sheet.Range["O1:P1"].MergeCells, "merge and unmerge");
                run("range.insert", "{sheet:'Scratch',address:'A1:A1',axis:'rows'}"); run("range.delete", "{sheet:'Scratch',address:'A1:A1',axis:'rows'}");
                run("table.create", "{sheet:'Sales',name:'SalesTable',address:'J1:M5',properties:{ShowTotals:true}}"); run("table.update", "{sheet:'Sales',name:'SalesTable',column:2,totals:1}"); Check((int)sheet.ListObjects.Count == 1, "table creation and totals");
                run("table.delete", "{sheet:'Sales',name:'SalesTable',keepData:true}"); Check((string)sheet.Range["J2"].Value2 == "West2", "table deletion keeps data");
                run("name.set", "{name:'Revenue',refersTo:'=Sales!$B$2:$B$5'}"); Check(run("name.list", "{}")["names"].Any(), "defined names"); run("name.delete", "{name:'Revenue'}");
                run("note.set", "{sheet:'Sales',address:'A1',text:'Reviewed'}"); Check((string)sheet.Range["A1"].Comment.Text() == "Reviewed", "cell note"); run("note.delete", "{sheet:'Sales',address:'A1'}");
                run("sparkline.add", "{sheet:'Sales',address:'O2:O5',source:'B2:C5'}"); Check((int)sheet.Range["O2:O5"].SparklineGroups.Count > 0, "sparklines"); run("sparkline.delete", "{sheet:'Sales',address:'O2:O5'}");
                run("chart.create", "{sheet:'Sales',address:'A1:C5',name:'RevenueChart',type:'xlLineMarkers'}");
                run("chart.edit", "{sheet:'Sales',chart:'RevenueChart',path:'Axes(xlValue, xlPrimary)',properties:{MinimumScale:0,MaximumScale:300}}");
                run("chart.edit", "{sheet:'Sales',chart:'RevenueChart',path:'SeriesCollection(1).Points(2).Format.Fill.ForeColor',properties:{RGB:255}}");
                Check(run("chart.inspect", "{sheet:'Sales',chart:'RevenueChart',path:'Axes(2,1)',properties:['MaximumScale']}")["MaximumScale"].Value<double>() == 300, "axis and point editing");
                run("chart.series", "{sheet:'Sales',chart:'RevenueChart',action:'edit',index:2,type:'xlColumnClustered',axisGroup:2}"); Check((int)sheet.ChartObjects("RevenueChart").Chart.SeriesCollection(2).AxisGroup == 2, "combo chart secondary axis");
                run("chart.trendline", "{sheet:'Sales',chart:'RevenueChart',series:1,type:-4132}"); run("chart.errorbars", "{sheet:'Sales',chart:'RevenueChart',series:1,type:2,amount:5}");
                run("pivot.create", "{sheet:'Sales',address:'A1:D5',destinationSheet:'Summary',name:'RevenuePivot',rows:['Region'],filters:['Year'],values:[{field:'Revenue',function:'sum',caption:'Total Revenue'}]}");
                run("pivot.edit", "{sheet:'Summary',pivot:'RevenuePivot',layout:1,repeatLabels:true}");
                run("pivot.field", "{sheet:'Summary',pivot:'RevenuePivot',field:'Cost',addValue:true,function:'sum',caption:'Total Cost'}");
                run("pivot.field", "{sheet:'Summary',pivot:'RevenuePivot',field:'Year',orientation:2}");
                run("pivot.filter", "{sheet:'Summary',pivot:'RevenuePivot',field:'Region',visibleItems:['East','West']}");
                run("pivot.sort", "{sheet:'Summary',pivot:'RevenuePivot',field:'Region',descending:true,by:'Total Revenue'}");
                Check(run("pivot.inspect", "{sheet:'Summary',pivot:'RevenuePivot',field:'Region'}")["fields"].Count() >= 4, "pivot inspect fields and items");
                run("pivot.filter", "{sheet:'Summary',pivot:'RevenuePivot',field:'Region',clear:true}");
                run("pivot.calculated", "{sheet:'Summary',pivot:'RevenuePivot',name:'Profit',formula:'=Revenue-Cost'}");
                run("pivot.refresh", "{sheet:'Summary',pivot:'RevenuePivot'}");
                var slicer = run("slicer.add", "{sheet:'Summary',pivot:'RevenuePivot',field:'Region',name:'Regions'}"); Check((bool)slicer["applied"], "Excel native slicer");
                if(double.Parse((string)app.Version,System.Globalization.CultureInfo.InvariantCulture)>=15){
                    sheet.Range["Q1:R5"].Value2=new object[,]{{"Date","Amount"},{45000.0,1.0},{45001.0,2.0},{45002.0,3.0},{45003.0,4.0}};sheet.Range["Q2:Q5"].NumberFormat="yyyy-mm-dd";
                    run("pivot.create","{sheet:'Sales',address:'Q1:R5',destinationSheet:'Summary',destination:'J20',name:'DatePivot',rows:['Date'],values:[{field:'Amount',function:'sum'}]}");
                    Check((bool)run("slicer.add","{sheet:'Summary',pivot:'DatePivot',field:'Date',timeline:true,name:'Dates'}")["applied"],"native date timeline on Excel 2013+");
                    run("pivot.delete","{sheet:'Summary',pivot:'DatePivot'}");
                }
                run("sheet.layout", "{sheet:'Sales',freezeRows:1,tabColor:'#117A55',address:'A1:D5',columnWidth:16,print:{Orientation:2}}");
                string version = (string)run("snapshot.create", "{label:'all objects'}")["result"]["results"][0]["versionId"];
                run("sheet.delete", "{sheet:'Summary'}"); Check((int)book.Worksheets.Count == 2, "sheet deletion with snapshot");
                ai.OpenVersion(version); dynamic restoredBook = app.ActiveWorkbook; Check((int)restoredBook.Worksheets["Summary"].PivotTables().Count == 1 && (int)restoredBook.Worksheets["Sales"].ChartObjects().Count == 1, "full snapshot restores deleted sheet, pivot and chart"); Check((int)book.Worksheets.Count == 2, "snapshot restore preserves original live workbook"); restoredBook.Close(false); Marshal.ReleaseComObject((object)restoredBook); book.Activate();
                var risky = run("vba.write", "{module:'RiskTest',code:'Sub Test()\n Shell \"cmd\"\nEnd Sub'}"); Check(!(bool)risky["applied"] && (bool)risky["requiresManualApproval"], "dangerous VBA cannot auto apply");
                bool refused = false; try { ai.ApplyBatch(new[] { (string)risky["proposalId"] }, false, false); } catch (InvalidOperationException) { refused = true; } Check(refused, "host independently rejects automatic risky VBA apply");
                ai.Discard(new[] { (string)risky["proposalId"] });
                Check(JObject.FromObject(ai.SnapshotHistory())["versions"].Count() > 20, "multiple durable history versions");
            }
        }
        finally { if (book != null) { try { book.Close(false); } catch { } Marshal.ReleaseComObject((object)book); } if (app != null) { try { app.Quit(); } catch { } Marshal.ReleaseComObject((object)app); } }
    }
}
