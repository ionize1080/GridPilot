using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Newtonsoft.Json.Linq;
using SheetPilot.Core;
using SheetPilot.Excel;

internal static partial class Program
{
    private static void SimpleMaskCoreTests()
    {
        Func<object,SimpleMaskOptions,string> apply=(v,o)=>{SimpleMask.Validate(o);return SimpleMask.Transform(v,o);};
        var basic=new SimpleMaskOptions();
        Check(apply("新加坡油气有限公司",basic)=="XJPYQYXGS","simple initials exact requested example");
        Check(apply("新加坡油氣有限公司",basic)=="XJPYQYXGS","simple traditional Chinese initials");
        Check(SimpleMask.Letters.Length==26&&SimpleMask.Letters.Distinct().Count()==26,"26 fixed nonphonetic English characters");
        Check(apply("ABCDEFGHIJKLMNOPQRSTUVWXYZ",basic)==SimpleMask.Letters,"fixed uppercase alphabet");
        Check(apply("abcdefghijklmnopqrstuvwxyz",basic)==string.Concat(SimpleMask.Letters.Select(c=>"〔小〕"+c)),"lowercase preserves case marker");
        Check(apply("0123456789",basic)=="7391058264","fixed digit permutation");
        Check(apply("A1! \t\r\n，。😀é",basic)==apply("A1! \t\r\n，。😀é",basic),"symbols whitespace emoji stable");
        string symbols="!@#$%^&*()[]{}<>?，。！？￥…—😀\t\n";Check(!symbols.Any(c=>c>127&&apply(symbols,basic).Contains(c)),"symbols never pass through unchanged");
        Check(apply("新加坡",new SimpleMaskOptions{Prefix="客户-",Suffix="-X"})=="客户-XJP-X","optional affixes");
        var swap=new SimpleMaskOptions{Mode="swap",EdgeCount=2};Check(apply("P123456",swap)=="56234P1","requested swap example");
        foreach(string value in new[]{"中文标点！?","😀A12é中","P123456","e\u0301A12😀Z","a\nb\tc\rd"})Check(apply(apply(value,swap),swap)==value,"swap inverse Unicode "+value);
        Check(apply("😀A中!",new SimpleMaskOptions{Mode="reverse"})=="!中A😀","reverse supplementary Unicode");
        Reject(()=>apply("abc",swap),"short values skip");Reject(()=>apply(" ",basic),"whitespace skip");Reject(()=>apply(true,basic),"boolean skip");
        var regex=new SimpleMaskOptions{Mode="regex"};Check(apply("P123456",regex)=="56234P1","regex captures swap");
        Check(apply("AB123",new SimpleMaskOptions{Mode="regex",Pattern="(?<head>[A-Z]+)(?<n>[0-9]+)",Replacement="${n}-${head}-$$"})=="123-AB-$","named capture and literal dollar");
        Check(apply("abc",new SimpleMaskOptions{Mode="regex",Pattern=".+",Replacement=""})=="","regex empty replacement valid");
        Reject(()=>apply("😀",new SimpleMaskOptions{Mode="regex",Pattern="^.",Replacement=""}),"regex cannot split surrogate");
        Reject(()=>SimpleMask.Validate(new SimpleMaskOptions{Mode="regex",Pattern="["}),"invalid regex rejected before batch");
        Reject(()=>apply("abc",new SimpleMaskOptions{Mode="regex",Pattern="[0-9]+"}),"unmatched regex skip");
        Reject(()=>apply(new string('a',32767),new SimpleMaskOptions{Mode="regex",Pattern="a",Replacement="$`$'$`$'"}),"bounded regex replacement expansion");
        foreach(var pair in new[]{new[]{"+","100","00123","223"},new[]{"-","3","-2.5","-5.5"},new[]{"*","1.5","2.2","3.3"},new[]{"/","4","10","2.5"},new[]{"+","1","1e3","1001"},new[]{"+","1","12345678901234567890123456","12345678901234567890123457"}})
            Check(apply(pair[2],new SimpleMaskOptions{Mode="number",Operator=pair[0],Operand=pair[1]})==pair[3],"exact decimal arithmetic "+pair[0]+" "+pair[2]);
        Check(apply(0.1,new SimpleMaskOptions{Mode="number",Operand="0.2"})=="0.3","numeric cells arithmetic");
        Check(apply("10",new SimpleMaskOptions{Mode="number",Operand="5",Prefix="=",Suffix="!"})=="=15!","numeric output affixes");
        foreach(string bad in new[]{"1,000","12%","2024/1/1","NaN","Infinity","12345678901234567890123456789","1e-40","1.23e-28","￥100"})Reject(()=>apply(bad,new SimpleMaskOptions{Mode="number"}),"bad/inexact number "+bad);
        Reject(()=>SimpleMask.Validate(new SimpleMaskOptions{Mode="number",Operator="/",Operand="0"}),"division by zero friendly validation");
        try{apply("1e28",new SimpleMaskOptions{Mode="number",Operator="*",Operand="100"});Check(false,"overflow");}catch(OverflowException){Check(true,"decimal overflow");}
        var snapshot=new SimpleMaskSnapshot();snapshot.Blocks.Add(new SimpleMaskBlock{Row=1,Column=1,Rows=3,Columns=1});
        snapshot.Cells.Add(new SimpleMaskCell{Address="$A$1",Row=1,Column=1,Before="1"});snapshot.Cells.Add(new SimpleMaskCell{Address="$A$2",Row=2,Column=1,Before="secret"});snapshot.Cells.Add(new SimpleMaskCell{Address="$A$3",Row=3,Column=1,Before="2"});
        var result=SimpleMask.Build(snapshot,new SimpleMaskOptions{Mode="number",Operand="1"},CancellationToken.None);
        Check(result.Changed==2&&result.Skipped==1&&result.Cells[2].After=="3","bad cell does not stop remaining work");
        var clip=JObject.FromObject(MaskPreview.Clipboard(SimpleMask.Output(snapshot,result),0,1,3));Check(!((string)clip["text"]).Contains("secret")&&((string)clip["text"]).Contains("\r\n\r\n"),"skipped sensitive values absent from clipboard and geometry retained");
        snapshot.Cells[0].Before="=1+1";result=SimpleMask.Build(snapshot,new SimpleMaskOptions{Mode="reverse",Prefix="="},CancellationToken.None);clip=JObject.FromObject(MaskPreview.Clipboard(SimpleMask.Output(snapshot,result),0,1,3));Check(((string)clip["text"]).StartsWith("'="),"clipboard formula protection");
        var cancel=new CancellationTokenSource();cancel.Cancel();try{SimpleMask.Build(snapshot,basic,cancel.Token);Check(false,"cancel");}catch(OperationCanceledException){Check(true,"batch cancellation");}
        snapshot.Cells.Clear();for(int i=0;i<100;i++)snapshot.Cells.Add(new SimpleMaskCell{Before=new string('a',10000)+"!"});var timer=Stopwatch.StartNew();result=SimpleMask.Build(snapshot,new SimpleMaskOptions{Mode="regex",Pattern="^(a+)+$"},CancellationToken.None);Check(result.Skipped==100&&timer.ElapsedMilliseconds<5000,"catastrophic regex total budget bounded");
        var culture=Thread.CurrentThread.CurrentCulture;try{Thread.CurrentThread.CurrentCulture=new CultureInfo("fr-FR");Check(apply("1.2",new SimpleMaskOptions{Mode="number",Operand="0.1"})=="1.3","numeric rules culture independent");}finally{Thread.CurrentThread.CurrentCulture=culture;}
        snapshot.Cells.Clear();for(int i=0;i<100000;i++)snapshot.Cells.Add(new SimpleMaskCell{Before="新加坡A123!"});timer.Restart();result=SimpleMask.Build(snapshot,basic,CancellationToken.None);Check(result.Changed==100000&&result.Cells[0].After==result.Cells[99999].After,"100k deterministic batch");Console.WriteLine("100k simple mask ms="+timer.ElapsedMilliseconds);
    }
    private static void SimpleMaskExcelTests()
    {
        dynamic app=null;string dir=Path.Combine(Path.GetTempPath(),"GridPilot-simple-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(dir);
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            dynamic book=app.Workbooks.Add(-4167),sheet=book.Worksheets[1];sheet.Name="合成数据";var gateway=new ExcelGateway(()=>(object)app);
            sheet.Range["A1:C4"].NumberFormat="@";sheet.Range["A1:C4"].Value2=new object[,]{{"新加坡油气有限公司","P123456","00123"},{"ABcd!😀","短","12.5"},{"  ","P765432","bad"},{"plain",null,null}};
            sheet.Range["D1"].Formula="=1+1";sheet.Range["D2"].Formula="=1/0";sheet.Range["E1:F1"].Merge();sheet.Range["E1"].Value2="合并";sheet.Range["G1"].Value2=true;
            sheet.Range["H1"].NumberFormat="@";sheet.Range["H1"].Value2="=1+1";sheet.Range["I1"].Value2=123.45;
            sheet.Range["A:A"].Select();var snapshot=gateway.CaptureSimpleMask(CancellationToken.None);Check(snapshot.Cells.Count==4&&snapshot.Blocks[0].Rows==4&&snapshot.Blocks[0].Columns==1,"whole column bounded to contents");
            sheet.Range["1:1"].Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);Check(snapshot.Cells.Count==8&&snapshot.Blocks[0].Rows==1,"whole row bounded and merged child absent");
            sheet.Range["A1"].Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);Check(snapshot.Cells.Count==1,"SpecialCells single cell cannot expand outside selection");
            dynamic union=app.Union(sheet.Range["A1:A4"],sheet.Range["C1:C3"]);union.Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);Check(snapshot.Cells.Count==7&&snapshot.Blocks.Count==2,"multi-area capture");
            sheet.Range["A1:I4"].Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);var result=SimpleMask.Build(snapshot,new SimpleMaskOptions(),CancellationToken.None);
            Check(result.Cells.Count(c=>c.Skip=="公式")==2&&result.Cells.Any(c=>c.Skip=="合并单元格")&&result.Cells.Any(c=>c.Skip=="空白")&&result.Cells.Any(c=>c.Skip=="布尔值"),"formula merge whitespace boolean skip");
            string export=Path.Combine(dir,"output.xlsx");gateway.ExportSimpleMask(snapshot,result,export,CancellationToken.None);Check((string)sheet.Range["A1"].Value2=="新加坡油气有限公司","export source untouched");
            dynamic output=app.Workbooks.Open(export,0,true),outSheet=output.Worksheets[1];Check((string)outSheet.Range["A1"].Value2=="XJPYQYXGS"&&(object)outSheet.Range["D1"].Value2==null&&(object)outSheet.Range["E1"].Value2==null,"actual XLSX masked and skipped slots blank");Check(!(bool)outSheet.Range["A1:I4"].HasFormula,"output contains no formulas");output.Close(false);
            Reject(()=>gateway.ExportSimpleMask(snapshot,result,export,CancellationToken.None),"existing destination not overwritten");
            Reject(()=>gateway.ApplySimpleMask(snapshot,result,false,CancellationToken.None),"backup confirmation enforced by host");
            string recovery=Path.Combine(dir,"recovery.json");gateway.ExportSimpleRecovery(snapshot,result,recovery);var json=JObject.Parse(File.ReadAllText(recovery));Check((string)json["schema"]=="GridPilot.SimpleMask.Recovery/1"&&json["cells"].Any(c=>(string)c["Before"]=="新加坡油气有限公司"),"recovery contains originals options types coordinates");
            // Unrelated workbook is active: writes must still resolve the original snapshot target.
            dynamic other=app.Workbooks.Add(-4167);other.Worksheets[1].Range["A1"].Value2="unrelated";
            sheet.Range["B1"].Value2="edited";string message=gateway.ApplySimpleMask(snapshot,result,true,CancellationToken.None);
            Check((string)sheet.Range["A1"].Value2=="XJPYQYXGS"&&(string)sheet.Range["B1"].Value2=="edited"&&(string)other.Worksheets[1].Range["A1"].Value2=="unrelated","writes original sheet and skips stale cells while continuing");
            Check((string)sheet.Range["D1"].Formula=="=1+1"&&(string)sheet.Range["E1"].Value2=="合并","formulas and merged cells untouched");
            Check(gateway.CanUndo,"original write offers undo");gateway.UndoLast();Check((string)sheet.Range["A1"].Value2=="新加坡油气有限公司"&&(double)sheet.Range["I1"].Value2==123.45&&(string)sheet.Range["I1"].NumberFormat=="General","undo restores text number and format");other.Close(false);
            book.Activate();sheet.Range["H1"].Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);result=SimpleMask.Build(snapshot,new SimpleMaskOptions{Mode="regex",Pattern=".*",Replacement="=1+1"},CancellationToken.None);gateway.ApplySimpleMask(snapshot,result,true,CancellationToken.None);Check(!(bool)sheet.Range["H1"].HasFormula,"in-place formula-like output stays literal text");gateway.UndoLast();
            sheet.Range["C1:C3"].Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);result=SimpleMask.Build(snapshot,new SimpleMaskOptions{Mode="number",Operator="*",Operand="2"},CancellationToken.None);gateway.ApplySimpleMask(snapshot,result,true,CancellationToken.None);Check((string)sheet.Range["C1"].Value2=="246"&&(string)sheet.Range["C2"].Value2=="25"&&(string)sheet.Range["C3"].Value2=="bad","numeric strings arithmetic and per-cell skip");gateway.UndoLast();
            sheet.Cells.Locked=true;sheet.Range["A1"].Locked=false;sheet.Protect("test");sheet.Range["A1:B1"].Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);Check(snapshot.Cells.Single(c=>c.Address=="$B$1").Skip=="受保护单元格","locked protected cell identified");sheet.Unprotect("test");
            sheet.Range["Z1"].Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);Check(snapshot.Cells.Count==0,"empty selection supported");
            var data=new object[5000,1];for(int i=0;i<5000;i++)data[i,0]="P123456";sheet.Range["L1:L5000"].NumberFormat="@";sheet.Range["L1:L5000"].Value2=data;
            sheet.Range["L:L"].Select();var clock=Stopwatch.StartNew();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);Console.WriteLine("5000 cell Excel capture ms="+clock.ElapsedMilliseconds);Check(snapshot.Cells.Count==5000,"5000 whole-column values captured without truncation");
            result=SimpleMask.Build(snapshot,new SimpleMaskOptions{Mode="swap"},CancellationToken.None);var partialCancel=new CancellationTokenSource();gateway.ApplySimpleMask(snapshot,result,true,partialCancel.Token,(phase,percent,current,total)=>{if(current>=128)partialCancel.Cancel();});
            Check((string)sheet.Range["L1"].Value2=="56234P1"&&(string)sheet.Range["L129"].Value2=="P123456"&&result.Cells.Any(c=>c.WriteStatus=="取消后未写入"),"cancel keeps completed writes and marks remaining cells");gateway.UndoLast();Check((string)sheet.Range["L1"].Value2=="P123456","partial canceled writes can be undone");
            sheet.Range["M1"].NumberFormat="@";sheet.Range["M1"].Value2="remove";sheet.Range["M1"].Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);result=SimpleMask.Build(snapshot,new SimpleMaskOptions{Mode="regex",Pattern=".+",Replacement=""},CancellationToken.None);gateway.ApplySimpleMask(snapshot,result,true,CancellationToken.None);Check(string.IsNullOrEmpty(SimpleMask.Text((object)sheet.Range["M1"].Value2)),"empty replacement writes blank");gateway.UndoLast();Check((string)sheet.Range["M1"].Value2=="remove","empty replacement undo restores source");
            sheet.Range["M1"].Select();snapshot=gateway.CaptureSimpleMask(CancellationToken.None);result=SimpleMask.Build(snapshot,new SimpleMaskOptions(),CancellationToken.None);sheet.Range["M1"].NumberFormat="General";sheet.Range["M1"].Formula="=123";gateway.ApplySimpleMask(snapshot,result,true,CancellationToken.None);Check((bool)sheet.Range["M1"].HasFormula,"formula added after preview is never overwritten");
                        Check((bool)app.EnableEvents,"Excel events restored");gateway.ReleaseUndoEvents();book.Close(false);
        }
        finally{if(app!=null){try{app.Quit();}catch{}Marshal.ReleaseComObject(app);}Directory.Delete(dir,true);}
    }
}
