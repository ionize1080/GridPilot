using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SheetPilot.Core;
using SheetPilot.Excel;

internal static partial class Program
{
    private static void MultiSortTests(bool excel,string sampleFolder=null)
    {
        var ranks=CustomSort.Ranks(new[]{"二","一","二","其他","别的"},new SortOptions(),out int matched);
        Check(CustomSort.Order(new[]{"国际二部","国际一部","业务二部","业务一部"},new SortOptions{Rules="(国际)[中文数字]\n(业务)[中文数字]"}).SequenceEqual(new[]{1,0,3,2}),"department example Chinese rank");
        Check(ranks[0]==ranks[2]&&ranks[3]==ranks[4]&&ranks[1]<ranks[0]&&matched==3,"equal semantic keys share ranks for next level");
        Check(SortBlocks.Create(8,new[]{new[]{1,0,3,1},new[]{3,1,3,1}}).Select(x=>x.Length).SequenceEqual(new[]{1,5,1,1}),"overlapping vertical merges form indivisible record");
        Reject(()=>new MultiSortOptions{Levels=Enumerable.Range(0,21).Select(i=>new SortLevel()).ToList()}.Validate(3),"21st sorting level rejected");
        new MultiSortOptions{Levels=Enumerable.Range(0,20).Select(i=>new SortLevel{Column=i}).ToList()}.Validate(20);Check(true,"20 levels accepted");
        if(!excel)return;
        dynamic app=null;var com=new MaskComLifetime();
        try
        {
            app=com.Hold(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true)));app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            dynamic book=com.Hold(app.Workbooks.Add(-4167)),sheet=com.Hold(book.Worksheets[1]);var gateway=new ExcelGateway(()=>(object)app);
            var options=new MultiSortOptions {Levels=new List<SortLevel>{new SortLevel{Column=0,Mode="custom",Rules="[中文数字]"},new SortLevel{Column=1,Descending=true},new SortLevel{Column=2,Mode="natural"}}};
            options=JObject.FromObject(options).ToObject<MultiSortOptions>();Check(options.Levels.Count==3&&options.Levels[0].Mode=="custom","IPC roundtrip does not prepend an invisible default sort level");
            sheet.Range["A1:E7"].Value2=new object[,]{{"部门","客户","项目","金额","公式"},{"二","A","项目10",10,null},{"一","B","项目2",20,null},{"一","A","项目1",30,null},{"一","B","项目1",40,null},{"二","A","项目2",50,null},{"一","B","项目1",60,null}};
            sheet.Range["E2:E7"].FormulaR1C1="=RC[-1]*2";sheet.Range["F2"].Value2="旁边数据";sheet.Range["G2"].Formula="=SUM(D2:D7)";sheet.Range["A1:E7"].Select();
            var snapshot=gateway.CaptureSort();var plan=gateway.PlanSort(snapshot,options);
            Check(plan.Order.SequenceEqual(new[]{3,5,1,2,4,0}),"three-level opposite directions and stable final ties");
            Check((int)book.Worksheets.Count==1&&(int)app.Workbooks.Count==1,"scratch workbook closed without source sheet change");
            gateway.ApplySort(snapshot,plan);Check((double)sheet.Range["D2"].Value2==40&&(double)sheet.Range["E2"].Value2==80,"apply matches preview and relative formulas move");
            Check((string)sheet.Range["F2"].Value2=="旁边数据"&&(double)sheet.Range["G2"].Value2==210,"occupied adjacent column and external reference preserved");gateway.UndoLast();
            Check((double)sheet.Range["D2"].Value2==10&&(double)sheet.Range["E2"].Value2==20,"multi sort exact undo");
            sheet.Range["A1:E7"].Select();snapshot=gateway.CaptureSort();plan=gateway.PlanSort(snapshot,options);sheet.Range["D2"].Value2=99;Reject(()=>gateway.ApplySort(snapshot,plan),"stale values rejected before sort");sheet.Range["D2"].Value2=10;
            sheet.Range["A1:E7"].Select();snapshot=gateway.CaptureSort();plan=gateway.PlanSort(snapshot,options);sheet.Range["B2:C2"].Merge();Reject(()=>gateway.ApplySort(snapshot,plan),"stale merge layout rejected");sheet.Range["B2:C2"].UnMerge();
            dynamic mixed=com.Hold(book.Worksheets.Add());mixed.Name="类型";mixed.Range["A1:B12"].Value2=new object[,]{{"键","ID"},{10,0},{2,1},{"'10",2},{"'2",3},{false,4},{true,5},{null,6},{"'001",7},{-2.5,8},{"'=1+1",9},{new ErrorWrapper(unchecked((int)0x800A07FA)),10}};mixed.Range["A2"].Select();
            var native=new MultiSortOptions();snapshot=gateway.CaptureSort();plan=gateway.PlanSort(snapshot,native);
            // Compare our native-key planning against Excel sorting the actual typed source.
            dynamic copy=com.Hold(book.Worksheets.Add());mixed.Range["A1:B12"].Copy(copy.Range["A1"]);copy.Sort.SortFields.Clear();copy.Sort.SortFields.Add(copy.Range["A2:A12"],0,1,Type.Missing,0);copy.Sort.SetRange(copy.Range["A1:B12"]);copy.Sort.Header=1;copy.Sort.Orientation=1;copy.Sort.Apply();
            var nativeIds=(Array)copy.Range["B2:B12"].Value2;Console.WriteLine("PLAN="+string.Join(",",plan.Order)+" NATIVE="+string.Join(",",Enumerable.Range(1,11).Select(i=>Convert.ToInt32(nativeIds.GetValue(i,1)))));for(int i=1;i<12;i++)Console.WriteLine("KEY "+(i-1)+" "+snapshot.Values[i,0]+" TYPE="+snapshot.Values[i,0]?.GetType().FullName);Check(plan.Order.SequenceEqual(Enumerable.Range(1,11).Select(i=>Convert.ToInt32(nativeIds.GetValue(i,1)))),"mixed typed values, blanks, errors and literal equals match Excel native");
            mixed.Activate();mixed.Range["A2"].Select();native.Levels[0].TextAsNumbers=true;native.Levels[0].Descending=true;plan=gateway.PlanSort(gateway.CaptureSort(),native);Check(plan.Order.Last()==6,"descending keeps true blanks at bottom");
            dynamic merges=com.Hold(book.Worksheets.Add());merges.Name="跨行合并";merges.Range["A1:D7"].Value2=new object[,]{{"部门","说明","数值","公式"},{"二","b1",20,null},{null,"b2",21,null},{"一","a1",10,null},{null,"a2",11,null},{null,"a3",12,null},{"三","c1",30,null}};
            merges.Range["A2:A3"].Merge();merges.Range["A4:A6"].Merge();merges.Range["D2:D7"].FormulaR1C1="=RC[-1]*2";merges.Range["A2"].Select();
            var mergedOptions=new MultiSortOptions{Levels=new List<SortLevel>{new SortLevel{Mode="custom",Rules="[中文数字]"}}};snapshot=gateway.CaptureSort();plan=gateway.PlanSort(snapshot,mergedOptions);
            Check(plan.Order.SequenceEqual(new[]{2,3,4,0,1,5})&&plan.Records==3,"different-height merged blocks planned without splitting");gateway.ApplySort(snapshot,plan);
            Check((int)merges.Range["A2"].MergeArea.Rows.Count==3&&(double)merges.Range["D3"].Value2==22&&(string)merges.Range["B5"].Value2=="b1","merged blocks move and formulas retain dependencies");gateway.UndoLast();Check((int)merges.Range["A2"].MergeArea.Rows.Count==2&&(string)merges.Range["A2"].Value2=="二","vertical merge undo restores original spans");
            dynamic large=com.Hold(book.Worksheets.Add());large.Name="20级20000行";const int count=20000;var bulk=new object[count+1,21];
            for(int c=0;c<21;c++)bulk[0,c]="列"+(c+1);
            for(int i=0;i<count;i++){for(int c=0;c<20;c++)bulk[i+1,c]=(i/(c+1))%7;bulk[i+1,20]=i;}
            large.Range["A1:U20001"].Value2=bulk;large.Range["A2"].Select();var twenty=new MultiSortOptions{Levels=Enumerable.Range(0,20).Select(c=>new SortLevel{Column=c,Descending=c%2==1}).ToList()};
            var clock=Stopwatch.StartNew();snapshot=gateway.CaptureSort();plan=gateway.PlanSort(snapshot,twenty);Console.WriteLine("20k x 20 levels capture+preview ms="+clock.ElapsedMilliseconds);
            var expected=Enumerable.Range(0,count).ToArray();Array.Sort(expected,(a,b)=>{for(int c=0;c<20;c++){int diff=((int)bulk[a+1,c]).CompareTo((int)bulk[b+1,c]);if(diff!=0)return c%2==1?-diff:diff;}return a.CompareTo(b);});
            Check(plan.Order.SequenceEqual(expected),"20-level 20k complete order matches independent comparator");var page=JObject.FromObject(gateway.SortPage(snapshot,plan,19900,100));Check(page["rows"].Count()==100&&(int)page["rows"].Last()["position"]==20000,"final page includes row 20000");
            clock.Restart();gateway.ApplySort(snapshot,plan);Console.WriteLine("20k x 21 apply ms="+clock.ElapsedMilliseconds);Check((double)large.Range["U2"].Value2==expected[0]&&(double)large.Range["U20001"].Value2==expected.Last(),"20k apply first and last rows agree with preview");gateway.UndoLast();Check((double)large.Range["U2"].Value2==0&&(double)large.Range["U20001"].Value2==19999,"20k undo");
            dynamic hiddenHeader=com.Hold(book.Worksheets.Add());hiddenHeader.Range["A1:B4"].Value2=new object[,]{{"键","值"},{3,30},{1,10},{2,20}};
            dynamic hiddenTable=com.Hold(hiddenHeader.ListObjects.Add(1,hiddenHeader.Range["A1:B4"],Type.Missing,1));hiddenTable.ShowHeaders=false;hiddenTable.DataBodyRange.Cells[1,1].Select();
            snapshot=gateway.CaptureSort();var hiddenOptions=new MultiSortOptions{HasHeader=false};plan=gateway.PlanSort(snapshot,hiddenOptions);
            Check(!snapshot.TableHasHeader&&plan.Rows==3&&plan.Order.SequenceEqual(new[]{1,2,0}),"hidden table header does not skip first record");gateway.ApplySort(snapshot,plan);
            Check(!(bool)hiddenTable.ShowHeaders&&(double)hiddenTable.DataBodyRange.Cells[1,1].Value2==1&&(double)hiddenTable.DataBodyRange.Cells[3,1].Value2==3,"hidden-header table native apply preserves layout");gateway.UndoLast();Check((double)hiddenTable.DataBodyRange.Cells[1,1].Value2==3,"hidden-header table undo");
            if(sampleFolder!=null)MultiSortSamples(app,gateway,sampleFolder);
            gateway.ReleaseUndoEvents();book.Close(false);
        }
        finally{if(app!=null){try{app.Quit();}catch{}}com.Dispose();}
    }
    private static void MultiSortSamples(dynamic app,ExcelGateway gateway,string folder)
    {
        var files=Directory.GetFiles(folder).Where(p=>new[]{".xlsx",".xlsm",".xls"}.Contains(Path.GetExtension(p).ToLowerInvariant())).Take(20).ToArray();
        foreach(var file in files)
        {
            string before;using(var sha=System.Security.Cryptography.SHA256.Create())before=Convert.ToBase64String(sha.ComputeHash(File.ReadAllBytes(file)));
            dynamic source=app.Workbooks.Open(Path.GetFullPath(file),0,false);int tested=0;
            try
            {
                for(int i=1;i<=(int)source.Worksheets.Count&&tested<3;i++)
                {
                    dynamic sheet=source.Worksheets[i];if((int)sheet.Visible!=-1||(bool)sheet.ProtectContents)continue;
                    sheet.Activate();dynamic used=sheet.UsedRange;int rows=Math.Min(20001,(int)used.Rows.Count),cols=Math.Min(30,(int)used.Columns.Count);if(rows<2)continue;
                    // Samples are tested in-memory and never saved. Use complete used range
                    // where feasible, including merged report titles; no header removal.
                    if((long)(int)used.Rows.Count*(int)used.Columns.Count<=2000000&&(int)used.Rows.Count<=200001)used.Select();else sheet.Cells[(int)used.Row,(int)used.Column].Resize[rows,cols].Select();
                    try
                    {
                        var snapshot=gateway.CaptureSort();var opts=new MultiSortOptions{HasHeader=false,Levels=new List<SortLevel>{new SortLevel{Column=0},new SortLevel{Column=Math.Min(1,snapshot.ColumnCount-1),Descending=true}}};
                        if(snapshot.IsTable)opts.HasHeader=snapshot.TableHasHeader;var plan=gateway.PlanSort(snapshot,opts);gateway.ApplySort(snapshot,plan);
                        if(plan.Order.Where((v,n)=>v!=n).Any())gateway.UndoLast();tested++;
                        Console.WriteLine("SAMPLE PASS "+Path.GetFileName(file)+" / "+(string)sheet.Name+" rows="+snapshot.RowCount+" merges="+snapshot.MergeCount);
                    }
                    catch(InvalidOperationException ex){Console.WriteLine("SAMPLE BOUNDARY "+Path.GetFileName(file)+" / "+(string)sheet.Name+": "+ex.Message);}
                }
            }
            finally{source.Close(false);}
            using(var sha=System.Security.Cryptography.SHA256.Create())Check(before==Convert.ToBase64String(sha.ComputeHash(File.ReadAllBytes(file))),"sample source unchanged: "+Path.GetFileName(file));
        }
    }
    private static void MultiSortUiLaunch(string addin,string folder,string file=null)
    {
        dynamic app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=true;app.DisplayAlerts=false;app.AutomationSecurity=3;
        var gateway=new ExcelGateway(()=>(object)app);if(file==null)gateway.InsertSortExample();else app.Workbooks.Open(Path.GetFullPath(file),0,false);dynamic book=app.ActiveWorkbook,sheet=app.ActiveSheet;
        book.Worksheets.Add(Type.Missing,sheet).Name="第二张表";sheet.Activate();sheet.Range["A2"].Select();
        app.RegisterXLL(Path.GetFullPath(addin));app.Run("GridPilot_ShowNavigation");app.Run("GridPilot_Sort");
        Console.WriteLine("MULTI UI READY");
        try{PumpUntil(()=>File.Exists(Path.Combine(folder,"stop-ui")),3600000);}
        finally{app.DisplayAlerts=false;while(app.Workbooks.Count>0)app.Workbooks[1].Close(false);app.Quit();Marshal.ReleaseComObject(app);}
    }
}
