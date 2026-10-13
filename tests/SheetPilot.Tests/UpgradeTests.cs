using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using SheetPilot.Core;
using SheetPilot.Excel;
internal static partial class Program
{
    private static void UpgradeCoreTests()
    {
        var s=new MaskSnapshot();s.Sheets.Add(new DataSheet {Name="主表",HeaderMerges={new HeaderMerge {Row=1,Column=1,Rows=1,Columns=4},new HeaderMerge {Row=4,Column=1,Rows=2,Columns=1},new HeaderMerge {Row=4,Column=2,Rows=1,Columns=2}}});
        Action<int,int,object> add=(r,c,v)=>s.Cells.Add(new DataCell {Sheet=0,Row=r,Column=c,Value=v});
        add(1,1,"贷款项目情况表");add(4,1,"机构名称");add(4,2,"金额");add(5,2,"本期金额");add(5,3,"上期余额");add(6,1,"机构甲");add(6,2,100.0);add(6,3,80.0);
        var rules=MaskInference.Suggest(s);
        Check(rules.All(r=>r.StartRow==6),"multi-row merged header ends before data");
        Check(rules.First(r=>r.Column==1).Kind==MaskKind.ChineseName,"merged identity header suggestion");
        Check(rules.First(r=>r.Column==2).Kind==MaskKind.Keep,"money remains opt-in");
        Check(rules.First(r=>r.Column==1).Header.Contains("机构名称"),"title is not mistaken for field header");
        s.Sheets.Add(new DataSheet {Name="附表"});s.Cells.Add(new DataCell{Sheet=1,Row=1,Column=1,Value="借款人名称"});s.Cells.Add(new DataCell{Sheet=1,Row=1,Column=2,Value="类型"});s.Cells.Add(new DataCell{Sheet=1,Row=2,Column=1,Value="机构甲"});s.Cells.Add(new DataCell{Sheet=1,Row=2,Column=2,Value="分类"});
        rules=MaskInference.Suggest(s);Check(rules.First(r=>r.Sheet==0 && r.Column==1).Domain==rules.First(r=>r.Sheet==1 && r.Column==1).Domain,"overlapping identity aliases share mapping");
        var frame=new FocusFrame{Viewport=new Rectangle(0,0,800,600),Panes=new[]{new Rectangle(0,0,800,600)},Cells=new[]{new Rectangle(360,280,80,30)},Bands=new Rectangle[0]};
        foreach(double t in new[]{0.0,.2,.5,.8}){var ripple=FocusOverlay.RippleFrame(frame,t,new ToolPreferences());Check(ripple.Ripple.All(b=>frame.Viewport.Contains(b.Bounds)&&!b.Bounds.IntersectsWith(frame.Cells[0])),"ripple clipped and active cell clear "+t);}
        Check(ReferenceEquals(FocusOverlay.RippleFrame(frame,1,new ToolPreferences()),frame),"ripple finite endpoint");
        for(int theme=0;theme<3;theme++)foreach(int size in new[]{32,48,64,96})using(var icon=RibbonIcons.Draw("GridPilot_Mask",size,theme))Check(icon.Width==size&&icon.GetPixel(0,0).A==0,"theme/icon transparency "+theme+"/"+size);
        var options=new ToolPreferences {RippleDuration=9999,RippleStrength=-1,IconTheme=99};options.Normalize();Check(options.RippleDuration==600&&options.RippleStrength==5&&options.IconTheme==2,"animation preference bounds");
        var matcher=new IdentityMatcher(new System.Collections.Generic.Dictionary<string,string>{{"Alpha","X"},{"Alpha Ltd","Y"}});
        Check(matcher.Replace("ALPHA LTD / alpha",CancellationToken.None)=="Y / X","longest case-insensitive identity matching");
        var mention=new MaskSnapshot();mention.Sheets.Add(new DataSheet{Name="数据"});mention.Cells.Add(new DataCell{Sheet=0,Row=2,Column=1,Value="机构甲"});mention.Cells.Add(new DataCell{Sheet=0,Row=2,Column=2,Value="备注：机构甲"});mention.Cells.Add(new DataCell{Sheet=0,Row=3,Column=1,Value="合计"});
        var mentionRule=new[]{new MaskRule{Sheet=0,Column=1,StartRow=2,Kind=MaskKind.ChineseName,Domain="机构"}};
        Reject(()=>MaskPlanner.Build(mention,mentionRule,CancellationToken.None),"mentions require an explicit replacement option");
        var linked=MaskPlanner.Build(mention,mentionRule,CancellationToken.None,true);
        Check(linked.Output.Cells[1].Value.Equals("备注："+linked.Output.Cells[0].Value),"same identity in free text shares replacement");
        Check(linked.Output.Cells[2].Value.Equals("合计") && linked.Changed==2,"summary placeholder preserved");
    }
    private static void UpgradeExcelTests(string sampleDirectory,string outputDirectory)
    {int pid=UpgradeExcelBody(sampleDirectory,outputDirectory);CopyExcelExitCheck(pid);if(ExcelGateway.LastMaskWriterProcess>0)CopyExcelExitCheck(ExcelGateway.LastMaskWriterProcess);}
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static int UpgradeExcelBody(string sampleDirectory,string outputDirectory)
    {
        outputDirectory=Path.GetFullPath(outputDirectory);Directory.CreateDirectory(outputDirectory);dynamic app=null;int pid=0;
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));app.Visible=false;app.DisplayAlerts=false;app.EnableEvents=false;app.AutomationSecurity=3;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out pid);Console.WriteLine("Upgrade Excel "+app.Version+" build "+app.Build+" pid="+pid);
            var gateway=new ExcelGateway(()=>(object)app);dynamic book=app.Workbooks.Add(-4167),sheet=book.Worksheets[1];
            sheet.Range["B2:E2"].Merge();sheet.Range["B2"].Value2="贷款项目情况表";
            sheet.Range["B4:B5"].Merge();sheet.Range["B4"].Value2="机构名称";sheet.Range["C4:D4"].Merge();sheet.Range["C4"].Value2="金额";sheet.Range["C5"].Value2="本期金额";sheet.Range["D5"].Value2="上期金额";
            sheet.Range["B6:D7"].Value2=new object[,]{{"机构甲",100.0,80.0},{"机构乙",200.0,170.0}};
            sheet.Range["B4:D7"].Interior.Color=0xE5F2DA;sheet.Range["B4:D7"].Font.Name="Arial";sheet.Range["B4:D7"].Font.Size=14;sheet.Range["C6:D7"].NumberFormat="#,##0.00";sheet.Range["B4:D7"].Borders.LineStyle=1;
            sheet.Columns[2].ColumnWidth=28;sheet.Rows[6].RowHeight=33;sheet.Range["XFD1048576"].Interior.Color=123456;
            book.Worksheets.Add(Type.Missing,sheet).Name="空白表";sheet.Activate();
            book.Saved=true;var capture=gateway.CaptureMaskWorkbook();Check(capture.Cells.Count==11,"inflated UsedRange counts content only");Check((bool)book.Saved,"snapshot does not dirty workbook");
            var rules=ExcelGateway.SuggestMaskRules(capture);Check(rules.First(r=>r.Column==2).StartRow==6,"actual Excel merged header detection");
            var result=MaskPlanner.Build(capture,rules,CancellationToken.None);string formatted=Path.Combine(outputDirectory,"format-preserved.xlsx");if(File.Exists(formatted))File.Delete(formatted);
            gateway.ExportMask(capture,result,formatted,CancellationToken.None,true);Check((bool)book.Saved,"format export does not dirty source");
            dynamic output=app.Workbooks.Open(formatted,0,true),target=output.Worksheets[1];
            Check((int)target.Range["B6"].Interior.Color==0xE5F2DA,"mask preserves fill");Check((string)target.Range["B6"].Font.Name=="Arial"&&(double)target.Range["B6"].Font.Size==14,"mask preserves fonts");
            Check((string)target.Range["C6"].NumberFormat=="#,##0.00","mask preserves number format");Check((bool)target.Range["B4"].MergeCells,"mask preserves merged cells");
            Check(Math.Abs((double)target.Rows[6].RowHeight-33)<.1 && Math.Abs((double)target.Columns[2].ColumnWidth-28)<.2,"mask preserves row height and column width");
            Check((string)target.Range["B6"].Value2!="机构甲","preserving styles does not restore identities");output.Close(false);book.Activate();
            string plain=Path.Combine(outputDirectory,"format-plain.xlsx");if(File.Exists(plain))File.Delete(plain);gateway.ExportMask(capture,result,plain,CancellationToken.None,false);
            output=app.Workbooks.Open(plain,0,true);Check(!(bool)output.Worksheets[1].Range["B4"].MergeCells,"format retention is optional");output.Close(false);book.Close(false);
            foreach(var format in new[]{new {Extension="xlsx",Code=51},new {Extension="xls",Code=56},new {Extension="xlsm",Code=52},new {Extension="xlsb",Code=50}})
            {
                book=app.Workbooks.Add(-4167);sheet=book.Worksheets[1];sheet.Range["A1:B3"].Value2=new object[,]{{"机构名称","类型"},{"测试机构甲","甲类"},{"测试机构乙","乙类"}};if(format.Extension!="xlsx")sheet.Range["A1:B3"].Interior.Color=0xDDEEFF;
                string inputPath=Path.Combine(outputDirectory,"format-source."+format.Extension);if(File.Exists(inputPath))File.Delete(inputPath);book.SaveAs(inputPath,format.Code);
                var formatSnapshot=gateway.CaptureMaskWorkbook();var formatResult=MaskPlanner.Build(formatSnapshot,ExcelGateway.SuggestMaskRules(formatSnapshot),CancellationToken.None,true);
                string formatOutput=Path.Combine(outputDirectory,"from-"+format.Extension+".xlsx");if(File.Exists(formatOutput))File.Delete(formatOutput);gateway.ExportMask(formatSnapshot,formatResult,formatOutput,CancellationToken.None,true);
                Console.WriteLine("MASK TIMING "+format.Extension+": "+ExcelGateway.LastMaskPerformance);
                output=app.Workbooks.Open(formatOutput,0,true);Check((format.Extension=="xlsx" || (int)output.Worksheets[1].Range["A2"].Interior.Color==0xDDEEFF) && (string)output.Worksheets[1].Range["A2"].Value2!="测试机构甲" && !(bool)output.HasVBProject,"formatted roundtrip from "+format.Extension);output.Close(false);book.Close(false);
            }
            foreach(string path in Directory.GetFiles(sampleDirectory,"*",SearchOption.AllDirectories).Where(p=>new[]{".xlsx",".xls",".xlsb",".xlsm"}.Contains(Path.GetExtension(p).ToLowerInvariant())))
            {
                var watch=Stopwatch.StartNew();dynamic sample=null;
                try
                {
                    sample=app.Workbooks.Open(path,0,true);bool saved=sample.Saved;
                    var snapshot=gateway.CaptureMaskWorkbook();var suggested=ExcelGateway.SuggestMaskRules(snapshot);
                    Console.WriteLine("SAMPLE "+Path.GetFileName(path)+" cells="+snapshot.Cells.Count+" fields="+suggested.Count+" capture+infer_ms="+watch.ElapsedMilliseconds);
                    Check((bool)sample.Saved==saved,"sample source saved state "+Path.GetFileName(path));
                    if(snapshot.Cells.Count<=30000)
                    {
                        try
                        {
                            var masked=MaskPlanner.Build(snapshot,suggested,CancellationToken.None,true);string file=Path.Combine(outputDirectory,Path.GetFileNameWithoutExtension(path)+"-masked.xlsx");if(File.Exists(file))File.Delete(file);
                            gateway.ExportMask(snapshot,masked,file,CancellationToken.None,true);Check(File.Exists(file),"sample roundtrip "+Path.GetFileName(path));
                        }
                        catch(InvalidOperationException ex){Console.WriteLine("EXPECTED REVIEW "+Path.GetFileName(path)+": "+ex.Message);}
                    }
                }
                catch(System.Runtime.InteropServices.COMException ex){Console.WriteLine("EXCEL OPEN REJECT "+Path.GetFileName(path)+": "+ex.Message);}
                catch(InvalidOperationException ex){Console.WriteLine("SAMPLE LIMIT "+Path.GetFileName(path)+": "+ex.Message);}
                finally{if(sample!=null)sample.Close(false);}
            }
        }
        finally{if(app!=null){try{app.Quit();}catch{}System.Runtime.InteropServices.Marshal.ReleaseComObject(app);}}
        return pid;
    }
}
