using System;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using SheetPilot.Excel;
internal static partial class Program
{
    private static void AiUsabilityTests(dynamic app)
    {
        dynamic book=app.Workbooks.Add(-4167),sheet=book.Worksheets[1];sheet.Name="交互验证";
        try { using(var ai=new AiOperations(()=>(object)app)) {
            string snapshot=(string)JObject.FromObject(ai.Capture("sheet"))["snapshotId"];
            Func<string,string,JObject,string> propose=(kind,address,args)=>{args["kind"]=kind;args["address"]=address;args["sheet"]="交互验证";args["snapshotId"]=snapshot;return (string)JObject.FromObject(ai.Propose(args))["proposalId"];};
            var ids=new List<string>();
            for(int i=1;i<=21;i++)ids.Add(propose("formula","A1",new JObject{["formula"]="="+i}));
            ids.Add(propose("formulas","A1:B1",new JObject{["data"]=new JArray{new JArray{"=30","=40"}}}));
            ids.Add(propose("formula","A1",new JObject{["formula"]="=50"}));
            var result=JObject.FromObject(ai.ApplyBatch(ids.AsEnumerable().Reverse().ToArray()));
            Check((double)sheet.Range["A1"].Value2==50&&(double)sheet.Range["B1"].Value2==40,"23 accumulated overlapping proposals use latest creation order and retain other cells");
            Check(result["retiredIds"].Count()==23,"batch retires all superseded suggestions");
            ai.RestoreCheckpoint(true);Check(sheet.Range["A1"].Value2==null&&sheet.Range["B1"].Value2==null,"overlap merge restores original cells");ai.RestoreCheckpoint(false);
            var style1=propose("format","D1:E2",new JObject{["style"]=new JObject{["bold"]=true,["fill_color"]="#FF0000"}});
            var style2=propose("format","E2:F3",new JObject{["style"]=new JObject{["fill_color"]="#00FF00"}});
            ai.ApplyBatch(new[]{style2,style1});Check((bool)sheet.Range["E2"].Font.Bold&&(int)sheet.Range["E2"].Interior.Color==65280,"overlapping format properties merge without dropping unrelated properties");ai.RestoreCheckpoint(true);Check(!(bool)sheet.Range["E2"].Font.Bold&&(int)sheet.Range["E2"].Interior.Pattern==-4142,"merged format restores correctly");
            var old=propose("formula","G1",new JObject{["formula"]="=1"});for(int i=0;i<9;i++)snapshot=(string)JObject.FromObject(ai.Capture("sheet"))["snapshotId"];
            var fresh=propose("formula","G2",new JObject{["formula"]="=2"});ai.ApplyBatch(new[]{old,fresh});Check(sheet.Range["G1"].Value2==null&&(double)sheet.Range["G2"].Value2==2,"expired suggestion does not block fresh batch");
            var discard=propose("formula","H1",new JObject{["formula"]="=9"});ai.Discard(new[]{discard});ai.ApplyBatch(new[]{discard});Check(sheet.Range["H1"].Value2==null,"clear removes server-side proposal");
            ai.Approval(true,"workbook");ai.ConversationChanged();Check((bool)JObject.FromObject(ai.CheckpointState())["automatic"],"workbook approval survives conversation changes");
            for(int i=0;i<10;i++){snapshot=(string)JObject.FromObject(ai.Capture("sheet"))["snapshotId"];propose("formula","J1",new JObject{["formula"]="="+(i+1)});ai.ConversationChanged();}
            Check((double)sheet.Range["J1"].Value2==10,"workbook approval survives ten rounds and context eviction");
            ai.RestoreCheckpoint(true);Check((bool)JObject.FromObject(ai.CheckpointState())["automatic"],"workbook approval remains enabled after restore");propose("formula","J2",new JObject{["formula"]="=12"});Check(sheet.Range["J1"].Value2==null&&(double)sheet.Range["J2"].Value2==12,"edit after restore starts fresh protected checkpoint");
            dynamic other=app.Workbooks.Add(-4167);Check(!(bool)JObject.FromObject(ai.CheckpointState())["automatic"],"approval does not leak to another workbook");book.Activate();Check((bool)JObject.FromObject(ai.CheckpointState())["automatic"],"returning to approved workbook preserves approval");other.Close(false);
            ai.Approval(true,"session");ai.ConversationChanged();Check(!(bool)JObject.FromObject(ai.CheckpointState())["automatic"],"session-only option expires on conversation change");
            ai.Approval(true,"workbook");book.Close(false);book=app.Workbooks.Add(-4167);Check(!(bool)JObject.FromObject(ai.CheckpointState())["automatic"],"closed workbook grant cannot authorize a new workbook");
        }} finally {book.Close(false);}
    }
}
