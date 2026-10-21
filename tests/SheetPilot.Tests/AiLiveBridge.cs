using System;
using System.Text;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SheetPilot.Excel;

internal static partial class Program
{
    // Isolated synthetic Excel host for the native Chatbox -> live provider -> MCP test.
    // It never opens existing workbooks and never receives the provider credential.
    private static void AiLiveBridge()
    {
        Console.InputEncoding=Encoding.UTF8;Console.OutputEncoding=new UTF8Encoding(false);
        dynamic app=null,book=null;
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));
            app.Visible=false;app.DisplayAlerts=false;app.AutomationSecurity=3;
            book=app.Workbooks.Add(-4167);dynamic sheet=book.Worksheets[1];sheet.Name="销售";
            sheet.Range["A1:C4"].Value2=new object[,]{{"项目","收入","成本"},{"甲",100.0,40.0},{"乙",200.0,90.0},{"丙",150.0,70.0}};
            sheet.Range["A1:C4"].Select();
            using(var ai=new AiOperations(()=>(object)app))
            {
                ai.Progress=(phase,percent,current,total)=>Console.WriteLine(JsonConvert.SerializeObject(new{type="progress",phase,percent,current,total}));
                string line;
                while((line=Console.ReadLine())!=null)
                {
                    var m=JObject.Parse(line);object result=null;
                    try
                    {
                        switch((string)m["action"])
                        {
                            case "aiReady": result=new{mode="window",version="1.20.3"};break;
                            case "aiMode":result=new{mode=(string)m["mode"]};break;
                            case "aiContext":result=ai.Context();break;
                            case "aiCapture":result=ai.Capture((string)m["scope"]);break;
                            case "aiRead":result=ai.Read(m);break;
                            case "aiInspect":result=ai.Metadata((string)m["snapshotId"]);break;
                            case "aiPropose":result=ai.Propose(m);break;
                            case "aiApplyBatch":result=ai.ApplyBatch(m["proposalIds"].ToObject<string[]>());break;
                            case "aiApproval":result=ai.Approval((bool)m["enabled"],(string)m["scope"]??"session");break;
                            case "aiConversation":result=ai.ConversationChanged();break;
                            case "aiDiscard":result=ai.Discard(m["proposalIds"]?.ToObject<string[]>());break;
                            case "aiCheckpoint":result=ai.CheckpointState();break;
                            case "aiUndo":result=ai.RestoreCheckpoint(true);break;
                            case "aiRestore":result=ai.RestoreCheckpoint((string)m["position"]=="before");break;
                            case "testPivot": dynamic ps=book.Worksheets[(string)m["sheet"]],pt=ps.PivotTables(1);result=new{count=(int)ps.PivotTables().Count,total=(double)pt.GetPivotData((string)pt.DataFields(1).Name).Value2};break;
                            case "testRead":
                                result=new{value=(object)sheet.Range[(string)m["address"]].Value2,formula=(object)sheet.Range[(string)m["address"]].Formula,bold=(object)sheet.Range[(string)m["address"]].Font.Bold,charts=(int)sheet.ChartObjects().Count};break;
                            case "testLarge":sheet.Range["K1:L100001"].Value2=1;result=new{ready=true};break;
                            default:throw new InvalidOperationException("Unknown bridge action");
                        }
                        Console.WriteLine(JsonConvert.SerializeObject(new{id=(int)m["id"],data=result}));
                    }
                    catch(Exception ex){Console.WriteLine(JsonConvert.SerializeObject(new{id=(int)m["id"],error=ex.Message}));}
                }
            }
        }
        finally {if(book!=null){try{book.Close(false);}catch{}Marshal.ReleaseComObject(book);}if(app!=null){try{app.Quit();}catch{}Marshal.ReleaseComObject(app);}}
    }
}
