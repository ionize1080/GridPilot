using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class ToolsModule
    {
        private void OpenSort() => OpenSortWorkspace(null);
        internal void OpenSheetSort(SheetItem source,bool? descending=null)
        {
            if(descending.HasValue)
            {
                var sheets=gateway.Capture(true).Sheets.Where(s=>s.WorkbookKey==source.WorkbookKey).ToArray();
                var option=settings.Sort.Clone();option.Natural=navigation.Settings.QuickSortMode==1;
                option.Alphabetical=navigation.Settings.QuickSortMode==0;option.Descending=descending.Value;
                Announce(gateway.SortSheets(sheets,option));navigation.Refresh();Ribbon.Refresh();return;
            }
            OpenSortWorkspace(source);
        }
        private void OpenSortWorkspace(SheetItem source)
        {
            SortSnapshot snapshot=null;SortPlan plan=null;SheetItem[] sheets=null;
            Func<string> restore=null;string planKey=null;bool applied=false,busy=false;
            var defaults=JsonConvert.DeserializeObject<MultiSortOptions>(JsonConvert.SerializeObject(settings.MultiSort??new MultiSortOptions()));
            if(source!=null)defaults=new MultiSortOptions{Levels=new List<SortLevel>{new SortLevel {Mode=settings.Sort.Natural?"natural":settings.Sort.Alphabetical?"values":"custom",Rules=settings.Sort.Rules,Sequences=settings.Sort.Sequences,Descending=settings.Sort.Descending}}};
            OpenPortable(async message=>
            {
                if(busy)throw new InvalidOperationException("请等待当前排序操作完成。");busy=true;
                try
                {
                    string action=(string)message["action"];
                    if(action=="sortNative"){gateway.OpenNativeSort();return new {message="已关闭 Excel 原生排序窗口，请重新读取数据。"};}
                    if(action=="init"||action=="sortCapture"||action=="sortSample")
                    {
                        snapshot=null;plan=null;planKey=null;applied=false;restore=null;
                        if(action=="sortSample"){gateway.InsertSortExample();source=null;}
                        string error=null;
                        if(source==null){try{snapshot=gateway.CaptureSort();defaults.Levels=defaults.Levels.Where(l=>l.Column<snapshot.ColumnCount).Take(20).ToList();if(defaults.Levels.Count==0)defaults.Levels.Add(new SortLevel());if(defaults.Levels.Count==1)defaults.Levels[0].Column=snapshot.InitialColumn;if(snapshot.IsTable)defaults.HasHeader=snapshot.TableHasHeader;}catch(InvalidOperationException ex){error=ex.Message;}}
                        else sheets=gateway.Capture(true).Sheets.Where(s=>s.WorkbookKey==source.WorkbookKey).ToArray();
                        return new {mode="sort",kind=source==null?"rows":"sheets",error,options=defaults,
                            description=source==null?snapshot?.Description:source.WorkbookName+" · 全部工作表（含隐藏页和图表页）",
                            headers=source==null?snapshot?.Headers:new[]{"工作表名称"},rowCount=source==null?snapshot?.RowCount:sheets.Length,
                            isTable=snapshot?.IsTable??false,merges=snapshot?.MergeCount??0,sequences=gateway.SortSequences()};
                    }
                    if(action=="sortUndo")
                    {
                        if(restore==null)throw new InvalidOperationException("本窗口没有可恢复的排序。");
                        var text=restore();restore=null;plan=null;planKey=null;applied=false;navigation.Refresh();Ribbon.Refresh();return new {message=text};
                    }
                    if(!new[]{"sortPreview","sortPage","sortApply"}.Contains(action))throw new InvalidOperationException("未知排序操作。");
                    if(source==null&&snapshot==null)throw new InvalidOperationException("请先读取数据，或插入示例。");
                    var options=message["options"]?.ToObject<MultiSortOptions>()??throw new InvalidOperationException("缺少排序选项。");
                    options.Validate(source==null?snapshot.ColumnCount:1);
                    if(source!=null&&options.Levels.Count!=1)throw new InvalidOperationException("工作表名称只需一个排序层级。");
                    string key=JsonConvert.SerializeObject(options);
                    if(action=="sortPreview")
                    {
                        if(applied)throw new InvalidOperationException("此预览已应用，请先重新读取。");
                        plan=null;planKey=null;
                        if(source==null)plan=gateway.PlanSort(snapshot,options);
                        else
                        {
                            var level=options.Levels[0];var so=level.Custom(options.MatchCase);so.Alphabetical=level.Mode=="values";
                            var order=await Task.Run(()=>CustomSort.Order(sheets.Select(s=>s.Name).ToArray(),so));
                            plan=new SortPlan{Order=order,Records=sheets.Length,Matched=new[]{sheets.Length}};
                        }
                        planKey=key;
                    }
                    if(plan==null||planKey!=key)throw new InvalidOperationException("规则已改变，请重新预览。");
                    if(action=="sortApply")
                    {
                        if(applied)throw new InvalidOperationException("已应用此预览，请重新读取。");
                        string result;var prior=gateway.UndoOperation;
                        if(source==null){result=gateway.ApplySort(snapshot,plan);settings.MultiSort=options;store.Save(settings);}
                        else
                        {
                            var level=options.Levels[0];var so=level.Custom(options.MatchCase);so.Alphabetical=level.Mode=="values";
                            result=gateway.SortSheets(sheets,so);settings.Sort=so;store.Save(settings);
                        }
                        restore=ReferenceEquals(prior,gateway.UndoOperation)?null:gateway.CaptureUndoAction();applied=true;
                        navigation.Refresh();Ribbon.Refresh();return new {message=result,canUndo=restore!=null};
                    }
                    int offset=(int?)message["offset"]??0,size=Math.Max(25,Math.Min(200,(int?)message["size"]??100));string search=(string)message["search"]??"";
                    if(source==null)return gateway.SortPage(snapshot,plan,offset,size,search);
                    var rows=plan.Order.Select((i,n)=>new {position=n+1,sourceRow=i+1,values=new[]{sheets[i].Name}}).Take(MultiSortOptions.PreviewLimit).Where(r=>r.values[0].IndexOf(search,StringComparison.OrdinalIgnoreCase)>=0).ToArray();
                    offset=Math.Max(0,Math.Min(offset,Math.Max(0,rows.Length-1)))/size*size;
                    return new {total=plan.Rows,records=plan.Rows,previewed=Math.Min(plan.Rows,MultiSortOptions.PreviewLimit),filtered=rows.Length,offset,size,matched=plan.Matched,headers=new[]{"工作表名称"},rows=rows.Skip(offset).Take(size)};
                }
                finally{busy=false;}
            });
        }
    }
    public sealed partial class ExcelGateway
    {
        public object[] SortSequences()
        {
            CheckThread();dynamic app=application();var result=new List<object>();
            for(int i=1;i<=(int)app.CustomListCount;i++)
            {
                // IDispatch can return a SAFEARRAY with a non-zero lower bound.
                // Reflection preserves that Array instead of a DLR object[] conversion.
                var raw=(Array)((object)app).GetType().InvokeMember("GetCustomListContents",System.Reflection.BindingFlags.InvokeMethod,null,(object)app,new object[]{i});
                var values=Enumerable.Range(0,raw.Length).Select(n=>Convert.ToString(raw.GetValue(n+raw.GetLowerBound(0)))).ToArray();
                result.Add(new {name="Excel · "+string.Join("、",values.Take(3))+"…",values});
            }
            return result.ToArray();
        }
    }
}
