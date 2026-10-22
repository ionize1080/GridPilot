using System;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private static readonly Dictionary<string,string> ModernFeatures=new Dictionary<string,string> {
            {"range.formula2","dynamicArrays"},{"range.spill","dynamicArrays"},{"range.calculate","base"},
            {"comment.list","threadedComments"},{"comment.add","threadedComments"},{"comment.reply","threadedComments"},{"comment.delete","threadedComments"},
            {"query.list","queries"},{"query.set","queries"},{"query.delete","queries"},
            {"connection.list","base"},{"model.inspect","dataModel"},{"slicer.list","base"},{"slicer.filter","base"},{"slicer.clear","base"},{"slicer.timeline","timeline"},{"table.resize","base"}
        };
        public static readonly Dictionary<string,int> ModernCharts=new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase){
            {"xlComboColumnClusteredLine",113},{"xlComboColumnClusteredLineSecondaryAxis",114},{"xlComboAreaStackedColumnClustered",115},{"xlOtherCombinations",116},{"xlCombo",-4152},
            {"xlTreemap",117},{"xlHistogram",118},{"xlWaterfall",119},{"xlSunburst",120},{"xlBoxwhisker",121},{"xlPareto",122},{"xlFunnel",123},
            {"xlColumnClusteredEx",124},{"xlColumnStackedEx",125},{"xlColumnStacked100Ex",126},{"xlLineEx",127},{"xlLineStackedEx",128},{"xlLineStacked100Ex",129},{"xlPieEx",130},{"xlDoughnutEx",131},{"xlBarClusteredEx",132},{"xlBarStackedEx",133},{"xlBarStacked100Ex",134},{"xlAreaEx",135},{"xlAreaStackedEx",136},{"xlAreaStacked100Ex",137},{"xlXYScatterEx",138},{"xlBubbleEx",139},{"xlRegionMap",140}};
        private static bool Probe(object target,string property)
        {object value=null;try{value=ComGet(target,property);return true;}catch{return false;}finally{Release(value);}}
        private JObject Capabilities(Snapshot s)
        {
            dynamic app=application();int major=int.Parse(((string)app.Version).Split('.')[0]);object sheets=null,sheet=null,range=null;
            var features=new JObject{["base"]=true,["dynamicArrays"]=false,["threadedComments"]=false,["queries"]=major>=16&&Probe(s.book,"Queries"),["dataModel"]=major>=15&&Probe(s.book,"Model"),["timeline"]=major>=15,["modernCharts"]=major>=16,["comboCharts"]=major>=15};
            try{sheets=((dynamic)s.book).Worksheets;if((int)((dynamic)sheets).Count>0){sheet=((dynamic)sheets).Item[1];range=((dynamic)sheet).Range["A1"];features["dynamicArrays"]=major>=16&&Probe(range,"Formula2")&&Probe(range,"HasSpill");features["threadedComments"]=major>=16&&Probe(range,"CommentThreaded");}}
            finally{Release(range);Release(sheet);Release(sheets);}
            return new JObject{["excelVersion"]=(string)app.Version,["build"]=Convert.ToString(app.Build),["features"]=features,["enabledOperations"]=new JArray(ModernFeatures.Where(p=>(bool)features[p.Value]).Select(p=>p.Key)),["chartTypes"]=JObject.FromObject(Excel2010Charts.Concat(ModernCharts.Where(p=>major>=16||major>=15&&(p.Value<=116||p.Value<0))).ToDictionary(p=>p.Key,p=>p.Value)),["note"]="Version + non-mutating runtime member probes. Excel 16.0 does not distinguish 2016/2019/2021/2024/Microsoft 365. Chart subtypes depend on installed build/data shape; native Excel validates execution. Workbook protection, VBA trust and licensing remain authoritative."};
        }
        private static bool IsModernOperation(string op)=>op=="capability.list"||ModernFeatures.ContainsKey(op);
        private void CheckModernCapability(Snapshot s,string op,JObject o)
        {
            if(ModernFeatures.TryGetValue(op,out string feature)&&feature!="base"&&!(bool)Capabilities(s)["features"][feature])throw new InvalidOperationException("当前 Excel 不支持 "+op+"；请查询 capability.list。");
            if((op=="chart.create"||op=="chart.series")&&o["type"]!=null){int type=ChartType(o["type"]);int major=int.Parse(((string)((dynamic)application()).Version).Split('.')[0]);if(ModernCharts.Values.Contains(type)&&major<(type<=116||type<0?15:16))throw new InvalidOperationException("当前 Excel 版本不支持此图表类型。");}
        }
        private void CheckDiscoveryBoundary(Snapshot s,string op,JObject o)
        {
            if(s.scope!="auto"&&s.scope!="all")return;
            if(s.parent==null&&op!="capability.list")throw new InvalidOperationException("这是目录，请先 context.capture 或 context.bind 绑定目标工作簿。");
            if(new[]{"workbook.open","workbook.import","csv.import","image.insert"}.Contains(op)||op=="workbook.list"&&o["directory"]!=null)throw new InvalidOperationException("当前模式的读取范围限本次目录中已打开的工作簿及 VBA，不读取磁盘文件。");
        }
        private object ModernOperation(Snapshot s,string op,JObject o)
        {
            if(op=="capability.list")return Capabilities(s);
            dynamic book=s.book;var owned=new List<object>();
            Func<object,object> own=value=>{owned.Add(value);return value;};
            int offset=Math.Max(0,(int?)o["offset"]??0),limit=Math.Max(1,Math.Min(200,(int?)o["limit"]??100));
            try
            {
                if(op.StartsWith("query."))
                {
                    dynamic queries=own(book.Queries);
                    if(op=="query.list"){var items=new JArray();int total=queries.Count;for(int i=offset+1;i<=Math.Min(total,offset+limit);i++){dynamic q=own(queries.Item(i));items.Add(new JObject{["name"]=(string)q.Name,["formula"]=(string)q.Formula,["description"]=(string)q.Description});}return new{items,total,nextOffset=offset+items.Count<total?(int?)(offset+items.Count):null};}
                    string name=Required(o,"name");object query=null;for(int i=1;i<=(int)queries.Count;i++){dynamic q=own(queries.Item(i));if((string)q.Name==name)query=q;}
                    if(op=="query.delete"){if(query==null)throw new ArgumentException("查询不存在。");((dynamic)query).Delete(false);}
                    else{string formula=Required(o,"formula");if(query==null)query=own(queries.Add(name,formula));else ((dynamic)query).Formula=formula;if(o["description"]!=null)((dynamic)query).Description=(string)o["description"];}
                    return new{updated=true,refreshed=false,note="Query definition only; no refresh or source access."};
                }
                if(op=="connection.list")
                {dynamic connections=own(book.Connections);var items=new JArray();int total=connections.Count;for(int i=offset+1;i<=Math.Min(total,offset+limit);i++){dynamic c=own(connections.Item(i));items.Add(new JObject{["name"]=(string)c.Name,["type"]=(int)c.Type,["description"]=(string)c.Description});}return new{items,total,nextOffset=offset+items.Count<total?(int?)(offset+items.Count):null};}
                if(op=="model.inspect")
                {dynamic model=own(book.Model),tables=own(model.ModelTables),relations=own(model.ModelRelationships);var items=new JArray();int total=tables.Count;for(int i=offset+1;i<=Math.Min(total,offset+limit);i++){dynamic t=own(tables.Item(i));items.Add(new JObject{["name"]=(string)t.Name,["recordCount"]=Convert.ToInt64(t.RecordCount)});}return new{tables=items,total,relationshipCount=(int)relations.Count,nextOffset=offset+items.Count<total?(int?)(offset+items.Count):null};}
                if(op.StartsWith("slicer."))
                {
                    dynamic caches=own(book.SlicerCaches);
                    if(op=="slicer.list"){var items=new JArray();int total=caches.Count;for(int i=offset+1;i<=Math.Min(total,offset+limit);i++){dynamic c=own(caches.Item(i));items.Add(new JObject{["name"]=(string)c.Name,["sourceName"]=(string)c.SourceName,["olap"]=(bool)c.OLAP});}return new{items,total,nextOffset=offset+items.Count<total?(int?)(offset+items.Count):null};}
                    dynamic cache=own(caches.Item(Required(o,"cache")));
                    if(op=="slicer.clear")cache.ClearManualFilter();
                    else if(op=="slicer.timeline"){dynamic timeline=own(cache.TimelineState);timeline.SetFilterDateRange(DateTime.Parse(Required(o,"start"),System.Globalization.CultureInfo.InvariantCulture),DateTime.Parse(Required(o,"end"),System.Globalization.CultureInfo.InvariantCulture));}
                    else{var selected=((JArray)o["items"]).Values<string>().ToHashSetCompat();if(selected.Count==0)throw new ArgumentException("至少选择一个筛选项，或使用 slicer.clear。");if((bool)cache.OLAP)cache.VisibleSlicerItemsList=selected.ToArray();else{dynamic items=own(cache.SlicerItems);var all=new List<object>();for(int i=1;i<=(int)items.Count;i++)all.Add(own(items.Item(i)));if(selected.Any(n=>!all.Any(v=>(string)((dynamic)v).Name==n)))throw new ArgumentException("筛选项不存在。");foreach(dynamic item in all)if(selected.Contains((string)item.Name))item.Selected=true;foreach(dynamic item in all)if(!selected.Contains((string)item.Name))item.Selected=false;}}
                    return new{updated=true};
                }
                dynamic sheet=own(LiveSheet(s,o));
                if(op=="table.resize"){dynamic tables=own(sheet.ListObjects),table=own(tables.Item(Required(o,"name")));string address=Required(o,"address");AiRange.Parse(address);table.Resize(own(sheet.Range[address]));return new{updated=true};}
                if(op=="comment.list")
                {
                    dynamic comments=own(sheet.CommentsThreaded);var items=new JArray();int total=comments.Count;
                    for(int i=offset+1;i<=Math.Min(total,offset+limit);i++){dynamic c=own(comments.Item(i)),parent=own(c.Parent),replies=own(c.Replies);int count=replies.Count,replyOffset=Math.Max(0,(int?)o["replyOffset"]??0);var reply=new JArray();for(int j=replyOffset+1;j<=Math.Min(count,replyOffset+limit);j++){dynamic v=own(replies.Item(j));reply.Add((string)v.Text());}items.Add(new JObject{["address"]=(string)parent.Address[false,false],["text"]=(string)c.Text(),["replies"]=reply,["replyCount"]=count,["nextReplyOffset"]=replyOffset+reply.Count<count?(int?)(replyOffset+reply.Count):null});}
                    return new{items,total,nextOffset=offset+items.Count<total?(int?)(offset+items.Count):null};
                }
                string addr=Required(o,"address");AiRange.Parse(addr);dynamic range=own(sheet.Range[addr]);
                if(op=="range.spill"){object has=range.HasSpill;string spill=null;if(Equals(has,true)){dynamic parent=own(range.SpillParent),spilling=own(parent.SpillingToRange);spill=spilling.Address[false,false];}return new{hasSpill=has,address=spill,formula2=range.Formula2};}
                if(op=="range.formula2"){ValidateAddress(addr);string formula=Required(o,"formula");ValidateFormula(formula);range.Formula2=formula;return new{written=true,formula2=range.Formula2};}
                if(op=="range.calculate"){range.Calculate();return new{calculated=true};}
                if(op=="comment.add"){own(range.AddCommentThreaded(Required(o,"text")));return new{added=true};}
                dynamic comment=own(range.CommentThreaded);if(comment==null)throw new ArgumentException("单元格没有线程批注。");
                if(op=="comment.reply")own(comment.AddReply(Required(o,"text")));else if(op=="comment.delete")comment.Delete();else throw new ArgumentException("未知操作。");
                return new{updated=true};
            }
            finally{for(int i=owned.Count-1;i>=0;i--)Release(owned[i]);}
        }
    }
    internal static class ModernCollections { internal static HashSet<T> ToHashSetCompat<T>(this IEnumerable<T> items)=>new HashSet<T>(items); }
}
