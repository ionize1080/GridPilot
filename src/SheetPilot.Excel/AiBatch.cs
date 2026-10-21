using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private sealed class Edit
        {
            public string address,formula;
            public object value,beforeValue,beforeFormula,afterFormula;
            public bool hadFormula,afterHasFormula;
        }
        private static bool Same(object a,object b) => JToken.DeepEquals(a==null?JValue.CreateNull():JToken.FromObject(a),b==null?JValue.CreateNull():JToken.FromObject(b));
        private object Worksheet(Snapshot snapshot,string name)
        {
            if(!snapshot.sheets.Contains(name))throw new InvalidOperationException("目标表不在当前上下文中，请重新选择范围。");
            object sheets=((dynamic)snapshot.book).Worksheets;
            try{return ((dynamic)sheets).Item[name];}finally{Release(sheets);}
        }
        public object Metadata(string id)
        {
            var snapshot=Get(id);var result=new List<object>();
            foreach(string name in snapshot.sheets)
            {
                object sheet=Worksheet(snapshot,name),tables=null,charts=null,pivots=null;
                try
                {
                    tables=((dynamic)sheet).ListObjects;charts=((dynamic)sheet).ChartObjects();pivots=((dynamic)sheet).PivotTables();
                    var ts=new List<object>();
                    for(int i=1;i<=(int)((dynamic)tables).Count;i++)
                    {
                        object t=((dynamic)tables).Item[i],range=null;
                        try{range=((dynamic)t).Range;ts.Add(new{name=(string)((dynamic)t).Name,address=(string)((dynamic)range).Address[false,false]});}
                        finally{Release(range);Release(t);}
                    }
                    result.Add(new{sheet=name,protectedContents=(bool)((dynamic)sheet).ProtectContents,tables=ts,chartCount=(int)((dynamic)charts).Count,pivotCount=(int)((dynamic)pivots).Count});
                }finally{Release(pivots);Release(charts);Release(tables);Release(sheet);}
            }
            return new{snapshotId=id,book=snapshot.bookName,sheets=result,ranges=snapshot.ranges};
        }
        public object Propose(JObject request)
        {
            var snapshot=Get((string)request["snapshotId"]);
            string name=(string)request["sheet"],address=(string)request["address"],kind=(string)request["kind"]??"formula";
            var bounds=AiRange.Parse(address);object sheet=Worksheet(snapshot,name),range=null;
            var p=new Proposal{id=Guid.NewGuid().ToString("N"),sequence=++proposalSequence,snapshot=snapshot.id,sheet=name,address=address,kind=kind,options=(JObject)request.DeepClone()};
            try
            {
                range=((dynamic)sheet).Range[address];
                CheckWritable(snapshot,sheet,range,kind=="chart"||kind=="pivot");
                if(kind=="formula"||kind=="values"||kind=="formulas")
                {
                    var data=request["data"] as JArray;
                    if(kind=="formula") { ValidateAddress(address);ValidateFormula((string)request["formula"]);data=new JArray{new JArray{(string)request["formula"]}}; }
                    if(data==null||data.Count!=bounds.LastRow-bounds.Row+1||data.Any(row=>!(row is JArray)||row.Count()!=bounds.LastColumn-bounds.Column+1))
                        throw new InvalidOperationException("二维数据的行列数必须与目标范围一致。");
                    var edits=new List<Edit>();
                    for(int r=0;r<data.Count;r++)for(int c=0;c<data[r].Count();c++)
                    {
                        var token=data[r][c];object value=(token as JValue)?.Value;
                        if(!(token is JValue)||value is string text&&text.Length>32767)throw new InvalidOperationException("内容必须是文本、数字、布尔值或空值；文本最多32767字符。");
                        string formula=kind=="values"?null:(string)token;
                        if(kind!="values")ValidateFormula(formula);
                        object cell=null;
                        try
                        {
                            string dest=Address(bounds.Row+r,bounds.Column+c);cell=((dynamic)sheet).Range[dest];
                            edits.Add(new Edit{address=dest,value=value,formula=formula,beforeValue=((dynamic)cell).Value2,beforeFormula=((dynamic)cell).Formula,hadFormula=(bool)((dynamic)cell).HasFormula});
                        }finally{Release(cell);}
                    }
                    p.edits=edits.ToArray();p.formula=(string)request["formula"];
                }
                else if(kind=="format")
                {
                    p.style=request["style"] as JObject??throw new InvalidOperationException("缺少格式。");ValidateStyle(p.style);p.beforeStyle=ReadStyle(range,p.style);
                    if(p.beforeStyle.Properties().Any(v=>v.Value.Type==JTokenType.Null))throw new InvalidOperationException("目标存在混合格式，请拆分成格式一致的范围后再提出建议。");
                }
                else if(kind=="pivot")
                {
                    if((bool)((dynamic)snapshot.book).ProtectStructure)throw new InvalidOperationException("工作簿结构受保护，不能创建透视表工作表。");
                    ValidatePivot(p,sheet);p.sourceFormula=((dynamic)range).Formula;
                }
                else if(kind=="chart")
                {
                    string type=(string)request["chartType"];
                    if(!ChartTypes.ContainsKey(type??""))throw new InvalidOperationException("支持 column、bar、line、pie、scatter、area 图表。");
                    if(((string)request["title"]??"").Length>255)throw new InvalidOperationException("图表标题过长。");
                    p.sourceFormula=((dynamic)range).Formula;
                }
                else throw new InvalidOperationException("未知的建议类型。");
                proposals.Add(p.id,p);
                bool applied = AutomaticFor(snapshot.book);
                if (applied) { try { ApplyBatch(new[] { p.id }); } catch { if(CurrentApproval()?.scope!="workbook")automatic=false; throw; } }
                return new{proposalId=p.id,snapshotId=p.snapshot,book=snapshot.bookName,sheet=name,address,kind,formula=p.formula,data=request["data"],style=p.style,title=request["title"],before=p.edits?.Take(6).Select(e=>e.beforeFormula).ToArray(),count=p.edits?.LongLength??bounds.Size,applied,checkpoint=CheckpointState(),status=applied?"已自动应用；已保留变更前后快照。":"待一键应用，尚未写入 Excel。"};
            }finally{Release(range);Release(sheet);}
        }
        private static readonly Dictionary<string,int> ChartTypes=new Dictionary<string,int>{{"column",51},{"bar",57},{"line",4},{"pie",5},{"scatter",-4169},{"area",1}};
        private static readonly string[] StyleKeys={"bold","italic","font_size","font_color","fill_color","number_format","horizontal_alignment","wrap_text"};
        private static void ValidateStyle(JObject style)
        {
            if(!style.HasValues)throw new InvalidOperationException("格式不能为空。");
            foreach(var field in style.Properties())
            {
                if(!StyleKeys.Contains(field.Name))throw new InvalidOperationException("不支持的格式："+field.Name);
                if(field.Name.EndsWith("color"))AiRange.Color((string)field.Value);
                else if(field.Name=="font_size"&&((double)field.Value<=0||(double)field.Value>409))throw new InvalidOperationException("字号需在0至409之间。");
                else if(field.Name=="horizontal_alignment"&&!new[]{"general","left","center","right"}.Contains((string)field.Value))throw new InvalidOperationException("不支持的对齐方式。");
                else if(new[]{"bold","italic","wrap_text"}.Contains(field.Name)&&field.Value.Type!=JTokenType.Boolean)throw new InvalidOperationException("格式开关需为布尔值。");
                else if(field.Name=="number_format"&&((string)field.Value)?.Length>255)throw new InvalidOperationException("数字格式过长。");
            }
        }
        private static JObject ReadStyle(dynamic range,JObject keys)
        {
            object font=range.Font,interior=range.Interior;var result=new JObject();
            try
            {
                foreach(var key in keys.Properties())
                {
                    object v=null;
                    switch(key.Name){case "bold":v=((dynamic)font).Bold;break;case "italic":v=((dynamic)font).Italic;break;case "font_size":v=((dynamic)font).Size;break;case "font_color":v=((dynamic)font).Color;break;case "fill_color":v=((dynamic)interior).Color;break;case "number_format":v=range.NumberFormat;break;case "horizontal_alignment":v=range.HorizontalAlignment;break;case "wrap_text":v=range.WrapText;break;case "fill_pattern":v=((dynamic)interior).Pattern;break;}
                    result[key.Name]=v==null||v is DBNull?JValue.CreateNull():JToken.FromObject(v);
                }
                if(keys["fill_color"]!=null)result["fill_pattern"]=JToken.FromObject(((dynamic)interior).Pattern);
                return result;
            }finally{Release(interior);Release(font);}
        }
        private static void WriteStyle(dynamic range,JObject style,bool raw)
        {
            object font=range.Font,interior=range.Interior;
            try
            {
                foreach(var f in style.Properties())
                {
                    switch(f.Name)
                    {
                        case "bold":((dynamic)font).Bold=(bool)f.Value;break;case "italic":((dynamic)font).Italic=(bool)f.Value;break;
                        case "font_size":((dynamic)font).Size=(double)f.Value;break;
                        case "font_color":((dynamic)font).Color=raw?(int)f.Value:AiRange.Color((string)f.Value);break;
                        case "fill_color":((dynamic)interior).Color=raw?(int)f.Value:AiRange.Color((string)f.Value);break;
                        case "fill_pattern":((dynamic)interior).Pattern=(int)f.Value;break;
                        case "number_format":range.NumberFormat=(string)f.Value;break;case "wrap_text":range.WrapText=(bool)f.Value;break;
                        case "horizontal_alignment":range.HorizontalAlignment=raw?(int)f.Value:new Dictionary<string,int>{{"general",1},{"left",-4131},{"center",-4108},{"right",-4152}}[(string)f.Value];break;
                    }
                }
            }finally{Release(interior);Release(font);}
        }
        private static JObject ChartState(dynamic obj)
        {
            object chart=obj.Chart,series=null,title=null;var formulas=new JArray();
            try
            {
                dynamic c=chart;series=c.SeriesCollection();
                for(int i=1;i<=(int)((dynamic)series).Count;i++){object item=((dynamic)series).Item(i);try{formulas.Add((string)((dynamic)item).Formula);}finally{Release(item);}}
                string text=null;if((bool)c.HasTitle){title=c.ChartTitle;text=((dynamic)title).Text;}
                return new JObject{["name"]=(string)obj.Name,["left"]=(double)obj.Left,["top"]=(double)obj.Top,["width"]=(double)obj.Width,["height"]=(double)obj.Height,["type"]=(int)c.ChartType,["style"]=JToken.FromObject(c.ChartStyle),["legend"]=(bool)c.HasLegend,["title"]=text,["series"]=formulas};
            }finally{Release(title);Release(series);Release(chart);}
        }
        private static void CheckWritable(Snapshot snapshot,dynamic sheet,dynamic range,bool chart=false)
        {
            if((bool)((dynamic)snapshot.book).ReadOnly||(bool)sheet.ProtectContents||(int)sheet.Visible!=-1)throw new InvalidOperationException("目标只读、受保护或已隐藏。");
            if(chart)return;
            object spill=false;
            try{spill=range.HasSpill;}catch(System.Runtime.InteropServices.COMException){}catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}
            if(!(spill is bool)||(bool)spill)throw new InvalidOperationException("目标包含动态数组溢出区域，请选择普通单元格。");
            object tables=null;
            try
            {
                object merged=range.MergeCells,array=range.HasArray;
                if(!(merged is bool)||!(array is bool)||(bool)merged||(bool)array)throw new InvalidOperationException("目标含合并单元格或数组公式，请选择普通区域。");
                // Test every intersecting table, including a range spanning more than one table.
                tables=sheet.ListObjects;
                for(int i=1;i<=(int)((dynamic)tables).Count;i++)
                {
                    object table=((dynamic)tables).Item[i],tr=null,hit=null,app=null;
                    try{tr=((dynamic)table).Range;app=sheet.Application;hit=((dynamic)app).Intersect(range,tr);if(hit!=null)throw new InvalidOperationException("目标与 Excel 表格相交，自动扩展可能影响其他行；请选表格外区域。");}
                    finally{Release(app);Release(hit);Release(tr);Release(table);}
                }
            }finally{Release(tables);}
        }
        private void CheckProposal(Proposal p,bool revert)
        {
            var snapshot=Get(p.snapshot);object sheet=Worksheet(snapshot,p.sheet),range=null;
            try
            {
                range=((dynamic)sheet).Range[p.address];CheckWritable(snapshot,sheet,range,p.kind=="chart"||p.kind=="pivot");
                if(p.edits!=null)foreach(var e in p.edits)
                {
                    object cell=((dynamic)sheet).Range[e.address];
                    try{if(!Same(((dynamic)cell).Formula,revert?e.afterFormula:e.beforeFormula)||(bool)((dynamic)cell).HasFormula!=(revert?e.afterHasFormula:e.hadFormula))throw new InvalidOperationException("目标内容已变化，整批未修改，请重新生成建议。");}
                    finally{Release(cell);}
                }
                else if(p.kind=="format")
                {if(!JToken.DeepEquals(ReadStyle(range,p.beforeStyle),revert?p.afterStyle:p.beforeStyle))throw new InvalidOperationException("目标格式已变化，整批未修改。");}
                else if(!Same(((dynamic)range).Formula,revert?p.afterSourceFormula??p.sourceFormula:p.sourceFormula))throw new InvalidOperationException("图表源数据已变化，请重新生成建议。");
                if(p.kind=="pivot")
                {
                    if((bool)((dynamic)snapshot.book).ProtectStructure)throw new InvalidOperationException("工作簿结构受保护。");
                    EnsurePivotDestination(p,revert);
                }
                if(revert&&p.kind=="chart")
                {
                    object charts=((dynamic)sheet).ChartObjects(),chart=null;
                    try{chart=((dynamic)charts).Item[p.chartName];if(!JToken.DeepEquals(ChartState(chart),p.chartState))throw new InvalidOperationException("图表已被修改，未撤销，请保留当前图表。");}
                    finally{Release(chart);Release(charts);}
                }
            }finally{Release(range);Release(sheet);}
        }
        private static void SetValue(dynamic cell,object value)
        {
            if(value is string text){object format=cell.NumberFormat;try{cell.NumberFormat="@";cell.Value2=text;}finally{cell.NumberFormat=format;}}
            else cell.Value2=value;
        }
        private void WriteProposal(Proposal p,bool revert)
        {
            var snapshot=Get(p.snapshot);object sheet=Worksheet(snapshot,p.sheet),range=null;
            try
            {
                range=((dynamic)sheet).Range[p.address];
                if(p.kind=="pivot"){WritePivot(p,range,revert);return;}
                if(p.edits!=null)foreach(var e in p.edits)
                {
                    object cell=((dynamic)sheet).Range[e.address];
                    try
                    {
                        if(revert){if(e.hadFormula)((dynamic)cell).Formula=e.beforeFormula;else SetValue(cell,e.beforeValue);}
                        else{if(e.formula!=null)((dynamic)cell).Formula=e.formula;else SetValue(cell,e.value);e.afterFormula=((dynamic)cell).Formula;e.afterHasFormula=(bool)((dynamic)cell).HasFormula;}
                    }finally{Release(cell);}
                }
                else if(p.kind=="format")
                {WriteStyle(range,revert?p.beforeStyle:p.style,revert);if(!revert)p.afterStyle=ReadStyle(range,p.beforeStyle);}
                else
                {
                    object charts=((dynamic)sheet).ChartObjects(),obj=null,chart=null;
                    try
                    {
                        if(revert){if(p.chartName==null)return;obj=((dynamic)charts).Item[p.chartName];((dynamic)obj).Delete();}
                        else
                        {
                            obj=((dynamic)charts).Add((double)((dynamic)range).Left+30,(double)((dynamic)range).Top+30,480,280);
                            p.chartName="GridPilot_AI_"+p.id;((dynamic)obj).Name=p.chartName;chart=((dynamic)obj).Chart;
                            ((dynamic)chart).SetSourceData(range);((dynamic)chart).ChartType=ChartTypes[(string)p.options["chartType"]];
                            string title=(string)p.options["title"];
                            if(!string.IsNullOrWhiteSpace(title)){((dynamic)chart).HasTitle=true;object ct=((dynamic)chart).ChartTitle;try{((dynamic)ct).Text=title;}finally{Release(ct);}}
                            p.chartState=ChartState(obj);
                        }
                    }finally{Release(chart);Release(obj);Release(charts);}
                }
            }finally{Release(range);Release(sheet);}
        }
        public object Apply(string id,bool revert=false) => revert?ApplyBatch(null,true):ApplyBatch(new[]{id});
        public object Discard(string[] ids) { foreach(var id in ids??new string[0])proposals.Remove(id);return new {message="已清空所选待应用建议。"}; }
        private Proposal[] ResolveBatch(Proposal[] source)
        {
            var result=new List<Proposal>();
            var seen=new HashSet<Tuple<object,string,string,string>>();
            foreach(var p in source.OrderByDescending(p=>p.sequence))
            {
                object book=Get(p.snapshot).book;
                if(p.edits!=null){var copy=p.Copy();copy.edits=p.edits.Where(e=>seen.Add(Tuple.Create(book,p.sheet,e.address,"content"))).ToArray();if(copy.edits.Length>0)result.Add(copy);}
                else if(p.kind=="format")
                {
                    var area=AiRange.Parse(p.address);
                    if(!source.Any(other=>other!=p&&other.kind=="format"&&other.sheet==p.sheet&&ReferenceEquals(Get(other.snapshot).book,book)&&RangesOverlap(area,AiRange.Parse(other.address)))){result.Add(p);continue;}
                    var bounds=AiRange.Parse(p.address);if(bounds.Size>100000)throw new InvalidOperationException("单次格式批量应用最多 100,000 格，请缩小范围。");
                    for(int r=bounds.Row;r<=bounds.LastRow;r++)for(int c=bounds.Column;c<=bounds.LastColumn;c++)
                    {
                        var copy=p.Copy();copy.address=Address(r,c);copy.style=new JObject();copy.beforeStyle=new JObject();
                        foreach(var f in p.style.Properties())if(seen.Add(Tuple.Create(book,p.sheet,copy.address,f.Name))){copy.style[f.Name]=f.Value.DeepClone();copy.beforeStyle[f.Name]=p.beforeStyle[f.Name].DeepClone();}
                        if(copy.style["fill_color"]!=null)copy.beforeStyle["fill_pattern"]=p.beforeStyle["fill_pattern"].DeepClone();
                        if(copy.style.HasValues)result.Add(copy);
                    }
                }
                else result.Add(p);
            }
            return result.OrderBy(p=>p.edits!=null?0:p.kind=="format"?1:2).ThenBy(p=>p.sequence).ToArray();
        }
        private static bool RangesOverlap(AiRange a,AiRange b) => a.Row<=b.LastRow&&b.Row<=a.LastRow&&a.Column<=b.LastColumn&&b.Column<=a.LastColumn;
        public object ApplyBatch(string[] ids,bool revert=false)
        {
            if(revert)return RestoreCheckpoint(true);
            Proposal[] batch;
                if(ids==null||ids.Length==0||ids.Distinct().Count()!=ids.Length)throw new InvalidOperationException("请选择待应用的建议。");
                var valid=ids.Where(id=>id!=null&&proposals.ContainsKey(id)).Select(id=>proposals[id]).ToArray();
                if(valid.Length==0)return new{message="待应用建议已过期，已清理；请重新生成。",appliedIds=new string[0],retiredIds=ids,checkpoint=CheckpointState()};
                batch=ResolveBatch(valid);
            foreach(var p in batch)CheckProposal(p,revert);
            bool batchAutomatic=batch.All(p=>AutomaticFor(Get(p.snapshot).book));
            CheckCheckpointExtension(batch,batchAutomatic);
            bool extend = batchAutomatic && !newCheckpoint && undo != null;
            string folder = extend ? checkpointFolder : SaveCheckpoint(batch, "before");
            Proposal[] candidate=null;
            dynamic app=application();bool events=app.EnableEvents,screen=app.ScreenUpdating;var written=new List<Proposal>();
            try
            {
                app.EnableEvents=false;app.ScreenUpdating=false;
                foreach(var p in revert?batch.Reverse():batch){written.Add(p);WriteProposal(p,revert);}
                candidate=CombineCheckpoint(extend?undo.Concat(batch).ToArray():batch);
                CompleteCheckpointSources(candidate);
                SaveCheckpoint(batch, "after-" + Guid.NewGuid().ToString("N"), folder);
            }
            catch(Exception ex)
            {
                var failures=new List<string>();
                foreach(var p in written.AsEnumerable().Reverse())try{WriteProposal(p,!revert);}catch(Exception restore){failures.Add(restore.Message);}
                throw new InvalidOperationException(failures.Count==0?"操作失败，已恢复整批修改："+ex.Message:"操作失败且恢复未完成，请检查工作簿："+ex.Message+"；"+string.Join("；",failures));
            }
            finally{app.EnableEvents=events;app.ScreenUpdating=screen;}
            undo=candidate;
            checkpointFolder=folder;checkpointBefore=false;newCheckpoint=false;
            foreach(var id in ids)proposals.Remove(id);
            foreach(var key in snapshots.Keys.Where(key=>!snapshotOrder.Contains(key)&&!undo.Any(p=>p.snapshot==key)).ToArray()){snapshots[key].Dispose();snapshots.Remove(key);}
            return new{message="已应用 "+valid.Length+" 项建议；重叠内容采用最新建议，前后快照已保存。",appliedIds=valid.Select(p=>p.id).ToArray(),retiredIds=ids,checkpoint=CheckpointState()};
        }
    }
}
