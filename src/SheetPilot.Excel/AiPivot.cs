using System;
using System.Linq;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private static readonly Dictionary<string,int> PivotFunctions=new Dictionary<string,int>{{"sum",-4157},{"count",-4112},{"average",-4106},{"min",-4139},{"max",-4136}};
        private void ValidatePivot(Proposal p, dynamic source)
        {
            string target=(string)p.options["destinationSheet"];
            if(string.IsNullOrWhiteSpace(target)||target.Length>31||target.IndexOfAny(new[]{'[',']',':','*','?','/','\\'})>=0||target.StartsWith("'")||target.EndsWith("'"))
                throw new InvalidOperationException("请给数据透视表指定合法的新工作表名称，最多31字符。");
            var bounds=AiRange.Parse(p.address);if(bounds.LastRow<=bounds.Row)throw new InvalidOperationException("透视表源需含标题及至少一行数据。");
            object header=null;
            try
            {
                header=source.Range[Address(bounds.Row,bounds.Column)+":"+Address(bounds.Row,bounds.LastColumn)];object data=((dynamic)header).Value2;
                var names=Enumerable.Range(0,bounds.LastColumn-bounds.Column+1).Select(c=>Convert.ToString(At(data,0,c))).ToArray();
                if(names.Any(string.IsNullOrWhiteSpace)||names.Distinct(StringComparer.OrdinalIgnoreCase).Count()!=names.Length)throw new InvalidOperationException("源数据首行必须是非空且不重复的字段名。");
                var rows=p.options["rows"] as JArray;var columns=p.options["columns"] as JArray??new JArray();var values=p.options["values"] as JArray;
                if(rows==null||rows.Count==0||values==null||values.Count==0)throw new InvalidOperationException("至少指定一个行字段和一个汇总字段。");
                var dimensions=rows.Concat(columns).Select(x=>(string)x).ToArray();
                if(dimensions.Distinct().Count()!=dimensions.Length||dimensions.Any(x=>!names.Contains(x)))throw new InvalidOperationException("行列字段必须来自源标题且不能重复。");
                foreach(var value in values)if(!names.Contains((string)value["field"])||!PivotFunctions.ContainsKey((string)value["function"]??"sum"))throw new InvalidOperationException("汇总字段必须存在；支持 sum、count、average、min、max。");
                EnsurePivotDestination(p,false);
            }finally{Release(header);}
        }
        private object PivotSheet(Proposal p)
        {
            object sheets=((dynamic)Get(p.snapshot).book).Worksheets;
            try { for(int i=1;i<=(int)((dynamic)sheets).Count;i++){object sheet=((dynamic)sheets).Item[i];if(string.Equals((string)((dynamic)sheet).Name,(string)p.options["destinationSheet"],StringComparison.OrdinalIgnoreCase))return sheet;Release(sheet);}return null; }
            finally{Release(sheets);}
        }
        private void EnsurePivotDestination(Proposal p,bool exists)
        {
            object sheet=PivotSheet(p);
            try
            {
                if(!exists&&sheet!=null)throw new InvalidOperationException("数据透视表目标工作表已存在，请使用新的名称。");
                if(exists&&(sheet==null||!JToken.DeepEquals(PivotState(sheet),p.chartState)))throw new InvalidOperationException("透视表或其工作表已被修改，未覆盖手工编辑；可使用本地完整备份。");
            }finally{Release(sheet);}
        }
        private static JObject PivotState(dynamic sheet)
        {
            object range=null,rows=null,columns=null,shapes=null,tables=null,formats=null,links=null;
            try
            {
                range=sheet.UsedRange;rows=((dynamic)range).Rows;columns=((dynamic)range).Columns;shapes=sheet.Shapes;tables=sheet.PivotTables();formats=((dynamic)range).FormatConditions;links=sheet.Hyperlinks;
                int nr=((dynamic)rows).Count,nc=((dynamic)columns).Count;
                if((long)nr*nc>100000)throw new InvalidOperationException("透视表结果超过10万格，无法提供安全的一键切换，请缩小透视维度。");
                var data=new StringBuilder(JsonConvert.SerializeObject(new{formula=(object)((dynamic)range).Formula,address=(string)((dynamic)range).Address[false,false],visible=(int)sheet.Visible,protect=(bool)sheet.ProtectContents,shapes=(int)((dynamic)shapes).Count,tables=(int)((dynamic)tables).Count,formats=(int)((dynamic)formats).Count,links=(int)((dynamic)links).Count}));
                for(int r=1;r<=nr;r++)for(int c=1;c<=nc;c++)
                {
                    object cell=((dynamic)range).Cells[r,c],font=null,fill=null,borders=null;
                    try
                    {
                        dynamic x=cell;font=x.Font;fill=x.Interior;borders=x.Borders;
                        data.Append(JsonConvert.SerializeObject(new{style=(object)x.Style,format=(object)x.NumberFormat,align=(object)x.HorizontalAlignment,vertical=(object)x.VerticalAlignment,wrap=(object)x.WrapText,merged=(object)x.MergeCells,height=(object)x.Height,width=(object)x.Width,font=(object)((dynamic)font).Name,size=(object)((dynamic)font).Size,bold=(object)((dynamic)font).Bold,italic=(object)((dynamic)font).Italic,color=(object)((dynamic)font).Color,fill=(object)((dynamic)fill).Color,pattern=(object)((dynamic)fill).Pattern}));
                        foreach(int index in new[]{5,6,7,8,9,10}){object b=((dynamic)borders).Item[index];try{data.Append(JsonConvert.SerializeObject(new{style=(object)((dynamic)b).LineStyle,color=(object)((dynamic)b).Color,weight=(object)((dynamic)b).Weight}));}finally{Release(b);}}
                    }finally{Release(borders);Release(fill);Release(font);Release(cell);}
                }

                for(int i=1;i<=(int)((dynamic)tables).Count;i++)
                {
                    object pivot=((dynamic)tables).Item[i],fields=null,cache=null;
                    try
                    {
                        cache=((dynamic)pivot).PivotCache();fields=((dynamic)pivot).PivotFields();
                        data.Append(JsonConvert.SerializeObject(new{name=(string)((dynamic)pivot).Name,source=(object)((dynamic)cache).SourceData,rowGrand=(bool)((dynamic)pivot).RowGrand,columnGrand=(bool)((dynamic)pivot).ColumnGrand}));
                        for(int f=1;f<=(int)((dynamic)fields).Count;f++)
                        {
                            object field=((dynamic)fields).Item[f];
                            try {data.Append(JsonConvert.SerializeObject(new{name=(string)((dynamic)field).Name,orientation=(int)((dynamic)field).Orientation}));}
                            finally{Release(field);}
                        }
                    }finally{Release(cache);Release(fields);Release(pivot);}
                }
                using(var sha=SHA256.Create())return new JObject{["hash"]=Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(data.ToString())))};
            }finally{Release(links);Release(formats);Release(tables);Release(shapes);Release(columns);Release(rows);Release(range);}
        }
        private void WritePivot(Proposal p,dynamic sourceRange,bool revert)
        {
            object target=null,sheets=null,cache=null,caches=null,pivot=null,destination=null,last=null;
            dynamic app=application();bool alerts=app.DisplayAlerts;
            try
            {
                if(revert){target=PivotSheet(p);if(target!=null){app.DisplayAlerts=false;((dynamic)target).Delete();}return;}
                dynamic book=Get(p.snapshot).book;
                sheets=book.Worksheets;last=((dynamic)sheets).Item[((dynamic)sheets).Count];target=((dynamic)sheets).Add(After:last);((dynamic)target).Name=(string)p.options["destinationSheet"];
                caches=book.PivotCaches();string source=sourceRange.Address[true,true,1,true];cache=((dynamic)caches).Create(1,source,3);
                destination=((dynamic)target).Range["A3"];pivot=((dynamic)cache).CreatePivotTable(destination,"GridPilot_Pivot_"+p.id);
                int row=0,column=0;
                foreach(var name in (JArray)p.options["rows"]){object field=((dynamic)pivot).PivotFields((string)name);try{((dynamic)field).Orientation=1;((dynamic)field).Position=++row;}finally{Release(field);}}
                foreach(var name in (JArray)p.options["columns"]??new JArray()){object field=((dynamic)pivot).PivotFields((string)name);try{((dynamic)field).Orientation=2;((dynamic)field).Position=++column;}finally{Release(field);}}
                int index=0;
                foreach(var value in (JArray)p.options["values"]){string name=(string)value["field"],fn=(string)value["function"]??"sum";object field=((dynamic)pivot).PivotFields(name),added=null;try{added=((dynamic)pivot).AddDataField(field,fn+" · "+name+" "+(++index),PivotFunctions[fn]);}finally{Release(added);Release(field);}}
                ((dynamic)pivot).RefreshTable();p.chartState=PivotState(target);
            }
            catch { if(!revert&&target!=null){app.DisplayAlerts=false;((dynamic)target).Delete();}throw; }
            finally{app.DisplayAlerts=alerts;Release(last);Release(destination);Release(pivot);Release(cache);Release(caches);Release(sheets);Release(target);}
        }
    }
}
