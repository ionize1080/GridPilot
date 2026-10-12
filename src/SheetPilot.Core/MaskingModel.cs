using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;

namespace SheetPilot.Core
{
    public enum MaskKind { Keep, ChineseName, EnglishName, Identifier, Amount, Date }
    public sealed class MaskRule
    {
        public int Sheet { get; set; }
        public string SheetName { get; set; }
        public int Column { get; set; }
        public int StartRow { get; set; } = 2;
        public string Header { get; set; }
        public MaskKind Kind { get; set; }
        public string Domain { get; set; } = "";
        public decimal Factor { get; set; } = 1.7m;
        public int DayShift { get; set; } = 37;
    }
    public sealed class DataCell
    {
        public int Sheet, Row, Column;
        public object Value;
        public string Formula;
        public string NumberFormat;
        public string Key => Sheet+":"+Row+":"+Column;
        public string Location => "S"+(Sheet+1)+"!"+Address(Row,Column);
        public static string Address(int row,int column) { string s=""; for(int c=column;c>0;c=(c-1)/26) s=(char)('A'+(c-1)%26)+s; return s+row; }
        public DataCell Clone() => (DataCell)MemberwiseClone();
    }
    public sealed class DataSheet
    {
        public string Name;
        public int Rows, Columns;
    }
    public sealed class MaskSnapshot
    {
        public string WorkbookKey;
        public bool Date1904;
        public List<DataSheet> Sheets = new List<DataSheet>();
        public List<DataCell> Cells = new List<DataCell>();
        public MaskSnapshot Clone() => new MaskSnapshot { WorkbookKey=WorkbookKey, Date1904=Date1904, Sheets=Sheets,Cells=Cells.Select(c=>c.Clone()).ToList() };
    }
    public sealed class MaskResult
    {
        public MaskSnapshot Output;
        public int Changed, Formulas;
        public string Summary => "已验证 "+Changed+" 个常量替换、"+Formulas+" 条公式；输出仅含单元格数据与受支持公式。";
        public Dictionary<string,object> Calculated = new Dictionary<string,object>();
    }
    public static class MaskPlanner
    {
        public static MaskResult Build(MaskSnapshot source,IReadOnlyList<MaskRule> rules,CancellationToken cancel)
        {
            if(rules.GroupBy(r=>r.Sheet+":"+r.Column).Any(g=>g.Count()>1)) throw new InvalidOperationException("同一列存在重复规则。");
            var byColumn=rules.ToDictionary(r=>r.Sheet+":"+r.Column);
            Func<DataCell,MaskRule> ruleFor=c=>byColumn.TryGetValue(c.Sheet+":"+c.Column,out var r)&&c.Row>=r.StartRow?r:null;
            foreach(var rule in rules)
            {
                if(rule.StartRow<1 || rule.StartRow>1048576 || rule.Column<1 || rule.Column>16384) throw new InvalidOperationException("字段范围无效。");
                if(rule.Kind!=MaskKind.Keep && string.IsNullOrWhiteSpace(rule.Domain)) throw new InvalidOperationException("请为每个脱敏字段设置逻辑字段名；对应字段使用相同名称。");
                if(rule.Kind==MaskKind.Amount && (rule.Factor<=0 || rule.Factor==1 || rule.Factor>1000)) throw new InvalidOperationException("金额系数须在0到1000之间且不等于1。");
                if(rule.Kind==MaskKind.Date && (rule.DayShift==0 || Math.Abs((long)rule.DayShift)>36500)) throw new InvalidOperationException("日期平移须为非零整数且不超过36500天。");
            }
            foreach(var g in rules.Where(r=>r.Kind!=MaskKind.Keep).GroupBy(r=>r.Domain,StringComparer.Ordinal))
                if(g.Select(r=>r.Kind+"|"+r.Factor+"|"+r.DayShift).Distinct().Count()!=1) throw new InvalidOperationException("同一逻辑字段的规则不一致。");
            var originals=new HashSet<string>(source.Cells.Select(c=>c.Value as string).Where(s=>!string.IsNullOrEmpty(s)),StringComparer.OrdinalIgnoreCase);
            var mappings=new Dictionary<string,Dictionary<string,string>>(StringComparer.Ordinal);
            var literalMap=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            var sourceIndex=source.Cells.ToDictionary(c=>c.Key);
            var output=source.Clone(); var result=new MaskResult { Output=output };
            foreach(var cell in output.Cells)
            {
                cancel.ThrowIfCancellationRequested(); var rule=ruleFor(cell);
                if(rule?.Kind==MaskKind.Date)cell.NumberFormat="yyyy-mm-dd hh:mm:ss";
                if(cell.Formula!=null || rule==null || rule.Kind==MaskKind.Keep || cell.Value==null || Equals(cell.Value,"")) continue;
                object replacement;
                if(rule.Kind==MaskKind.Amount || rule.Kind==MaskKind.Date)
                {
                    if(!(cell.Value is double number) || double.IsNaN(number) || double.IsInfinity(number)) throw new InvalidOperationException(cell.Location+"：金额／日期列含非数值，请调整字段范围。");
                    decimal changed;
                    try { changed=rule.Kind==MaskKind.Amount?checked((decimal)number*rule.Factor):checked((decimal)number+rule.DayShift); }
                    catch(OverflowException) { throw new InvalidOperationException(cell.Location+"：数值超出支持精度。"); }
                    if(rule.Kind==MaskKind.Date && (number<(source.Date1904?0:1) || number>=(source.Date1904?2957004:2958466) || changed<(source.Date1904?0:1) || changed>=(source.Date1904?2957004:2958466) || (!source.Date1904 && (Math.Floor(number)==60 || Math.Floor((double)changed)==60)))) throw new InvalidOperationException(cell.Location+"：日期越界或涉及1900年虚构闰日。");
                    replacement=(double)changed;
                    if(rule.Kind==MaskKind.Amount && number!=0 && ((double)replacement==number || (double)replacement==0)) throw new InvalidOperationException(cell.Location+"：数值精度不足以安全缩放。");
                }
                else
                {
                    if(!(cell.Value is string original)) throw new InvalidOperationException(cell.Location+"：身份字段必须为文本，数值编号可能已经丢失精度。");
                    if(!mappings.TryGetValue(rule.Domain,out var map)) mappings.Add(rule.Domain,map=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase));
                    if(!map.TryGetValue(original,out var mapped))
                    {
                        mapped=NewName(rule.Kind,originals); originals.Add(mapped); map.Add(original,mapped);
                        if(literalMap.TryGetValue(original,out var other) && other!=mapped) throw new InvalidOperationException(cell.Location+"：同一原文被分配到不同逻辑字段，无法无歧义保持公式关系。");
                        literalMap[original]=mapped;
                    }
                    replacement=mapped;
                }
                cell.Value=replacement; if(!Equals(replacement,sourceIndex[cell.Key].Value)) result.Changed++;
            }
            if(result.Changed==0) throw new InvalidOperationException("没有需要脱敏的常量，请选择字段并检查起始行。");
            // Formula strings are tokenized: only complete string literals and sheet references change.
            foreach(var cell in output.Cells.Where(c=>c.Formula!=null))
                cell.Formula=SafeFormula.Rewrite(cell.Formula,source.Sheets,literalMap);
            var before=new FormulaWorkbook(source,false,cancel); var after=new FormulaWorkbook(output,true,cancel);
            foreach(var cell in source.Cells.Where(c=>c.Formula!=null))
            {
                cancel.ThrowIfCancellationRequested();
                var first=before.Evaluate(cell.Key); var second=after.Evaluate(cell.Key);
                if(!Equivalent(first.Value,cell.Value)) throw new InvalidOperationException(cell.Location+"：源公式缓存与验证计算不一致，请先重算，或该公式超出支持范围。");
                if(first.Trace!=second.Trace) throw new InvalidOperationException(cell.Location+"：脱敏改变了查找命中、条件分支或汇总记录，已停止导出。");
                object expected=cell.Value; var rule=ruleFor(cell);
                if(expected is string text && literalMap.TryGetValue(text,out var mapped)) expected=mapped;
                else if(expected is double d && rule!=null && rule.Kind==MaskKind.Amount) expected=(double)((decimal)d*rule.Factor);
                else if(expected is double date && rule!=null && rule.Kind==MaskKind.Date) expected=date+rule.DayShift;
                if(!Equivalent(expected,second.Value)) throw new InvalidOperationException(cell.Location+"：公式结果未满足保留／映射规则；请为相关计算列设置对应字段，或调整脱敏范围。");
                result.Calculated[cell.Key]=second.Value; result.Formulas++;
            }
            long residualWork=0;
            foreach(var cell in output.Cells)
            {
                cancel.ThrowIfCancellationRequested();
                var texts=cell.Formula==null?new[]{cell.Value as string}:SafeFormula.TextLiterals(cell.Formula).Concat(new[]{result.Calculated[cell.Key] as string});
                if(texts.Any(text=>text!=null && literalMap.Keys.Any(original=>
                {
                    if((++residualWork & 1023)==0)cancel.ThrowIfCancellationRequested();
                    if(residualWork>20000000)throw new InvalidOperationException("身份原文残留扫描超过本版计算上限，请缩小工作簿。");
                    return text.IndexOf(original,StringComparison.OrdinalIgnoreCase)>=0;
                })))
                    throw new InvalidOperationException(cell.Location+"：存在所选身份字段的原文残留，请纳入对应字段或移除不支持的文本引用。");
            }
            // Verify equality and strict order within every transformed numerical domain after double conversion.
            foreach(var group in source.Cells.Where(c=>c.Formula==null && ruleFor(c)!=null && (ruleFor(c).Kind==MaskKind.Amount || ruleFor(c).Kind==MaskKind.Date) && c.Value is double).GroupBy(c=>ruleFor(c).Domain))
            {
                var index=output.Cells.ToDictionary(c=>c.Key); var pairs=group.Select(c=>new { Before=(double)c.Value,After=(double)index[c.Key].Value }).OrderBy(p=>p.Before).ToArray();
                for(int i=1;i<pairs.Length;i++) if(pairs[i].Before==pairs[i-1].Before ? pairs[i].After!=pairs[i-1].After : pairs[i].After<=pairs[i-1].After) throw new InvalidOperationException("数值写回精度不能保持相等与排序关系。");
            }
            return result;
        }
        public static bool Equivalent(object a,object b)
        { if(a is double x && b is double y) return Math.Abs(x-y)<=1e-10*Math.Max(1,Math.Max(Math.Abs(x),Math.Abs(y))); return Equals(a,b) || a==null && Equals(b,0.0) || b==null && Equals(a,0.0); }
        private static string NewName(MaskKind kind,HashSet<string> used)
        {
            string[] chinese={"青禾","蓝杉","银枫","晨星","云帆","松泉","清岚","星河","海棠","远山","秋实","明川","春晓","云栖","月桂","锦溪"};
            string[] english={"Cedar","Maple","Willow","Birch","Orchid","River","Meadow","Haven","Aspen","Elm","Laurel","Brook","Ocean","Dawn","Pine","Valley"};
            using(var rng=RandomNumberGenerator.Create()) for(int attempts=0;attempts<10000;attempts++)
            {
                var bytes=new byte[8]; rng.GetBytes(bytes); uint id=BitConverter.ToUInt32(bytes,0)%1000000;
                string value=kind==MaskKind.ChineseName?chinese[bytes[4]%16]+chinese[bytes[5]%16]+id.ToString("D6",CultureInfo.InvariantCulture)+"有限公司":
                    kind==MaskKind.EnglishName?english[bytes[4]%16]+" "+english[bytes[5]%16]+" "+id.ToString("D6",CultureInfo.InvariantCulture)+" Ltd.":"TEST"+BitConverter.ToUInt32(bytes,0).ToString("X8")+BitConverter.ToUInt32(bytes,4).ToString("X8");
                if(!used.Contains(value)) return value;
            }
            throw new InvalidOperationException("无法生成无冲突替代值。");
        }
    }
}
