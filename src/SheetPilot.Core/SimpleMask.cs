using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace SheetPilot.Core
{
    public sealed class SimpleMaskOptions
    {
        public string Mode { get; set; } = "initials";
        public string Prefix { get; set; } = "";
        public string Suffix { get; set; } = "";
        public int EdgeCount { get; set; } = 2;
        public string Pattern { get; set; } = "^(.{2})([\\s\\S]*)(.{2})$";
        public string Replacement { get; set; } = "$3$2$1";
        public string Operator { get; set; } = "+";
        public string Operand { get; set; } = "100";
    }
    public sealed class SimpleMaskCell
    {
        public int Block { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public string Address { get; set; }
        public object Before { get; set; }
        public string After { get; set; }
        public string Skip { get; set; }
        public string WriteStatus { get; set; }
        public string NumberFormat { get; set; }
    }
    public sealed class SimpleMaskBlock
    {
        public string Address { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public int Rows { get; set; }
        public int Columns { get; set; }
    }
    public sealed class SimpleMaskSnapshot
    {
        public string BookKey { get; set; }
        public string SheetKey { get; set; }
        public string BookName { get; set; }
        public string SheetName { get; set; }
        public List<SimpleMaskBlock> Blocks { get; set; } = new List<SimpleMaskBlock>();
        public List<SimpleMaskCell> Cells { get; set; } = new List<SimpleMaskCell>();
    }
    public sealed class SimpleMaskResult
    {
        public SimpleMaskOptions Options { get; set; }
        public List<SimpleMaskCell> Cells { get; set; } = new List<SimpleMaskCell>();
        public int Changed => Cells.Count(c=>c.Skip==null && c.After!=SimpleMask.Text(c.Before));
        public int Skipped => Cells.Count(c=>c.Skip!=null);
        public string Summary => "读取 "+Cells.Count+" 格 · 变更 "+Changed+" 格 · 跳过 "+Skipped+" 格。复制和导出中的跳过项留空。";
    }
    public static class SimpleMask
    {
        public const string MappingVersion = "GP-SIMPLE-1 / Unihan-17.0.0";
        public const string Letters = "岚砚穹榕汐棠霁珀笙岑翎澜槐璟沐筠霜屿檀洵苒曜栖珩蔚桐";
        public const string Digits = "7391058264";
        private static readonly Lazy<Dictionary<int,char>> Initials = new Lazy<Dictionary<int,char>>(()=>
        {
            var map=new Dictionary<int,char>();
            using(var stream=typeof(SimpleMask).Assembly.GetManifestResourceStream("SheetPilot.Core.Resources.pinyin-initials-v1.tsv"))
            using(var reader=new StreamReader(stream))
            { string line;while((line=reader.ReadLine())!=null){var parts=line.Split('\t');map.Add(int.Parse(parts[0],NumberStyles.HexNumber,CultureInfo.InvariantCulture),parts[1][0]);} }
            return map;
        });
        public static string Text(object value) => Convert.ToString(value,CultureInfo.InvariantCulture)??"";
        public static void Validate(SimpleMaskOptions o)
        {
            if(o==null || !new[]{"initials","swap","reverse","regex","number"}.Contains(o.Mode))throw new InvalidOperationException("请选择有效的脱敏方式。");
            if((o.Prefix??"").Length>256 || (o.Suffix??"").Length>256)throw new InvalidOperationException("前后缀各最多 256 个字符。");
            if(o.Mode=="swap" && (o.EdgeCount<1||o.EdgeCount>1000))throw new InvalidOperationException("首尾交换长度应为 1–1000。");
            if(o.Mode=="regex")
            {
                if(string.IsNullOrEmpty(o.Pattern)||o.Pattern.Length>2048||(o.Replacement??"").Length>4096)throw new InvalidOperationException("请输入正则（最多 2048 字符），替换最多 4096 字符。");
                try{new Regex(o.Pattern,RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(50));}catch(ArgumentException){throw new InvalidOperationException("正则表达式无效，请检查括号与转义。");}
            }
            if(o.Mode=="number")
            {
                if(!new[]{"+","-","*","/"}.Contains(o.Operator)||!TryDecimal(o.Operand,out var number))throw new InvalidOperationException("请输入有效运算符和十进制数（小数点用 .，不含千分位）。");
                if(o.Operator=="/"&&number==0)throw new InvalidOperationException("除数不能为 0。");
            }
        }
        private static bool TryDecimal(string value,out decimal number)
        {
            number=0;value=(value??"").Trim();
            if(!Regex.IsMatch(value,@"^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]{1,3})?$",RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(50)))return false;
            // Decimal.TryParse silently rounds excess digits: reject instead of losing identifiers.
            var digits=value.Split('e','E')[0].Where(char.IsDigit).ToArray();int first=Array.FindIndex(digits,c=>c!='0');
            if(first>=0 && digits.Length-first>28)return false;
            var pieces=value.Split('e','E');string mantissa=pieces[0];int exponent=pieces.Length==2?int.Parse(pieces[1],CultureInfo.InvariantCulture):0;
            int dot=mantissa.IndexOf('.'),fraction=dot<0?0:mantissa.Length-dot-1,trailing=0;
            for(int i=digits.Length-1;i>=0&&digits[i]=='0';i--)trailing++;
            if(first>=0 && fraction-exponent-trailing>28)return false;
            if(!decimal.TryParse(value,NumberStyles.Float,CultureInfo.InvariantCulture,out number))return false;
            return number!=0||first<0;
        }
        public static SimpleMaskResult Build(SimpleMaskSnapshot source,SimpleMaskOptions options,CancellationToken token,Action<int,int> progress=null)
        {
            Validate(options);var result=new SimpleMaskResult{Options=options};int timeouts=0;
            Regex regex=options.Mode=="regex"?new Regex(options.Pattern,RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(50)):null;
            for(int i=0;i<source.Cells.Count;i++)
            {
                token.ThrowIfCancellationRequested();var c=source.Cells[i];var item=new SimpleMaskCell{Block=c.Block,Row=c.Row,Column=c.Column,Address=c.Address,Before=c.Before,Skip=c.Skip,NumberFormat=c.NumberFormat};
                if(item.Skip==null)
                {
                    if(timeouts>=10)item.Skip="正则累计超时，已跳过剩余项；请简化表达式";
                    else try{item.After=Transform(c.Before,options,regex);}
                    catch(RegexMatchTimeoutException){timeouts++;item.Skip="正则超时";}
                    catch(Exception ex) when(ex is InvalidOperationException || ex is OverflowException || ex is ArgumentException){item.Skip=ex is OverflowException?"运算超出十进制范围":ex.Message;}
                }
                result.Cells.Add(item);if(i%256==0)progress?.Invoke(i,source.Cells.Count);
            }
            progress?.Invoke(source.Cells.Count,source.Cells.Count);return result;
        }
        public static string Transform(object value,SimpleMaskOptions o,Regex regex=null)
        {
            string text=Text(value);if(string.IsNullOrWhiteSpace(text))throw new InvalidOperationException("空白");
            if(value is bool)throw new InvalidOperationException("布尔值");
            string output;
            switch(o.Mode)
            {
                case "initials":
                    var b=new StringBuilder();
                    for(int i=0;i<text.Length;i++)
                    {
                        int cp=char.ConvertToUtf32(text,i);if(cp>0xffff)i++;
                        if(cp>='A'&&cp<='Z')b.Append(Letters[cp-'A']);
                        else if(cp>='a'&&cp<='z')b.Append("〔小〕").Append(Letters[cp-'a']);
                        else if(cp>='0'&&cp<='9')b.Append(Digits[cp-'0']);
                        else if(Initials.Value.TryGetValue(cp,out char initial))b.Append(initial);
                        else b.Append("〔符").Append(((cp+0x5A17)%0x110000).ToString("X6",CultureInfo.InvariantCulture)).Append('〕');
                    }
                    output=b.ToString();break;
                case "swap": case "reverse":
                    // Text elements keep supplementary characters and combining marks together.
                    int[] positions=StringInfo.ParseCombiningCharacters(text);var elements=positions.Select((p,i)=>text.Substring(p,(i+1<positions.Length?positions[i+1]:text.Length)-p)).ToArray();
                    if(o.Mode=="reverse")output=string.Concat(elements.Reverse());
                    else {int n=o.EdgeCount;if(elements.Length<n*2)throw new InvalidOperationException("长度不足首尾各 "+n+" 个字符");output=string.Concat(elements.Skip(elements.Length-n).Concat(elements.Skip(n).Take(elements.Length-n*2)).Concat(elements.Take(n)));}break;
                case "regex":
                    regex=regex??new Regex(o.Pattern,RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(50));
                    if(!regex.IsMatch(text))throw new InvalidOperationException("正则未匹配");
                    // Avoid Regex.Replace allocating unbounded output for zero-width matches.
                    var buffer=new StringBuilder();int end=0;var timer=System.Diagnostics.Stopwatch.StartNew();
                    foreach(Match match in regex.Matches(text))
                    {if(timer.ElapsedMilliseconds>50)throw new RegexMatchTimeoutException();string replacement=BoundedReplacement(match,o.Replacement??"");if((long)buffer.Length+match.Index-end+replacement.Length>32767)throw new InvalidOperationException("结果超过 Excel 32767 字符限制");buffer.Append(text,end,match.Index-end).Append(replacement);end=match.Index+match.Length;}
                    buffer.Append(text,end,text.Length-end);output=buffer.ToString();break;
                case "number":
                    if(!TryDecimal(text,out decimal number))throw new InvalidOperationException("非纯数字或超出 28 位精度范围");
                    if(!TryDecimal(o.Operand,out decimal operand))throw new InvalidOperationException("运算参数无效");
                    decimal answer;checked{switch(o.Operator){case "+":answer=number+operand;break;case "-":answer=number-operand;break;case "*":answer=number*operand;break;case "/":if(operand==0)throw new InvalidOperationException("除数不能为 0");answer=number/operand;break;default:throw new InvalidOperationException("运算符无效");}}
                    output=answer.ToString("G29",CultureInfo.InvariantCulture);break;
                default:throw new InvalidOperationException("未知模式");
            }
            output=(o.Prefix??"")+output+(o.Suffix??"");if(output.Length>32767)throw new InvalidOperationException("结果超过 Excel 32767 字符限制");
            // .NET regex operates on UTF-16; reject expressions that split a surrogate pair.
            for(int i=0;i<output.Length;i++){if(char.IsHighSurrogate(output[i])){if(i+1>=output.Length||!char.IsLowSurrogate(output[++i]))throw new InvalidOperationException("正则拆分了 Unicode 字符");}else if(char.IsLowSurrogate(output[i]))throw new InvalidOperationException("正则拆分了 Unicode 字符");}
            return output;
        }
        private static string BoundedReplacement(Match match,string replacement)
        {
            var output=new StringBuilder();int end=0;
            foreach(Match token in Regex.Matches(replacement,@"\$(\$|&|`|'|\+|_|[0-9]+|\{[^}]+\})"))
            {
                string expanded=match.Result(token.Value);
                if((long)output.Length+token.Index-end+expanded.Length>32767)throw new InvalidOperationException("结果超过 Excel 32767 字符限制");
                output.Append(replacement,end,token.Index-end).Append(expanded);end=token.Index+token.Length;
            }
            if(output.Length+replacement.Length-end>32767)throw new InvalidOperationException("结果超过 Excel 32767 字符限制");
            return output.Append(replacement,end,replacement.Length-end).ToString();
        }
        public static MaskResult Output(SimpleMaskSnapshot source,SimpleMaskResult result)
        {
            var snapshot=new MaskSnapshot();foreach(var b in source.Blocks)snapshot.Sheets.Add(new DataSheet{Name="脱敏结果"+(snapshot.Sheets.Count+1),Rows=b.Rows,Columns=b.Columns});
            foreach(var c in result.Cells.Where(c=>c.Skip==null)){var b=source.Blocks[c.Block];snapshot.Cells.Add(new DataCell{Sheet=c.Block,Row=c.Row-b.Row+1,Column=c.Column-b.Column+1,Value=c.After});}
            return new MaskResult{Output=snapshot,Changed=result.Changed};
        }
    }
}
