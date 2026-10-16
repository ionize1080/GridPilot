using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace SheetPilot.Core
{
    [DataContract]
    public sealed class SortOptions
    {
        [DataMember] public string Rules { get; set; } = "[中文数字]";
        [DataMember] public string Sequences { get; set; } = "";
        [DataMember] public bool Descending { get; set; }
        [DataMember] public bool HasHeader { get; set; } = true;
        [DataMember] public bool MatchCase { get; set; }
        [DataMember] public bool Natural { get; set; }
        public SortOptions Clone() => (SortOptions)MemberwiseClone();
        [OnDeserializing] private void Init(StreamingContext c) { Rules="[中文数字]"; Sequences=""; HasHeader=true; }
    }

    // A permutation is calculated before Excel is touched. Equal / unmatched keys retain source order.
    public static class CustomSort
    {
        public const string ChineseCharacters = "零〇一二两三四五六七八九十百千万亿兆京垓秭穰沟涧正载壹贰叁肆伍陆柒捌玖拾佰仟萬億兩";
        private const string Digits = "零一二三四五六七八九";
        private static readonly Dictionary<char,int> Units = new Dictionary<char,int> {{'十',1},{'百',2},{'千',3},{'万',4},{'亿',8},{'兆',12},{'京',16},{'垓',20},{'秭',24},{'穰',28},{'沟',32},{'涧',36},{'正',40},{'载',44}};
        public static BigInteger ChineseNumber(string text)
        {
            if(string.IsNullOrEmpty(text) || text.Length>512) throw new ArgumentException("中文数字为空或超过 512 字符。");
            const string from="〇两兩壹贰叁肆伍陆柒捌玖拾佰仟萬億", to="零二二一二三四五六七八九十百千万亿";
            text=new string(text.Select(c=>from.IndexOf(c)>=0?to[from.IndexOf(c)]:c).ToArray());
            if(text.Any(c=>Digits.IndexOf(c)<0&&!Units.ContainsKey(c)))throw new ArgumentException("包含无效中文数字。");
            return ParseChinese(text);
        }
        private static BigInteger ParseChinese(string text)
        {
            if(text.Length==0)return BigInteger.Zero;
            int highest=0,at=-1;
            for(int i=0;i<text.Length;i++)if(Units.TryGetValue(text[i],out int power)&&power>=highest){highest=power;at=i;}
            if(at<0){BigInteger n=0;foreach(char c in text)n=n*10+Digits.IndexOf(c);return n;}
            return (at==0?BigInteger.One:ParseChinese(text.Substring(0,at)))*BigInteger.Pow(10,highest)+ParseChinese(text.Substring(at+1));
        }
        private sealed class Part { public string Name; public Dictionary<string,int> Sequence; }
        private sealed class Rule { public Regex Regex; public List<Part> Parts=new List<Part>(); }
        private sealed class Key { public int Index,Group; public List<BigInteger> Values=new List<BigInteger>(); public string Text; public object[] NaturalParts; }
        public static int[] Order(IReadOnlyList<string> values, SortOptions options)
        {
            if(options==null)throw new ArgumentNullException(nameof(options));
            if(values.Count>200000)throw new InvalidOperationException("一次最多排序 200,000 行。");
            var rules=Compile(options);var clock=Stopwatch.StartNew();var keys=new List<Key>(values.Count);
            try
            {
                for(int i=0;i<values.Count;i++)
                {
                    if(clock.ElapsedMilliseconds>5000)throw new InvalidOperationException("排序匹配超过 5 秒，请简化规则或缩小范围。");
                    string text=values[i]??"";if(text.Length>32767)throw new InvalidOperationException("排序文本过长。");
                    var key=new Key {Index=i,Group=rules.Count,Text=text};
                    if(options.Natural)key.NaturalParts=Regex.Split(text,"([0-9]+)").Select((v,j)=>j%2==1?(object)BigInteger.Parse(v,CultureInfo.InvariantCulture):v).ToArray();
                    for(int r=0;r<rules.Count;r++)
                    {
                        Match match=rules[r].Regex.Match(text);if(!match.Success)continue;
                        key.Group=r;
                        foreach(var part in rules[r].Parts)
                        {
                            string captured=match.Groups[part.Name].Value;
                            if(captured.Length==0)throw new InvalidOperationException("序列标记必须参与每次匹配；请把可选分支拆成多条规则。");
                            key.Values.Add(part.Sequence==null?ChineseNumber(captured):new BigInteger(part.Sequence[captured]));
                        }
                        break;
                    }
                    keys.Add(key);
                }
            }
            catch(RegexMatchTimeoutException){throw new InvalidOperationException("正则匹配超时，请简化嵌套或重复表达式。");}
            keys.Sort((a,b)=>
            {
                int n=a.Group.CompareTo(b.Group);if(n!=0)return n; // Rule priority is explicit, independent of direction.
                if(options.Natural)n=NaturalCompare(a.NaturalParts,b.NaturalParts,options.MatchCase);
                else for(int j=0;j<a.Values.Count&&n==0;j++)n=a.Values[j].CompareTo(b.Values[j]);
                if(n!=0)return options.Descending?-Math.Sign(n):Math.Sign(n);
                return a.Index.CompareTo(b.Index);
            });
            return keys.Select(k=>k.Index).ToArray();
        }
        private static List<Rule> Compile(SortOptions options)
        {
            var result=new List<Rule>();if(options.Natural)return result;
            if((options.Rules??"").Length>8192 || (options.Sequences??"").Length>8192)throw new InvalidOperationException("规则或自定义序列最多 8192 字符。");
            var comparison=options.MatchCase?StringComparer.Ordinal:StringComparer.OrdinalIgnoreCase;
            var sequences=new Dictionary<string,string[]>(StringComparer.Ordinal){{"大中小微",new[]{"大","中","小","微"}}};
            foreach(string line in Lines(options.Sequences))
            {
                int equal=line.IndexOf('=');if(equal<1)throw new InvalidOperationException("自定义序列格式：名称=第一项,第二项,第三项");
                string name=line.Substring(0,equal).Trim();var terms=line.Substring(equal+1).Split(',','，').Select(s=>s.Trim()).ToArray();
                if(name=="中文数字"||name.IndexOfAny(new[]{'[',']'})>=0||sequences.ContainsKey(name))throw new InvalidOperationException("序列名重复或包含保留字符："+name);
                if(terms.Length<2||terms.Length>200||terms.Any(string.IsNullOrEmpty)||terms.Distinct(comparison).Count()!=terms.Length)throw new InvalidOperationException("序列需包含 2–200 个不同的非空项目。");
                sequences.Add(name,terms);
            }
            var lines=Lines(options.Rules);if(lines.Length==0||lines.Length>32)throw new InvalidOperationException("请输入 1–32 条排序规则，每行一条。");
            foreach(string line in lines)
            {
                if(line.Length>512)throw new InvalidOperationException("每条规则最多 512 字符。");
                var rule=new Rule();
                // Unknown brackets remain ordinary .NET regex character classes.
                string pattern=Regex.Replace(line,@"\[([^\[\]]+)\]",m=>
                {
                    string name=m.Groups[1].Value;if(name!="中文数字"&&!sequences.ContainsKey(name))return m.Value;
                    var part=new Part {Name="GPToken"+rule.Parts.Count};string body;
                    if(name=="中文数字")body="["+ChineseCharacters+"]+";
                    else {var terms=sequences[name];part.Sequence=terms.Select((t,i)=>new{t,i}).ToDictionary(x=>x.t,x=>x.i,comparison);body=string.Join("|",terms.OrderByDescending(t=>t.Length).Select(Regex.Escape));}
                    rule.Parts.Add(part);return "(?<"+part.Name+">"+body+")";
                });
                if(line.Contains("GPToken"))throw new InvalidOperationException("GPToken 为保留捕获组名称。");
                try{rule.Regex=new Regex("^(?:"+pattern+")",RegexOptions.CultureInvariant|(options.MatchCase?RegexOptions.None:RegexOptions.IgnoreCase),TimeSpan.FromMilliseconds(40));}
                catch(ArgumentException ex){throw new InvalidOperationException("规则无效："+line+"。"+ex.Message);}
                result.Add(rule);
            }
            return result;
        }
        private static string[] Lines(string text)=>(text??"").Split(new[]{'\r','\n'},StringSplitOptions.RemoveEmptyEntries).Select(x=>x.Trim()).Where(x=>x.Length>0).ToArray();
        private static int NaturalCompare(object[] x,object[] y,bool matchCase)
        {
            for(int i=0;i<Math.Min(x.Length,y.Length);i++)
            {
                int n=i%2==1?((BigInteger)x[i]).CompareTo((BigInteger)y[i]):
                    CultureInfo.GetCultureInfo("zh-CN").CompareInfo.Compare((string)x[i],(string)y[i],matchCase?CompareOptions.None:CompareOptions.IgnoreCase);
                if(n!=0)return n;
            }
            return x.Length.CompareTo(y.Length);
        }
    }
}
