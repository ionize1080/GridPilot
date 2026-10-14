using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Security.Cryptography;

namespace SheetPilot.Core
{
    public sealed class FormulaResult
    {
        public object Value;
        public string Trace;
    }
    internal sealed class FormulaRange
    {
        internal int Sheet, Row, Column, Rows, Columns;
    }
    public sealed class FormulaWorkbook
    {
        internal readonly MaskSnapshot Snapshot;
        internal readonly bool NeutralNames;
        private readonly Dictionary<long,DataCell> cells;
        private readonly Dictionary<string,FormulaResult> cache=new Dictionary<string,FormulaResult>();
        private readonly HashSet<string> visiting=new HashSet<string>();
        private readonly CancellationToken cancellation;
        private long operations;
        public FormulaWorkbook(MaskSnapshot snapshot,bool neutralNames=false,CancellationToken cancellation=default(CancellationToken))
        { Snapshot=snapshot; NeutralNames=neutralNames; cells=snapshot.Cells.ToDictionary(c=>CellKey(c.Sheet,c.Row,c.Column));this.cancellation=cancellation; }
        private static long CellKey(int sheet,int row,int column)=>((long)sheet*1048576+row-1)*16384+column-1;
        internal void Step()
        { cancellation.ThrowIfCancellationRequested();if(++operations>Math.Max(5000000L,(long)Snapshot.Cells.Count*100))throw new InvalidOperationException("公式验证计算量超过本版上限，请缩小工作簿。"); }
        public FormulaResult Evaluate(string key)
        {
            Step();
            if(cache.TryGetValue(key,out var cached)) return cached;
            var parts=key.Split(':');
            if(!cells.TryGetValue(CellKey(int.Parse(parts[0],CultureInfo.InvariantCulture),int.Parse(parts[1],CultureInfo.InvariantCulture),int.Parse(parts[2],CultureInfo.InvariantCulture)),out var cell)) return new FormulaResult { Value=null,Trace="" };
            if(cell.Formula==null) return new FormulaResult { Value=cell.Value,Trace="" };
            if(visiting.Count>=128 || !visiting.Add(key)) throw new InvalidOperationException(cell.Location+"：循环引用或依赖深度超限。");
            try
            {
                var trace=new StringBuilder(); object value=new SafeFormula(cell.Formula,this,cell.Sheet,trace).Evaluate();
                string fingerprint;
                using(var hash=SHA256.Create())fingerprint=Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(trace.ToString())));
                var result=new FormulaResult { Value=value,Trace=fingerprint }; cache[key]=result; return result;
            }
            catch(InvalidOperationException ex) { throw new InvalidOperationException(cell.Location+"："+ex.Message); }
            finally { visiting.Remove(key); }
        }
        internal object Read(int sheet,int row,int column,StringBuilder trace)
        {
            Step();if(!cells.TryGetValue(CellKey(sheet,row,column),out var cell))return null;
            if(cell.Formula==null)return cell.Value;
            var key=cell.Key;var v=Evaluate(key);if(v.Trace.Length>0)trace.Append('{').Append(key).Append(':').Append(v.Trace).Append('}');return v.Value;
        }
        internal int SheetIndex(string name)
        {
            for(int i=0;i<Snapshot.Sheets.Count;i++) if(string.Equals(NeutralNames?"Sheet"+(i+1):Snapshot.Sheets[i].Name,name,StringComparison.OrdinalIgnoreCase)) return i;
            throw new InvalidOperationException("公式引用了未纳入的工作表。");
        }
    }
    // An allow-list parser, not an Excel formula replacement engine. Unknown syntax blocks export.
    public sealed class SafeFormula
    {
        private static readonly Regex Lexer=new Regex(@"\G(?:\s+|""(?:[^""]|"""")*""|'(?:[^']|'')*'|\$?[A-Za-z_\u0080-\uffff][A-Za-z0-9_.$\u0080-\uffff]*|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?|<>|<=|>=|[=<>+*/^&(),!:%-])",RegexOptions.CultureInvariant,TimeSpan.FromMilliseconds(200));
        private static readonly Regex Cell=new Regex(@"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})$",RegexOptions.CultureInvariant);
        private readonly List<string> tokens;
        private readonly FormulaWorkbook workbook;
        private readonly int sheet;
        private readonly StringBuilder trace;
        private int position, depth, evaluationDepth;
        private sealed class Node { internal string Op, Text; internal Node Left,Right; internal List<Node> Args; internal FormulaRange Range; }
        internal SafeFormula(string formula,FormulaWorkbook workbook,int sheet,StringBuilder trace)
        { tokens=Tokenize(formula); this.workbook=workbook;this.sheet=sheet;this.trace=trace; }
        public static string Rewrite(string formula,IReadOnlyList<DataSheet> sheets,Dictionary<string,string> mappings)
        {
            var tokens=Tokenize(formula);
            for(int i=0;i<tokens.Count;i++)
            {
                string t=tokens[i];
                if(t.StartsWith("\"",StringComparison.Ordinal))
                {
                    string raw=Unquote(t),result=raw;
                    if(mappings.TryGetValue(raw,out var replacement)) result=replacement;
                    else foreach(string op in new[]{"<>","<=",">=","=","<",">"})
                        if(raw.StartsWith(op,StringComparison.Ordinal) && mappings.TryGetValue(raw.Substring(op.Length),out replacement)){result=op+replacement;break;}
                    tokens[i]="\""+result.Replace("\"","\"\"")+"\"";
                }
                else if(i+1<tokens.Count && tokens[i+1]=="!")
                {
                    string name=t.StartsWith("'",StringComparison.Ordinal)?t.Substring(1,t.Length-2).Replace("''","'"):t;
                    int index=-1; for(int j=0;j<sheets.Count;j++) if(string.Equals(name,sheets[j].Name,StringComparison.OrdinalIgnoreCase)) { index=j;break; }
                    if(index<0) throw new InvalidOperationException("公式引用了未纳入的工作表。");
                    tokens[i]="'Sheet"+(index+1)+"'";
                }
            }
            return "="+string.Concat(tokens);
        }
        private static List<string> Tokenize(string formula)
        {
            if(formula==null || !formula.StartsWith("=",StringComparison.Ordinal) || formula.Length>8192) throw new InvalidOperationException("公式格式或长度不受支持。");
            var result=new List<string>(); int index=1;
            while(index<formula.Length)
            {
                var match=Lexer.Match(formula,index); if(!match.Success) throw new InvalidOperationException("公式包含未支持的语法（名称、外链、数组或结构化引用等）。");
                if(match.Value=="&" || match.Value=="^")throw new InvalidOperationException("字符串拼接及幂运算不在本版支持范围。");
                if(!string.IsNullOrWhiteSpace(match.Value)) result.Add(match.Value); index+=match.Length;
            }
            return result;
        }
        private string Peek => position<tokens.Count?tokens[position]:"";
        private string Take() { if(position>=tokens.Count) throw new InvalidOperationException("公式不完整。"); return tokens[position++]; }
        private void Expect(string token) { if(Take()!=token) throw new InvalidOperationException("公式语法不受支持。"); }
        internal object Evaluate()
        {
            var root=Parse(0); if(position!=tokens.Count) throw new InvalidOperationException("公式包含未支持的运算符。");
            return Scalar(Eval(root));
        }
        private static int Precedence(string op) { switch(op) { case "=":case "<>":case "<":case ">":case "<=":case ">=":return 1;case "&":return 2;case "+":case "-":return 3;case "*":case "/":return 4;default:return -1; } }
        private Node Parse(int minimum)
        {
            if(++depth>64) throw new InvalidOperationException("公式嵌套层数超限。");
            try
            {
                Node left; string token=Take();
                if(token=="+" || token=="-") left=new Node { Op="unary"+token,Left=Parse(5) };
                else if(token=="(") { left=Parse(0);Expect(")"); }
                else if(token.StartsWith("\"",StringComparison.Ordinal)) left=new Node { Op="text",Text=Unquote(token) };
                else if(double.TryParse(token,NumberStyles.Float,CultureInfo.InvariantCulture,out var number)) left=new Node { Op="number",Text=token };
                else if(Peek=="(")
                {
                    if(!new[]{"IF","SUM","COUNT","AVERAGE","MIN","MAX","AND","OR","NOT","SUMIF","SUMIFS","COUNTIF","COUNTIFS","VLOOKUP","MATCH","INDEX"}.Contains(token.ToUpperInvariant())) throw new InvalidOperationException("暂不支持函数 "+token+"，已停止导出。");
                    Take(); var args=new List<Node>(); if(Peek!=")") { do { args.Add(Parse(0)); if(Peek!=",") break; Take(); } while(true); } Expect(")");
                    left=new Node { Op="function",Text=token.ToUpperInvariant(),Args=args };
                }
                else if(string.Equals(token,"TRUE",StringComparison.OrdinalIgnoreCase) || string.Equals(token,"FALSE",StringComparison.OrdinalIgnoreCase)) left=new Node { Op="bool",Text=token.ToUpperInvariant() };
                else
                {
                    int target=sheet;
                    if(Peek=="!") { Take(); target=workbook.SheetIndex(token.StartsWith("'",StringComparison.Ordinal)?token.Substring(1,token.Length-2).Replace("''","'"):token); token=Take(); }
                    int row,column,endRow,endColumn;
                    if(Peek==":" && Regex.IsMatch(token,@"^\$?[A-Za-z]{1,3}$"))
                    {
                        ParseAddress(token+"1",out row,out column);Take();string end=Take();
                        if(!Regex.IsMatch(end,@"^\$?[A-Za-z]{1,3}$"))throw new InvalidOperationException("整列引用不完整。");
                        ParseAddress(end+"1048576",out endRow,out endColumn);
                    }
                    else
                    {
                        ParseAddress(token,out row,out column);endRow=row;endColumn=column;
                        if(Peek==":"){Take();ParseAddress(Take(),out endRow,out endColumn);}
                    }
                    if(endRow<row || endColumn<column || (long)(endRow-row+1)*(endColumn-column+1)>8000000) throw new InvalidOperationException("单个公式引用区域超过800万格验证预算或方向不受支持。");
                    left=new Node { Op="range",Range=new FormulaRange { Sheet=target,Row=row,Column=column,Rows=endRow-row+1,Columns=endColumn-column+1 } };
                }
                if(Peek=="%") { Take();left=new Node { Op="percent",Left=left }; }
                while(Precedence(Peek)>=minimum) { string op=Take();left=new Node { Op=op,Left=left,Right=Parse(Precedence(op)+1) }; }
                return left;
            }
            finally { depth--; }
        }
        private static void ParseAddress(string token,out int row,out int column)
        {
            var match=Cell.Match(token); if(!match.Success) throw new InvalidOperationException("暂不支持名称、整行列、动态或结构化引用。");
            column=0;foreach(char c in match.Groups[1].Value.ToUpperInvariant()) column=column*26+c-'A'+1;
            row=int.Parse(match.Groups[2].Value,CultureInfo.InvariantCulture); if(column>16384 || row>1048576) throw new InvalidOperationException("引用超出工作表范围。");
        }
        private object Eval(Node node)
        {
            if(++evaluationDepth>128)throw new InvalidOperationException("公式运算嵌套超过本版上限。");
            try {return EvalNode(node);}finally {evaluationDepth--;}
        }
        private object EvalNode(Node node)
        {
            workbook.Step();
            switch(node.Op)
            {
                case "text": return node.Text;
                case "number": return double.Parse(node.Text,CultureInfo.InvariantCulture);
                case "bool": return node.Text=="TRUE";
                case "range": return node.Range;
                case "unary+":return Number(Scalar(Eval(node.Left)));
                case "unary-":return -Number(Scalar(Eval(node.Left)));
                case "percent":return Number(Scalar(Eval(node.Left)))/100;
                case "function":return Function(node.Text,node.Args);
            }
            object a=Scalar(Eval(node.Left)),b=Scalar(Eval(node.Right));
            switch(node.Op)
            {
                case "+":return Number(a)+Number(b);case "-":return Number(a)-Number(b);case "*":return Number(a)*Number(b);
                case "/": if(Number(b)==0) throw new InvalidOperationException("除零公式不支持导出。"); return Number(a)/Number(b);
                case "&":throw new InvalidOperationException("字符串拼接需人工声明约束，本版停止导出。");
                default: bool compared=Compare(a,b,node.Op);trace.Append('C').Append(compared?'1':'0');return compared;
            }
        }
        private object Function(string name,List<Node> args)
        {
            if(name=="IF") { Arity(args,3,3);bool condition=Boolean(Scalar(Eval(args[0]))); trace.Append("IF").Append(condition?'1':'0');return Eval(args[condition?1:2]); }
            if(name=="SUM" || name=="COUNT" || name=="AVERAGE" || name=="MIN" || name=="MAX")
            {
                Arity(args,1,255);var numbers=new List<double>(); foreach(var arg in args) foreach(var value in Values(Eval(arg))) if(value is double d) numbers.Add(d); else if(value!=null && !(value is string) && !(value is bool)) throw new InvalidOperationException("公式引用了错误值。");
                if(name=="COUNT") return (double)numbers.Count;
                if(name=="SUM") return numbers.Sum();if(numbers.Count==0) throw new InvalidOperationException("空数值汇总不受支持。");
                if(name=="AVERAGE")return numbers.Average();if(name=="MIN")return numbers.Min();return numbers.Max();
            }
            if(name=="AND" || name=="OR" || name=="NOT")
            { Arity(args,1,name=="NOT"?1:255);var flags=args.Select(n=>Boolean(Scalar(Eval(n)))).ToArray();return name=="NOT"?!flags[0]:name=="AND"?flags.All(v=>v):flags.Any(v=>v); }
            if(name=="SUMIF" || name=="COUNTIF" || name=="SUMIFS" || name=="COUNTIFS")
            {
                var ranges=new List<FormulaRange>();var criteria=new List<object>();FormulaRange sum=null;
                if(name=="SUMIF" || name=="COUNTIF")
                { Arity(args,2,name=="SUMIF"?3:2);ranges.Add(Range(Eval(args[0])));criteria.Add(Scalar(Eval(args[1])));if(name=="SUMIF") sum=args.Count==3?Range(Eval(args[2])):ranges[0]; }
                else
                { int start=name=="SUMIFS"?1:0; if(args.Count<start+2 || (args.Count-start)%2!=0) throw new InvalidOperationException("条件汇总参数不完整。");if(start==1)sum=Range(Eval(args[0]));for(int i=start;i<args.Count;i+=2) {ranges.Add(Range(Eval(args[i])));criteria.Add(Scalar(Eval(args[i+1])));} }
                int count=ranges[0].Rows*ranges[0].Columns;
                if(ranges.Any(r=>r.Rows!=ranges[0].Rows || r.Columns!=ranges[0].Columns) || sum!=null && (sum.Rows!=ranges[0].Rows || sum.Columns!=ranges[0].Columns)) throw new InvalidOperationException("条件范围尺寸不一致。");
                var values=ranges.Select(r=>Values(r).ToArray()).ToArray();var sums=sum==null?null:Values(sum).ToArray(); double total=0;
                trace.Append(name).Append('[');
                for(int i=0;i<count;i++)
                {
                    bool match=true; for(int j=0;j<ranges.Count;j++) if(!Criterion(values[j][i],criteria[j])) match=false;
                    if(match) {trace.Append(i).Append(','); if(sums==null) total++;else if(sums[i] is double number) total+=number;}
                }
                trace.Append(']');return total;
            }
            if(name=="VLOOKUP")
            {
                Arity(args,4,4);object query=Scalar(Eval(args[0]));var range=Range(Eval(args[1]));int column=Integer(Scalar(Eval(args[2])));
                if(Boolean(Scalar(Eval(args[3])))) throw new InvalidOperationException("仅支持精确 VLOOKUP（第四参数为 FALSE）。");
                if(column<1 || column>range.Columns) throw new InvalidOperationException("查找列号超限。");
                for(int r=0;r<range.Rows;r++) if(LookupEqual(Read(range,r,0),query)) {trace.Append("V").Append(r).Append(':').Append(column);return Read(range,r,column-1);}
                throw new InvalidOperationException("精确查找未命中。");
            }
            if(name=="MATCH")
            {
                Arity(args,3,3);object query=Scalar(Eval(args[0]));var range=Range(Eval(args[1]));if(Number(Scalar(Eval(args[2])))!=0 || range.Rows!=1 && range.Columns!=1) throw new InvalidOperationException("仅支持一维精确 MATCH。");
                var values=Values(range).ToArray();for(int i=0;i<values.Length;i++) if(LookupEqual(values[i],query)) {trace.Append('M').Append(i);return (double)(i+1);}throw new InvalidOperationException("MATCH 未命中。");
            }
            if(name=="INDEX")
            {
                Arity(args,2,3);var range=Range(Eval(args[0]));int r=Integer(Scalar(Eval(args[1]))),c=args.Count==3?Integer(Scalar(Eval(args[2]))):1;
                if(args.Count==2 && range.Rows==1) {c=r;r=1;} if(r<1 || r>range.Rows || c<1 || c>range.Columns) throw new InvalidOperationException("INDEX 超出范围。");
                trace.Append('I').Append(r).Append(':').Append(c);return Read(range,r-1,c-1);
            }
            throw new InvalidOperationException("暂不支持函数 "+name+"，已停止导出。");
        }
        private object Read(FormulaRange r,int row,int column)=>workbook.Read(r.Sheet,r.Row+row,r.Column+column,trace);
        private object Scalar(object value)
        { if(!(value is FormulaRange r)) return value;if(r.Rows!=1 || r.Columns!=1) throw new InvalidOperationException("不支持隐式交叉或数组运算。");return Read(r,0,0); }
        private IEnumerable<object> Values(object value)
        { if(value is FormulaRange r) {for(int row=0;row<r.Rows;row++)for(int col=0;col<r.Columns;col++)yield return Read(r,row,col);} else yield return value; }
        private static FormulaRange Range(object v)=>v as FormulaRange??throw new InvalidOperationException("此参数必须是单元格范围。");
        private static double Number(object v) {if(v==null)return 0;if(v is double d)return d;if(v is bool b)return b?1:0;throw new InvalidOperationException("不支持公式中的隐式文本转数值。");}
        private static int Integer(object v) {double d=Number(v);if(d!=Math.Truncate(d)||d<int.MinValue||d>int.MaxValue)throw new InvalidOperationException("索引必须是整数。");return (int)d;}
        private static bool Boolean(object v) {if(v is bool b)return b;if(v==null)return false;if(v is double d)return d!=0;throw new InvalidOperationException("逻辑参数类型不受支持。");}
        private static void Arity(List<Node> a,int min,int max) {if(a.Count<min || a.Count>max)throw new InvalidOperationException("函数参数数量不受支持。");}
        private static bool Criterion(object value,object criterion)
        {
            string op="=";object target=criterion;
            if(criterion is string text)
            {
                if(text.IndexOfAny(new[]{'*','?','~'})>=0) throw new InvalidOperationException("通配符条件需独立约束，本版停止导出。");
                foreach(string candidate in new[]{"<>","<=",">=","=","<",">"}) if(text.StartsWith(candidate,StringComparison.Ordinal)) {op=candidate;text=text.Substring(candidate.Length);break;}
                if(double.TryParse(text,NumberStyles.Float,CultureInfo.InvariantCulture,out double number)) target=number; else target=text;
            }
            return Compare(value,target,op);
        }
        private static bool LookupEqual(object value,object query)
        {
            if(query is string text && text.IndexOfAny(new[]{'*','?','~'})>=0)throw new InvalidOperationException("通配符查找不在本版支持范围。");
            return Compare(value,query,"=");
        }
        public static IEnumerable<string> TextLiterals(string formula)=>Tokenize(formula).Where(t=>t.StartsWith("\"",StringComparison.Ordinal)).Select(Unquote);
        private static bool Compare(object a,object b,string op)
        {
            int c;if(a is string x && b is string y)c=string.Compare(x,y,StringComparison.OrdinalIgnoreCase);
            else if((a==null||a is double) && (b==null||b is double)) c=Number(a).CompareTo(Number(b));
            else if(a is bool aa && b is bool bb)c=aa.CompareTo(bb);
            else if(op=="=" || op=="<>") return op=="<>";
            else throw new InvalidOperationException("混合类型大小比较不受支持。");
            switch(op) {case "=":return c==0;case "<>":return c!=0;case "<":return c<0;case ">":return c>0;case "<=":return c<=0;case ">=":return c>=0;default:throw new InvalidOperationException("运算不受支持。");}
        }
        private static string Unquote(string text)=>text.Substring(1,text.Length-2).Replace("\"\"","\"");
    }
}
