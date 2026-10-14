using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace SheetPilot.Core
{
    public sealed class HeaderMerge { public int Row, Column, Rows, Columns; }
    public static class MaskInference
    {
        private static string Clean(string text) => Regex.Replace(text ?? "", @"[\s\p{P}\p{S}]", "").ToLowerInvariant();
        private static readonly string[] Labels = {"名称","姓名","客户","借款人","被投资人","机构","项目","代码","编号","证件","账号","账户","电话","手机","地址","金额","余额","日期","时间","序号","类型","类别","比例","增量","批复","是否","name","customer","amount","date","id"};
        public static IReadOnlyList<MaskRule> Suggest(MaskSnapshot source,CancellationToken cancel=default(CancellationToken))
        {
            var rules = new List<MaskRule>();
            var samples=new Dictionary<MaskRule,HashSet<string>>();
            var affixSamples=new Dictionary<MaskRule,string[]>();
            foreach(var sg in source.Cells.GroupBy(c=>c.Sheet))
            {
                var sheet=source.Sheets[sg.Key];
                var cells=sg.ToArray();
                int firstDataRow=cells.Min(x=>x.Row);
                var top=cells.Where(c=>c.Row<=firstDataRow+39).ToArray();
                var candidates=top.GroupBy(c=>c.Row).Select(g=>new { Row=g.Key, Score=g.Count(c=>c.Value is string t && Labels.Any(l=>Clean(t).Contains(l))) }).OrderByDescending(g=>g.Score).ThenBy(g=>g.Row).ToArray();
                int header=candidates.Length>0 && candidates[0].Score>=2 ? candidates[0].Row : 0;
                if(header==0 && top.Length>0)
                {
                    var first=top.Min(c=>c.Row);var row=top.Where(c=>c.Row==first).ToArray();
                    if(row.Length>=2 && row.All(c=>c.Value is string) && row.Any(c=>Labels.Any(l=>Clean(Convert.ToString(c.Value)).Contains(l)))) header=first;
                }
                int end=header;
                if(header>0)
                {
                    foreach(var m in sheet.HeaderMerges.Where(m=>m.Row<=header && m.Row+m.Rows>header)) end=Math.Max(end,m.Row+m.Rows-1);
                    // A subordinate header row has multiple known labels and no numeric records.
                    for(int r=end+1;r<=header+4;r++)
                    { var row=top.Where(c=>c.Row==r).ToArray();if(row.Length<2 || row.Any(c=>c.Value is double) || row.Count(c=>Labels.Any(l=>Clean(Convert.ToString(c.Value)).Contains(l)))<2) break;end=r; }
                }
                foreach(var cg in cells.GroupBy(c=>c.Column))
                {
                    cancel.ThrowIfCancellationRequested();
                    var first=cg.OrderBy(c=>c.Row).First();
                    var labels=new List<string>();
                    if(header>0) for(int r=header;r<=end;r++)
                    {
                        var cell=top.FirstOrDefault(c=>c.Row==r && c.Column==cg.Key);
                        var merge=sheet.HeaderMerges.FirstOrDefault(m=>r>=m.Row && r<m.Row+m.Rows && cg.Key>=m.Column && cg.Key<m.Column+m.Columns);
                        if(cell==null && merge!=null)cell=top.FirstOrDefault(c=>c.Row==merge.Row && c.Column==merge.Column);
                        string label=Convert.ToString(cell?.Value);if(!string.IsNullOrWhiteSpace(label) && !labels.Contains(label))labels.Add(label);
                    }
                    string name=string.Join(" / ",labels);string clean=Clean(name);
                    var rule=new MaskRule {Sheet=sg.Key,SheetName=sheet.Name,Column=cg.Key,StartRow=header>0?end+1:first.Row,Header=name.Length>0?name:"未识别表头 · "+DataCell.Address(first.Row,cg.Key),Domain=clean.Length>0?clean:"字段"+cg.Key,Reason=header>0?"根据表头建议，请确认":"未识别表头，请确认数据起始行"};
                    var data=cg.Where(c=>c.Row>=rule.StartRow && c.Value!=null).Take(100).ToArray();
                    bool strings=data.Length>0 && data.All(c=>c.Value is string);
                    if(strings && Regex.IsMatch(clean,"机构名称|公司名称|企业名称|客户名称|借款人|被投资人|姓名|客户$|customer|name"))
                    {rule.Kind=MaskKind.ChineseName;rule.Domain=Regex.IsMatch(clean,"机构|公司|企业")?"机构":Regex.IsMatch(clean,"姓名|联系人")?"个人姓名":"客户";}
                    else if(strings && Regex.IsMatch(clean,"项目名称")){rule.Kind=MaskKind.ChineseName;rule.Domain="项目";}
                    else if(strings && Regex.IsMatch(clean,"代码|编号|证件|账号|账户|电话|手机|email")){rule.Kind=MaskKind.Identifier;rule.Domain=Regex.IsMatch(clean,"信用代码|信用编码")?"信用代码":clean;}
                    // Money/date/category fields are deliberately opt-in: a guess must not break business constraints.
                    var sample=AffixAnalyzer.Sample(cg.Where(c=>c.Row>=rule.StartRow),cancel);
                    if(rule.Kind==MaskKind.ChineseName && sample.Length>0 && sample.Count(v=>Regex.IsMatch(v,@"^[\x20-\x7E]+$"))>=sample.Length*.8)rule.Kind=MaskKind.EnglishName;
                    affixSamples[rule]=sample;samples[rule]=new HashSet<string>(sample,StringComparer.Ordinal);
                    rules.Add(rule);
                }
            }
            // Aliases need both compatible identity semantics and substantial observed value overlap.
            var identity=rules.Where(r=>r.Kind==MaskKind.ChineseName || r.Kind==MaskKind.EnglishName).ToArray();
            for(int i=0;i<identity.Length;i++) for(int j=0;j<i;j++)
            {
                cancel.ThrowIfCancellationRequested();
                var a=identity[i];var b=identity[j];if(a.Kind!=b.Kind || a.Domain==b.Domain)continue;if(a.Domain=="项目" || b.Domain=="项目" || a.Domain=="个人姓名" || b.Domain=="个人姓名")continue;
                int minimum=Math.Min(samples[a].Count,samples[b].Count),overlap=samples[a].Count(samples[b].Contains);
                if(minimum>0 && overlap>=Math.Min(2,minimum) && overlap/(double)minimum>=.6){a.Domain=b.Domain;a.Reason="与「"+b.SheetName+" · "+b.Header+"」有重复值，建议共用映射";}
            }
            foreach(var group in rules.Where(r=>r.Kind>=MaskKind.ChineseName && r.Kind<=MaskKind.Identifier).GroupBy(r=>new {r.Domain,r.Kind}))
            {
                var first=group.First();AffixAnalyzer.Apply(first,group.SelectMany(r=>affixSamples[r]));
                foreach(var rule in group){rule.Prefix=first.Prefix;rule.Suffix=first.Suffix;rule.Digits=first.Digits;rule.ShortName=first.ShortName;}
            }
            return rules.OrderBy(r=>r.Sheet).ThenBy(r=>r.Column).ToArray();
        }
    }
}
