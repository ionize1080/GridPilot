using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace SheetPilot.Core
{
    public sealed class HeaderProbeCell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public object Value { get; set; }
        public bool Formula { get; set; }
        public bool Bold { get; set; }
        public string Style { get; set; }
    }
    public sealed class SimpleHeaderBand
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int FirstColumn { get; set; }
        public int LastColumn { get; set; }
        public string Source { get; set; }
        public string Reason { get; set; }
        public double Score { get; set; }
        public List<SimpleHeaderField> Fields { get; set; } = new List<SimpleHeaderField>();
        public bool Contains(int row,int column) => row>=StartRow&&row<=EndRow&&column>=FirstColumn&&column<=LastColumn;
    }
    public sealed class SimpleHeaderField
    {
        public int Column { get; set; }
        public string Label { get; set; }
    }
    public sealed class SimpleHeaderAnalysis
    {
        public List<SimpleHeaderBand> Bands { get; set; } = new List<SimpleHeaderBand>();
        public string Message { get; set; }
        public string Role(int row,int column)
        {
            if(Bands.Any(b=>b.Contains(row,column)))return "header";
            var first=Bands.Where(b=>column>=b.FirstColumn&&column<=b.LastColumn).OrderBy(b=>b.StartRow).FirstOrDefault();
            return first!=null&&row<first.StartRow?"preamble":"data";
        }
    }
    public static class SimpleHeaderDetector
    {
        // Anchored labels, not substring matching: a company/customer value containing a field word is not a header.
        private static readonly HashSet<string> Labels=new HashSet<string>(("序号,行号,月份,年月,年度,年份,季度,日期,时间,姓名,名称,客户,借款人,客户名,客户号,项目名,项目号,贷款号,机构,部门,币种,贷款用途,余额,金额,数量,单价,总价,合计,备注,摘要,状态,类别,种类,类型,单位,地址,电话,手机,邮箱,账号,账户,产品,地区,国家,省份,城市,规格,型号,性别,年龄,国籍,等级,期限,利率,汇率,比例,负责人,联系人,导出者,导出人,制表人,统计时间,导出时间,报告期,年初,本期,期末,上期,本月,上月,收入,支出,借方,贷方,id,no,name,customer,customerid,customername,account,accountnumber,date,time,month,year,amount,balance,quantity,price,total,description,notes,remarks,status,type,currency,address,email,phone,code,department,product,category,country,city").Split(','),StringComparer.Ordinal);
        private static string Clean(string value)=>Regex.Replace(value??"",@"[\s\p{P}\p{S}]","").ToLowerInvariant();
        private static bool Metadata(string value)
        {
            string text=(value??"").Trim();
            return Regex.IsMatch(text,@"^(导出|打印|制表|填报|报送|生成|统计|截至|截止|单位|机构|操作员|经办人|报表日期|数据日期|exported|generated|printed|prepared|as of).{0,24}[:：].+",RegexOptions.IgnoreCase)
                ||Regex.IsMatch(text,@"^(导出|打印|制表|填报|报送|生成|统计|截至|截止).*(?:20\d{2}[-/年.]|\d{1,2}:\d{2})")
                ||Regex.IsMatch(text,@"^(导出者|导出人|制表人|操作员|经办人)\s+\S+");
        }
        private static double Label(object value)
        {
            if(!(value is string s)||s.Length>60||Metadata(s))return 0;string t=Clean(s);
            if(Labels.Contains(t))return 1;
            if(t.Length>30||Regex.IsMatch(t,@"\d|有限公司|有限责任公司|股份公司|集团公司|银行$"))return 0;
            if(Regex.IsMatch(t,@"(?:名称|姓名|代码|编码|编号|证件号|账号|客户号|项目号|贷款号|合同号|流水号|月份|日期|时间|金额|余额|数量|比例|利率|汇率|标识|描述|类别|种类|用途|期限|规模|发放日|到期日|开户日|电话|邮箱)$"))return .9;
            return 0;
        }
        private static bool DataLike(object value)
        {
            if(value==null)return false;if(!(value is string))return true;string s=((string)value).Trim();
            return Regex.IsMatch(s,@"\d|有限公司|有限责任公司|集团|银行$|@|^[A-Z]{3}$");
        }
        public static SimpleHeaderAnalysis Analyze(SimpleMaskSnapshot snapshot,SimpleMaskOptions options,CancellationToken token)
        {
            var result=new SimpleHeaderAnalysis();
            if(options.HeaderMode=="off"){result.Message="已关闭表头识别，选区内内容按规则处理。";return result;}
            if(options.HeaderMode=="manual")
            {
                foreach(var row in ParseRows(options.HeaderRows))result.Bands.Add(new SimpleHeaderBand{StartRow=row,EndRow=row,FirstColumn=1,LastColumn=16384,Source="manual",Score=1,Reason="用户指定的表头行"});
                result.Message="手动指定表头：第 "+options.HeaderRows+" 行。";return result;
            }
            if(snapshot.HeaderReadWarning!=null){result.Message=snapshot.HeaderReadWarning;return result;}
            var cells=snapshot.HeaderContext.Where(c=>c.Value!=null&&!string.IsNullOrWhiteSpace(SimpleMask.Text(c.Value))).ToArray();
            if(cells.Length==0){result.Message="未可靠识别表头；可手动指定表头行。";return result;}
            var byRow=cells.GroupBy(c=>c.Row).ToDictionary(g=>g.Key,g=>g.OrderBy(c=>c.Column).ToArray());
            foreach(var native in snapshot.NativeHeaders)
            {var band=new SimpleHeaderBand{StartRow=native.StartRow,EndRow=native.EndRow,FirstColumn=native.FirstColumn,LastColumn=native.LastColumn,Source="excel-table",Score=1,Reason="Excel 表格的正式标题行"};FillFields(band,cells,snapshot.HeaderMerges);result.Bands.Add(band);}
            var candidates=new List<SimpleHeaderBand>();
            foreach(int row in byRow.Keys.OrderBy(r=>r))
            {
                token.ThrowIfCancellationRequested();var initial=byRow[row];
                if(initial.Any(c=>Metadata(SimpleMask.Text(c.Value))))continue;
                if(initial.All(c=>Label(c.Value)==0))continue;
                // A lone label spanning several columns is a report title, even if it says “余额”.
                if(initial.Length==1&&snapshot.HeaderMerges.Any(m=>m.Row==row&&m.Column==initial[0].Column&&m.Columns>1))continue;
                int end=row;
                // Follow vertical labels and horizontal parent/child groups, capped at five rows.
                for(int pass=0;pass<5;pass++)
                {
                    int old=end;
                    foreach(var m in snapshot.HeaderMerges.Where(m=>m.Row>=row&&m.Row<=end&&m.Rows>1))end=Math.Max(end,Math.Min(row+4,m.Row+m.Rows-1));
                    if(byRow.TryGetValue(end+1,out var next)&&end<row+4)
                    {
                        bool parents=snapshot.HeaderMerges.Any(m=>m.Row>=row&&m.Row<=end&&m.Columns>1);
                        if(parents&&next.Length>=2&&next.All(c=>c.Value is string&&!DataLike(c.Value)&&!Metadata(SimpleMask.Text(c.Value)))&&next.Count(c=>Label(c.Value)>0)>=Math.Max(1,next.Length*.5))end++;
                    }
                    if(old==end)break;
                }
                var head=cells.Where(c=>c.Row>=row&&c.Row<=end).ToArray();
                if(head.Any(c=>c.Formula)||head.Any(c=>Metadata(SimpleMask.Text(c.Value))))continue;
                int left=head.Min(c=>c.Column),right=head.Max(c=>c.Column);
                foreach(var m in snapshot.HeaderMerges.Where(m=>m.Row>=row&&m.Row<=end)){left=Math.Min(left,m.Column);right=Math.Max(right,m.Column+m.Columns-1);}
                var band=new SimpleHeaderBand{StartRow=row,EndRow=end,FirstColumn=left,LastColumn=right,Source="inferred"};FillFields(band,cells,snapshot.HeaderMerges);
                var below=cells.Where(c=>c.Row>end&&c.Row<=end+6&&c.Column>=left&&c.Column<=right).ToArray();
                var dataRows=below.GroupBy(c=>c.Row).Where(g=>g.Count()>=Math.Max(1,band.Fields.Count*.4))
                    .Where(g=>g.Count(c=>band.Fields.Any(f=>f.Column==c.Column&&Clean(f.Label)==Clean(SimpleMask.Text(c.Value))))<band.Fields.Count*.6).Take(4).ToArray();
                if(dataRows.Length==0)continue;
                double labels=head.Average(c=>Label(c.Value)),texts=head.Count(c=>c.Value is string)/(double)head.Length;
                double coverage=band.Fields.Count/(double)Math.Max(1,right-left+1);
                double bodyLabels=dataRows.SelectMany(g=>g).Average(c=>Label(c.Value));
                double data=dataRows.SelectMany(g=>g).Count(c=>DataLike(c.Value))/(double)dataRows.Sum(g=>g.Count());
                double style=head.Count(c=>c.Bold||dataRows.SelectMany(g=>g).Any(d=>d.Column==c.Column&&d.Style!=c.Style&&!string.IsNullOrEmpty(c.Style)))/(double)head.Length;
                if(labels<.45||texts<.7||bodyLabels>=labels-.2)continue;
                if(head.Length==1&&(labels<.95||style<.5&&data<.5))continue;
                band.Score=Math.Min(.99,.44*labels+.16*texts+.12*coverage+.12*Math.Min(1,data*2)+.1*(labels-bodyLabels)+.06*style);
                if(band.Score<.72)continue;
                band.Reason="字段名称与下方记录一致，结合"+(style>0?"样式差异、":"")+(end>row?"多层合并结构、":"")+"列覆盖及数据类型判断";
                candidates.Add(band);
            }
            // Prefer the complete parent/child band over a stronger-looking lower tier alone.
            foreach(var candidate in candidates.OrderByDescending(c=>c.Score+.04*(c.EndRow-c.StartRow)).ThenBy(c=>c.StartRow))
            {
                if(result.Bands.Any(b=>candidate.StartRow<=b.EndRow&&candidate.EndRow>=b.StartRow))continue;
                if(result.Bands.Any(b=>candidate.FirstColumn<=b.LastColumn&&candidate.LastColumn>=b.FirstColumn)&&!IsRepeat(candidate,result.Bands,cells))continue;
                result.Bands.Add(candidate);
            }
            result.Bands=result.Bands.OrderBy(b=>b.StartRow).ToList();
            result.Message=result.Bands.Count==0?"未可靠识别表头；不会自动保留疑似数据，可手动指定表头行。":"识别表头："+string.Join("；",result.Bands.Select(b=>"第 "+b.StartRow+(b.EndRow==b.StartRow?"":"–"+b.EndRow)+" 行（"+b.Fields.Count+" 个字段）"))+"。请核对，可手动修正。";
            if(snapshot.HeaderContextLimited)result.Message+=" 自动识别仅扫描首 80 个使用行、256 列及正式 Excel 表头。";
            return result;
        }
        private static bool IsRepeat(SimpleHeaderBand candidate,List<SimpleHeaderBand> previous,HeaderProbeCell[] cells) => previous.Any(b=>b.Fields.Count>=2&&b.Fields.Count==candidate.Fields.Count&&b.Fields.Zip(candidate.Fields,(a,c)=>a.Column==c.Column&&Clean(a.Label)==Clean(c.Label)).All(v=>v));
        private static void FillFields(SimpleHeaderBand band,HeaderProbeCell[] cells,List<HeaderMerge> merges)
        {
            for(int col=band.FirstColumn;col<=band.LastColumn;col++)
            {
                var labels=new List<string>();
                for(int row=band.StartRow;row<=band.EndRow;row++)
                {
                    var c=cells.FirstOrDefault(v=>v.Row==row&&v.Column==col);
                    if(c==null){var m=merges.FirstOrDefault(v=>row>=v.Row&&row<v.Row+v.Rows&&col>=v.Column&&col<v.Column+v.Columns);if(m!=null)c=cells.FirstOrDefault(v=>v.Row==m.Row&&v.Column==m.Column);}
                    string text=SimpleMask.Text(c?.Value);if(text.Length>0&&!labels.Contains(text))labels.Add(text);
                }
                if(labels.Count>0)band.Fields.Add(new SimpleHeaderField{Column=col,Label=string.Join(" / ",labels)});
            }
        }
        public static int[] ParseRows(string text)
        {
            var rows=new SortedSet<int>();foreach(var item in (text??"").Replace('，',',').Split(','))
            {
                var match=Regex.Match(item.Trim(),@"^(\d{1,7})(?:\s*-\s*(\d{1,7}))?$");
                if(!match.Success)throw new InvalidOperationException("表头行格式例如 2、4-5 或 2,8-9。");
                int first=int.Parse(match.Groups[1].Value,CultureInfo.InvariantCulture),last=match.Groups[2].Success?int.Parse(match.Groups[2].Value,CultureInfo.InvariantCulture):first;
                if(first<1||last<first||last>1048576||last-first>99)throw new InvalidOperationException("表头行超出有效范围；最多指定 100 行。");
                for(int row=first;row<=last;row++)rows.Add(row);if(rows.Count>100)throw new InvalidOperationException("最多指定 100 个表头行。");
            }
            if(rows.Count==0)throw new InvalidOperationException("请输入表头行。");return rows.ToArray();
        }
    }
}
