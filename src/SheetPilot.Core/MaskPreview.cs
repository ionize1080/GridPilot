using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;

namespace SheetPilot.Core
{
    public static class MaskPreview
    {
        public static object Page(MaskSnapshot source,MaskResult result,string search,int offset,int size,CancellationToken token)
        {
            size=Math.Max(1,Math.Min(200,size));offset=Math.Max(0,Math.Min(19900,offset));search=search??"";
            var matches=result.Output.Cells.Select((c,i)=>new {Cell=c,Before=source.Cells[i].Value,After=c.Formula==null?c.Value:result.Calculated[c.Key]})
                .Where(p=>{token.ThrowIfCancellationRequested();return !Equals(p.Before,p.After)&&(search.Length==0||new[]{p.Cell.Location,Convert.ToString(p.Before,CultureInfo.InvariantCulture),Convert.ToString(p.After,CultureInfo.InvariantCulture)}.Any(v=>v.IndexOf(search,StringComparison.OrdinalIgnoreCase)>=0));}).Take(20001).ToArray();
            return new {summary=result.Summary,changed=result.Changed,formulas=result.Formulas,total=Math.Min(20000,matches.Length),capped=matches.Length>20000,offset,pageSize=size,
                rows=matches.Take(20000).Skip(offset).Take(size).Select(p=>new {location=p.Cell.Location,before=p.Before,after=p.After}).ToArray()};
        }
        public static object Clipboard(MaskResult result,int sheet,int firstRow,int lastRow,int rowLimit=20000)
        {
            if(sheet<0||sheet>=result.Output.Sheets.Count)throw new InvalidOperationException("请选择要复制的工作表。");
            var spec=result.Output.Sheets[sheet];if(firstRow<1||lastRow<firstRow||lastRow>spec.Rows||lastRow-firstRow+1>rowLimit)throw new InvalidOperationException("复制范围必须位于工作表内，单次最多 "+rowLimit+" 行。");
            if((long)(lastRow-firstRow+1)*spec.Columns>1000000)throw new InvalidOperationException("复制范围超过 100 万格，请缩小行范围。");
            var cells=result.Output.Cells.Where(c=>c.Sheet==sheet&&c.Row>=firstRow&&c.Row<=lastRow).ToDictionary(c=>(long)c.Row*16384+c.Column);
            var plain=new StringBuilder();var html=new StringBuilder("<html><body><table>");
            for(int row=firstRow;row<=lastRow;row++)
            {
                html.Append("<tr>");for(int col=1;col<=spec.Columns;col++)
                {
                    string value="";if(cells.TryGetValue((long)row*16384+col,out var c))value=Convert.ToString(c.Formula==null?c.Value:result.Calculated[c.Key],CultureInfo.InvariantCulture)??"";
                    if(col>1)plain.Append('\t');
                    // Excel's HTML clipboard is explicitly text, preserving identifiers and preventing formula execution.
                    string safe=value.Length>0&&"=+-@".IndexOf(value.TrimStart().FirstOrDefault())>=0?"'"+value:value;
                    plain.Append(safe.IndexOfAny(new[]{'\t','\r','\n','"'})>=0?"\""+safe.Replace("\"","\"\"")+"\"":safe);
                    html.Append("<td style=\"mso-number-format:'\\@';white-space:pre-wrap\">").Append(value.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace("\"","&quot;")).Append("</td>");
                }
                plain.Append("\r\n");html.Append("</tr>");if(plain.Length+html.Length>16*1024*1024)throw new InvalidOperationException("复制内容超过 16 MB，请缩小范围。");
            }
            html.Append("</table></body></html>");return new {text=plain.ToString(),html=html.ToString(),rows=lastRow-firstRow+1,columns=spec.Columns};
        }
    }
}
