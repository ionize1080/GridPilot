using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace SheetPilot.Core
{
    public sealed class TextInput
    {
        public string Sheet { get; set; }
        public string Address { get; set; }
        public string Text { get; set; }
    }
    public sealed class TextOutput
    {
        public TextInput Source { get; set; }
        public IReadOnlyList<WordToken> Words { get; set; }
        public string Preview => string.Join(" / ",Words.Select(w=>w.Text));
    }
    public static class TextBatch
    {
        public static object[,] Layout(IReadOnlyList<TextOutput> rows, SegmentationOptions options)
        {
            int count=options.OutputMode==2?rows.Sum(r=>r.Words.Count):rows.Count;
            int columns=options.OutputMode==1?3+rows.Select(r=>r.Words.Count).DefaultIfEmpty().Max():options.OutputMode==2?5:4;
            if(count+1>1048576 || columns>16384 || (long)(count+1)*columns>2000000) throw new InvalidOperationException("结果超过行列或 200 万单元格上限，请缩小选择范围。");
            var values=new object[count+1,columns];
            values[0,0]="来源工作表"; values[0,1]="来源单元格"; values[0,2]="词数";
            if(options.OutputMode==2) { values[0,2]="词序"; values[0,3]="字符位置（从1开始）"; values[0,4]="词"; }
            else for(int j=3;j<columns;j++) values[0,j]=options.OutputMode==1?"词"+(j-2):"分词结果";
            int row=1;
            foreach(var entry in rows)
            {
                if(options.OutputMode==2)
                { for(int j=0;j<entry.Words.Count;j++) { values[row,0]=entry.Source.Sheet; values[row,1]=entry.Source.Address; values[row,2]=j+1; values[row,3]=entry.Words[j].Offset+1; values[row++,4]=entry.Words[j].Text; } }
                else
                {
                    values[row,0]=entry.Source.Sheet; values[row,1]=entry.Source.Address; values[row,2]=entry.Words.Count;
                    if(options.OutputMode==1) for(int j=0;j<entry.Words.Count;j++) values[row,j+3]=entry.Words[j].Text;
                    else { string text=string.Join(options.Separator??" / ",entry.Words.Select(w=>w.Text)); if(text.Length>32767) throw new InvalidOperationException("合并后的分词结果超过单元格文字上限，请改用纵向明细。"); values[row,3]=text; }
                    row++;
                }
            }
            return values;
        }
    }
}
