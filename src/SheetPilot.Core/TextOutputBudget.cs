using System;
namespace SheetPilot.Core
{
    public sealed class TextOutputBudget
    {
        private readonly SegmentationOptions options;
        private readonly int rows;
        private long total, count, length;
        public TextOutputBudget(SegmentationOptions options,int rows){this.options=options;this.rows=rows;}
        public void BeginRow(){count=length=0;}
        public void Accept(string word)
        {
            length+=word.Length+(count==0?0:(options.Separator??" / ").Length);count++;total++;
            if(word.Length>32767 || (options.OutputMode==0 && length>32767))throw new InvalidOperationException("分词输出超过单元格文字上限，请改用纵向明细。");
            long cells=options.OutputMode==2?(total+1)*5:options.OutputMode==1?(rows+1L)*(3+count):(rows+1L)*4;
            if(cells>2000000 || (options.OutputMode==1 && count+3>16384) || (options.OutputMode==2 && total+1>1048576))throw new InvalidOperationException("分词输出超过行列或200万格预算，已提前停止。");
        }
    }
}
