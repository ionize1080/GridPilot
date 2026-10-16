using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace SheetPilot.Core
{
    [DataContract]
    public sealed class SortLevel
    {
        [DataMember] public int Column { get; set; }
        [DataMember] public string Mode { get; set; } = "values";
        [DataMember] public bool Descending { get; set; }
        [DataMember] public bool TextAsNumbers { get; set; }
        [DataMember] public string Rules { get; set; } = "[中文数字]";
        [DataMember] public string Sequences { get; set; } = "";
        public SortOptions Custom(bool matchCase) => new SortOptions { Natural=Mode=="natural", Descending=Descending, MatchCase=matchCase, Rules=Rules, Sequences=Sequences };
    }

    [DataContract]
    public sealed class MultiSortOptions
    {
        public const int MaxLevels = 20;
        public const int PreviewLimit = 20000;
        [DataMember] public bool HasHeader { get; set; } = true;
        [DataMember] public bool MatchCase { get; set; }
        [DataMember] public bool StrokeOrder { get; set; }
        [DataMember] public List<SortLevel> Levels { get; set; } = new List<SortLevel> { new SortLevel() };
        [OnDeserializing] private void Initialize(StreamingContext context) { HasHeader=true;Levels=new List<SortLevel>(); }
        public void Validate(int columns)
        {
            if(Levels==null||Levels.Count<1||Levels.Count>MaxLevels)throw new InvalidOperationException("请设置 1–20 个排序层级。");
            foreach(var level in Levels)
            {
                if(level==null||level.Column<0||level.Column>=columns)throw new InvalidOperationException("排序列超出当前区域，请重新选择。");
                if(!new[]{"values","natural","custom"}.Contains(level.Mode))throw new InvalidOperationException("未知排序方式。");
            }
        }
    }

    // Merge-connected rows form one record. A permutation must never split a record.
    public static class SortBlocks
    {
        public static int[][] Create(int rows, IEnumerable<int[]> merges)
        {
            var ends=Enumerable.Range(0,rows).ToArray();
            foreach(var m in merges)
            {
                if(m.Length<4||m[0]<0||m[2]<1||m[0]+m[2]>rows)throw new InvalidOperationException("合并区域跨越表头或排序边界，请选择完整记录。");
                ends[m[0]]=Math.Max(ends[m[0]],m[0]+m[2]-1);
            }
            var blocks=new List<int[]>();
            for(int start=0;start<rows;)
            {
                int end=ends[start];for(int i=start+1;i<=end;i++)end=Math.Max(end,ends[i]);
                blocks.Add(Enumerable.Range(start,end-start+1).ToArray());start=end+1;
            }
            return blocks.ToArray();
        }
        public static int[] Expand(int[][] blocks,IReadOnlyList<int> order) => order.SelectMany(i=>blocks[i]).ToArray();
        public static int[] Inverse(IReadOnlyList<int> order)
        {
            if(order==null||order.Distinct().Count()!=order.Count||order.Any(i=>i<0||i>=order.Count))throw new InvalidOperationException("无效排序计划。");
            var ranks=new int[order.Count];for(int i=0;i<order.Count;i++)ranks[order[i]]=i;return ranks;
        }
    }
}
