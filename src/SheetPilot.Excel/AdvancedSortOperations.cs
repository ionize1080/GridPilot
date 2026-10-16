using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed class SortPlan
    {
        internal int[] Order;
        internal bool Header;
        internal string SourceFingerprint;
        public int Records { get; internal set; }
        public int Rows => Order.Length;
        public int[] Matched { get; internal set; }
    }

    public sealed partial class ExcelGateway
    {
        // Use Excel itself for native values: numbers, dates, numeric text, logical
        // values, errors, blanks, locale, pinyin/stroke and case rules stay native.
        // Only detached keys are copied; no formula is evaluated in the scratch book.
        public SortPlan PlanSort(SortSnapshot snapshot,MultiSortOptions options)
        {
            CheckThread();options.Validate(snapshot.Columns);
            if(snapshot.Table!=null&&options.HasHeader!=snapshot.TableHasHeader)throw new InvalidOperationException("Excel 表格的标题行选项须与实际布局一致。");
            VerifySortSnapshot(snapshot);
            int skip=options.HasHeader?1:0,count=snapshot.Rows-skip;
            var merges=snapshot.Merges.Where(m=>m[0]+m[2]>skip).Select(m=>new[]{m[0]-skip,m[1],m[2],m[3]}).ToArray();
            int[][] blocks=SortBlocks.Create(count,merges);int n=blocks.Length,levels=options.Levels.Count;
            var data=new object[n,levels+1];var matched=new int[levels];
            for(int c=0;c<levels;c++)
            {
                var level=options.Levels[c];
                var values=blocks.Select(b=>snapshot.Values[b[0]+skip,level.Column]).ToArray();
                if(level.Mode=="values")
                {
                    matched[c]=n;
                    for(int r=0;r<n;r++)data[r,c]=values[r] is string s?"'"+s:values[r];
                }
                else
                {
                    var ranks=CustomSort.Ranks(values.Select(v=>Convert.ToString(v,CultureInfo.InvariantCulture)??"").ToArray(),level.Custom(options.MatchCase),out matched[c]);
                    for(int r=0;r<n;r++)data[r,c]=ranks[r];
                }
            }
            for(int r=0;r<n;r++)data[r,levels]=r;
            dynamic app=application(),scratch=null,active=app.ActiveSheet,selection=app.Selection;
            bool screen=app.ScreenUpdating,events=app.EnableEvents;
            var com=new MaskComLifetime();
            try
            {
                app.ScreenUpdating=false;app.EnableEvents=false;
                // Do not final-release the host Workbooks collection: other features own it.
                scratch=com.Hold(app.Workbooks.Add(-4167));
                dynamic sheet=com.Hold(scratch.Worksheets[1]),range=com.Hold(sheet.Cells[1,1].Resize[n,levels+1]);
                range.Value2=data;dynamic sort=com.Hold(sheet.Sort);sort.SortFields.Clear();
                for(int c=0;c<levels;c++)
                {
                    var level=options.Levels[c];dynamic key=com.Hold(sheet.Cells[1,c+1].Resize[n,1]);
                    sort.SortFields.Add(key,0,level.Mode=="values"&&level.Descending?2:1,Type.Missing,level.Mode=="values"&&level.TextAsNumbers?1:0);
                }
                sort.SortFields.Add(sheet.Cells[1,levels+1].Resize[n,1],0,1,Type.Missing,0);
                sort.SetRange(range);sort.Header=2;sort.Orientation=1;sort.MatchCase=options.MatchCase;sort.SortMethod=options.StrokeOrder?2:1;sort.Apply();
                var ids=SortMatrix((object)sheet.Cells[1,levels+1].Resize[n,1].Value2,n,1);
                var order=SortBlocks.Expand(blocks,Enumerable.Range(0,n).Select(i=>Convert.ToInt32(ids[i,0])).ToArray());
                return new SortPlan {Order=order,Header=options.HasHeader,Records=n,Matched=matched,SourceFingerprint=snapshot.Fingerprint};
            }
            finally
            {
                try{if(scratch!=null)scratch.Close(false);active?.Activate();try{selection?.Select();}catch(COMException){}}
                finally{app.EnableEvents=events;app.ScreenUpdating=screen;com.Dispose();}
            }
        }
        private void VerifySortSnapshot(SortSnapshot snapshot)
        {
            dynamic book=FindBook(snapshot.Book),sheet=FindPage(book,snapshot.Sheet),range=sheet.Range[snapshot.Address];
            if(SortFingerprint(range)!=snapshot.Fingerprint||JsonConvert.SerializeObject(RecordMerges(range))!=JsonConvert.SerializeObject(snapshot.Merges))
                throw new InvalidOperationException("数据、公式或合并布局在读取后已改变，请重新读取并预览。");
            ValidateSortRange(book,sheet,range);
        }
        public object SortPage(SortSnapshot snapshot,SortPlan plan,int offset,int size,string search="")
        {
            int skip=plan.Header?1:0;size=Math.Max(25,Math.Min(200,size));
            var rows=plan.Order.Select((source,i)=>new{source,index=i}).Take(MultiSortOptions.PreviewLimit);
            if(!string.IsNullOrEmpty(search))rows=rows.Where(r=>Enumerable.Range(0,snapshot.Columns).Any(c=>SortText(snapshot.Values[r.source+skip,c]).IndexOf(search,StringComparison.OrdinalIgnoreCase)>=0));
            var selected=rows.ToArray();offset=Math.Max(0,Math.Min(offset,Math.Max(0,selected.Length-1)));offset=offset/size*size;
            return new {total=plan.Rows,records=plan.Records,previewed=Math.Min(plan.Rows,MultiSortOptions.PreviewLimit),filtered=selected.Length,offset,size,matched=plan.Matched,
                headers=Enumerable.Range(0,snapshot.Columns).Select(c=>plan.Header?SortText(snapshot.Values[0,c]):"第 "+(c+1)+" 列"),
                rows=selected.Skip(offset).Take(size).Select(r=>new {position=r.index+1,sourceRow=snapshot.Row+skip+r.source,values=Enumerable.Range(0,snapshot.Columns).Select(c=>SortDisplay(snapshot,r.source+skip,c)).ToArray()}).ToArray()};
        }
        private static string SortDisplay(SortSnapshot snapshot,int row,int column)
        {
            object value=snapshot.Values[row,column];string format=snapshot.Formats?[column];
            if(value is double number&&format!=null)
            {
                var plain=System.Text.RegularExpressions.Regex.Replace(format.ToLowerInvariant(),"\"[^\"]*\"|\\\\.|\\[[^\\]]*\\]","");
                if(plain.IndexOfAny(new[]{'y','d'})>=0&&number>=0&&number<2957000)
                {try{return DateTime.FromOADate(number+(snapshot.Date1904?1462:0)).ToString(plain.Contains("h")?"yyyy-MM-dd HH:mm:ss":"yyyy-MM-dd",CultureInfo.InvariantCulture);}catch(ArgumentException){}}
            }
            return SortText(value);
        }
        private static string SortText(object value)
        {
            if(value is ErrorWrapper error){switch(error.ErrorCode & 0xFFFF){case 2000:return "#NULL!";case 2007:return "#DIV/0!";case 2015:return "#VALUE!";case 2023:return "#REF!";case 2029:return "#NAME?";case 2036:return "#NUM!";case 2042:return "#N/A";default:return "#Excel错误("+error.ErrorCode+")";}}
            return Convert.ToString(value,CultureInfo.InvariantCulture)??"";
        }
        public string ApplySort(SortSnapshot snapshot,SortPlan plan)
        {
            CheckThread();VerifySortSnapshot(snapshot);
            if(plan.SourceFingerprint!=snapshot.Fingerprint||plan.Rows!=snapshot.Rows-(plan.Header?1:0))throw new InvalidOperationException("预览与数据区域不一致，请重新预览。");
            var ranks=SortBlocks.Inverse(plan.Order);
            if(plan.Order.Select((v,i)=>v==i).All(x=>x))return "顺序已符合规则，无需调整。";
            WatchUndo();SortRows(snapshot,plan.Header,ranks);
            dynamic book=FindBook(snapshot.Book),sheet=FindPage(book,snapshot.Sheet),range=sheet.Range[snapshot.Address];
            string after=SortFingerprint(range),mergeAfter=JsonConvert.SerializeObject(RecordMerges(range));undoInvalidated=false;
            undo=()=>
            {
                dynamic b=FindBook(snapshot.Book),s=FindPage(b,snapshot.Sheet),live=s.Range[snapshot.Address];
                if(SortFingerprint(live)!=after||JsonConvert.SerializeObject(RecordMerges(live))!=mergeAfter)throw new InvalidOperationException("排序后的内容或合并布局已改变，不能恢复原序。");
                ValidateSortRange(b,s,live);SortRows(snapshot,plan.Header,plan.Order);return "已恢复排序前的记录顺序。";
            };
            return "已排序 "+plan.Records+" 条记录（"+plan.Rows+" 行），保留公式、格式和合并。可恢复原序；未自动保存。";
        }
        public void OpenNativeSort()
        {
            CheckThread();dynamic app=application();WorksheetFocus.Restore(new IntPtr((int)app.Hwnd));app.Dialogs[39].Show();
        }
        public void InsertSortExample()
        {
            CheckThread();dynamic app=application(),book=app.Workbooks.Add(-4167),sheet=book.Worksheets[1];sheet.Name="多级排序示例";
            sheet.Range["A1:E13"].Value2=new object[,] {
                {"所属部门","客户名称","项目名称","金额","日期"},
                {"业务二部","海川","项目十",120,46200},{"业务一部","远山","项目二",80,46190},
                {"业务二部","海川","项目一",300,46205},{"业务一部","远山","项目十一",60,46201},
                {"国际一部","青禾","项目三",200,46195},{"业务一部","青禾","项目一",40,46210},
                {"国际一部","青禾","项目二",150,46200},{"业务二部","远山","项目九",90,46204},
                {"国际二部","海川","项目二十",420,46209},{"业务一部","青禾","项目十",70,46206},
                {"国际二部","海川","项目二",310,46203},{"业务二部","海川","项目三",100,46208}};
            sheet.Range["A:E"].ColumnWidth=19;sheet.Range["A1:E1"].Font.Bold=true;sheet.Range["A1:E1"].Interior.Color=14807768;
            sheet.Range["D2:D13"].NumberFormat="#,##0.00";sheet.Range["E2:E13"].NumberFormat="yyyy-mm-dd";sheet.Range["A2"].Select();
        }
    }
}
