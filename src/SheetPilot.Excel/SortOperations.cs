using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed class SortSnapshot
    {
        internal string Book,Sheet,Address,Fingerprint,Table;
        internal int Row,Column,Rows,Columns;
        internal object[,] Values;
        public string Description { get; internal set; }
        public string[] Headers { get; internal set; }
        public int InitialColumn { get; internal set; }
        public IReadOnlyList<string> Keys(int column,bool header)
        {
            if(column<0||column>=Columns)throw new InvalidOperationException("请选择排序依据列。");
            if(Table!=null&&!header)throw new InvalidOperationException("Excel 表格的标题行必须保留。");
            return Enumerable.Range(header?1:0,Rows-(header?1:0)).Select(i=>Convert.ToString(Values[i,column],CultureInfo.InvariantCulture)??"").ToArray();
        }
    }
    public sealed partial class ExcelGateway
    {
        private static object[,] SortMatrix(object raw,int rows,int columns)
        {
            var result=new object[rows,columns];if(raw is Array a){for(int r=0;r<rows;r++)for(int c=0;c<columns;c++)result[r,c]=a.GetValue(r+a.GetLowerBound(0),c+a.GetLowerBound(1));}else result[0,0]=raw;return result;
        }
        private static string SortFingerprint(dynamic range) => JsonConvert.SerializeObject(new object[]{(object)range.Formula,(object)range.Value2});
        public SortSnapshot CaptureSort()
        {
            CheckThread();dynamic app=application(),selected=app.Selection;
            if(selected==null||Microsoft.VisualBasic.Information.TypeName((object)selected).TrimStart('_')!="Range"||(int)selected.Areas.Count!=1)throw new InvalidOperationException("请在连续数据区域内选中一个单元格。");
            dynamic sheet=selected.Worksheet,book=sheet.Parent;
            if((bool)book.ReadOnly||(bool)sheet.ProtectContents)throw new InvalidOperationException("工作簿只读或工作表受保护，无法排序。");
            dynamic range=selected.CurrentRegion;
            dynamic table=selected.Cells[1,1].ListObject;
            if(table!=null)range=table.Range;
            else if((bool)sheet.AutoFilterMode)
            {dynamic filter=sheet.AutoFilter;if(filter!=null&&app.Intersect(selected,filter.Range)!=null)range=filter.Range;}
            if(table!=null&&(bool)table.ShowTotals)range=range.Resize[(int)range.Rows.Count-1,(int)range.Columns.Count];
            int rows=range.Rows.Count,columns=range.Columns.Count;
            if(rows<2||rows>200001||(long)rows*columns>2000000)throw new InvalidOperationException("排序区域需有至少 2 行，最多 200,001 行 / 2,000,000 个单元格。");
            ValidateSortRange(book,sheet,range);
            var values=SortMatrix((object)range.Value2,rows,columns);
            return new SortSnapshot {Book=WorkbookKey(book),Sheet=SheetKey(sheet),Address=range.Address,Row=range.Row,Column=range.Column,Rows=rows,Columns=columns,Table=table==null?null:(string)table.Name,
                Values=values,Fingerprint=SortFingerprint(range),InitialColumn=Math.Max(0,Math.Min(columns-1,(int)app.ActiveCell.Column-(int)range.Column)),
                Headers=Enumerable.Range(0,columns).Select(c=>"第 "+(c+1)+" 列 · "+Convert.ToString(values[0,c],CultureInfo.InvariantCulture)).ToArray(),
                Description=(string)book.Name+" / "+(string)sheet.Name+" / "+(string)range.Address+" · "+rows+" 行 × "+columns+" 列。自动包含整个连续区域或筛选表；包含隐藏行。"};
        }
        private static void ValidateSortRange(dynamic book,dynamic sheet,dynamic range)
        {
            if((bool)book.ReadOnly||(bool)sheet.ProtectContents)throw new InvalidOperationException("工作簿只读或工作表受保护，无法排序。");
            if(!Equals((object)range.MergeCells,false))throw new InvalidOperationException("排序区域包含合并单元格，请先在副本中取消合并。");
            if(!Equals((object)range.HasArray,false))throw new InvalidOperationException("排序区域包含数组公式，无法安全排序。");
            try{if(!Equals((object)range.HasSpill,false))throw new InvalidOperationException("排序区域包含动态数组，请先复制为值。");}
            catch(System.Runtime.InteropServices.COMException){/* Older Excel has no HasSpill. */}
            catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){ }
            if((int)sheet.PivotTables().Count>0)
                for(int i=1;i<=(int)sheet.PivotTables().Count;i++)
                    if(sheet.Application.Intersect(range,sheet.PivotTables(i).TableRange2)!=null)throw new InvalidOperationException("请在数据透视表的源数据中排序。");
        }
        public string ApplySort(SortSnapshot snapshot,int column,SortOptions options)
        {
            CheckThread();var order=CustomSort.Order(snapshot.Keys(column,options.HasHeader),options);
            if(order.Select((value,i)=>value==i).All(x=>x))return "顺序已符合规则，无需调整。";
            dynamic book=FindBook(snapshot.Book),sheet=FindPage(book,snapshot.Sheet),range=sheet.Range[snapshot.Address];
            if(SortFingerprint(range)!=snapshot.Fingerprint)throw new InvalidOperationException("数据在预览后发生变化，请关闭窗口重新读取。");
            ValidateSortRange(book,sheet,range);WatchUndo();
            var ranks=new int[order.Length];for(int i=0;i<order.Length;i++)ranks[order[i]]=i;
            SortRows(snapshot,options.HasHeader,ranks);
            string after=SortFingerprint(range);undoInvalidated=false;
            undo=()=>
            {
                dynamic b=FindBook(snapshot.Book),s=FindPage(b,snapshot.Sheet),live=s.Range[snapshot.Address];
                if(SortFingerprint(live)!=after)throw new InvalidOperationException("排序后的内容已改变，不能恢复原序。");
                ValidateSortRange(b,s,live);SortRows(snapshot,options.HasHeader,order);return "已恢复排序前的记录顺序。";
            };
            return "已排序 "+order.Length+" 条记录，保留公式和单元格格式。可通过 GridPilot 撤销；未自动保存。";
        }
        private void SortRows(SortSnapshot snapshot,bool header,int[] ranks)
        {
            dynamic app=application(),book=FindBook(snapshot.Book),sheet=FindPage(book,snapshot.Sheet),range=sheet.Range[snapshot.Address];
            ValidateSortRange(book,sheet,range);
            if(snapshot.Column+snapshot.Columns>(int)sheet.Columns.Count)throw new InvalidOperationException("数据已到最后一列，缺少临时排序键空间。");
            dynamic helper=sheet.Cells[snapshot.Row,snapshot.Column+snapshot.Columns].Resize[snapshot.Rows+(snapshot.Table!=null&&(bool)sheet.ListObjects[snapshot.Table].ShowTotals?1:0),1];
            if((double)app.WorksheetFunction.CountA(helper)!=0||!Equals((object)helper.MergeCells,false))throw new InvalidOperationException("数据右侧紧邻列不是空白，请先留出一列再排序。");
            dynamic table=snapshot.Table==null?null:sheet.ListObjects[snapshot.Table],added=null;
            if(table!=null&&(!header||(string)table.Range.Resize[snapshot.Rows,snapshot.Columns].Address!=snapshot.Address))throw new InvalidOperationException("Excel 表格范围已改变，请重新读取。");
            bool events=app.EnableEvents,screen=app.ScreenUpdating;object autoExpand=null;
            dynamic sort=table==null?sheet.Sort:table.Sort;
            try
            {
                app.EnableEvents=false;app.ScreenUpdating=false;
                // Keep Excel from absorbing a temporary key into a neighbouring table.
                autoExpand=app.AutoCorrect.AutoExpandListRange;app.AutoCorrect.AutoExpandListRange=false;
                if(table!=null){added=table.ListColumns.Add();helper=added.Range;}
                var keys=new object[ranks.Length,1];for(int i=0;i<ranks.Length;i++)keys[i,0]=ranks[i];
                dynamic key=sheet.Cells[snapshot.Row+(header?1:0),snapshot.Column+snapshot.Columns].Resize[ranks.Length,1];key.Value2=keys;
                if(header)helper.Cells[1,1].Value2="__GridPilotSort";
                sort.SortFields.Clear();sort.SortFields.Add(key,0,1,Type.Missing,0);
                if(table==null)sort.SetRange(range.Resize[snapshot.Rows,snapshot.Columns+1]);
                sort.Header=header?1:2;sort.MatchCase=false;sort.Orientation=1;sort.Apply();
            }
            finally
            {
                try{sort.SortFields.Clear();if(added!=null)added.Delete();else helper.ClearContents();}
                finally{try{if(autoExpand!=null)app.AutoCorrect.AutoExpandListRange=autoExpand;}finally{try{app.EnableEvents=events;}finally{app.ScreenUpdating=screen;}}}
            }
        }
        public string SortSheets(IReadOnlyList<SheetItem> expected,SortOptions options)
        {
            CheckThread();if(expected.Count==0)return "没有工作表。";
            var order=CustomSort.Order(expected.Select(s=>s.Name).ToArray(),options);
            var desired=order.Select(i=>expected[i]).ToArray();ReorderSheets(expected,desired);
            undoInvalidated=false;undo=()=>{ReorderSheets(desired,expected);return "已恢复工作表原始顺序。";};return "已排序 "+expected.Count+" 个工作表标签；可通过 GridPilot 撤销。";
        }
        private void ReorderSheets(IReadOnlyList<SheetItem> expected,IReadOnlyList<SheetItem> desired)
        {
            dynamic book=FindBook(expected[0].WorkbookKey),app=application();
            if((bool)book.ReadOnly||(bool)book.ProtectStructure)throw new InvalidOperationException("工作簿只读或结构已保护。");
            var live=Capture(true).Sheets.Where(s=>s.WorkbookKey==expected[0].WorkbookKey).ToArray();
            if(live.Length!=expected.Count||live.Where((s,i)=>s.Key!=expected[i].Key||s.Name!=expected[i].Name||s.Visibility!=expected[i].Visibility).Any())throw new InvalidOperationException("工作表列表在预览后已变化，请重新读取。");
            bool screen=app.ScreenUpdating,events=app.EnableEvents;dynamic active=app.ActiveSheet;
            Action<IReadOnlyList<SheetItem>> move=items=>{for(int i=items.Count-1;i>=0;i--){dynamic page=FindPage(book,items[i].Key);if((int)page.Index!=1)page.Move(book.Sheets[1]);}};
            try{app.ScreenUpdating=false;app.EnableEvents=false;
                // Excel ignores a hidden Before anchor. Make anchors visible only while moving,
                // then restore Hidden / VeryHidden exactly before repainting.
                foreach(var item in expected)if(!item.CanNavigate)((dynamic)FindPage(book,item.Key)).Visible=-1;
                try{move(desired);}catch{move(expected);throw;}}
            finally{try{foreach(var item in expected)if(!item.CanNavigate)((dynamic)FindPage(book,item.Key)).Visible=item.Visibility==SheetVisibility.VeryHidden?2:0;
                active?.Activate();}finally{app.EnableEvents=events;app.ScreenUpdating=screen;}}
        }
    }
}
