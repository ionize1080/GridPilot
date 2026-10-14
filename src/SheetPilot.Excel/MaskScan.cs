using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class ExcelGateway
    {
        // Adaptive rectangles: Find trims empty margins at every split. No per-cell Find,
        // no UsedRange inflation, no dependence on the user's current Find settings.
        private static void ScanMaskSheet(object sheetObject, int sheetIndex, CancellationToken cancel, Action<DataCell> visit, Action<long> progress = null, MaskComLifetime owned = null)
        {
            dynamic sheet=sheetObject;
            int rowCount,columnCount;
            using(var scope=owned?.Scope())
            {dynamic rows=sheet.Rows,columns=sheet.Columns;if(owned!=null){owned.Hold((object)rows);owned.Hold((object)columns);}rowCount=rows.Count;columnCount=columns.Count;}
            var pending=new Stack<CellRect>();pending.Push(new CellRect(1,1,rowCount,columnCount));
            long count=0;
            while(pending.Count>0)
            {
                cancel.ThrowIfCancellationRequested();var rect=pending.Pop();
                using(var scope=owned?.Scope())
                {
                    Func<object,dynamic> hold=value=>owned==null?value:owned.Hold(value);
                    dynamic region=hold(sheet.Range[rect.Address]);
                    if(rect.Count>32768)
                    {
                        dynamic firstAfter=hold(sheet.Range[DataCell.Address(rect.Bottom,rect.Right)]),lastAfter=hold(sheet.Range[DataCell.Address(rect.Row,rect.Column)]);
                        dynamic first=hold(region.Find("*",firstAfter,-4123,2,1,1,false,false,false));if(first==null)continue;
                        dynamic last=hold(region.Find("*",lastAfter,-4123,2,1,2,false,false,false));
                        dynamic left=hold(region.Find("*",firstAfter,-4123,2,2,1,false,false,false));
                        dynamic right=hold(region.Find("*",lastAfter,-4123,2,2,2,false,false,false));
                        rect=new CellRect((int)first.Row,(int)left.Column,(int)last.Row-(int)first.Row+1,(int)right.Column-(int)left.Column+1);
                        if(rect.Count>32768)
                        {
                            if(rect.Rows>=rect.Columns)
                            { int n=rect.Rows/2;pending.Push(new CellRect(rect.Row+n,rect.Column,rect.Rows-n,rect.Columns));pending.Push(new CellRect(rect.Row,rect.Column,n,rect.Columns)); }
                            else
                            { int n=rect.Columns/2;pending.Push(new CellRect(rect.Row,rect.Column+n,rect.Rows,rect.Columns-n));pending.Push(new CellRect(rect.Row,rect.Column,rect.Rows,n)); }
                            continue;
                        }
                        region=hold(sheet.Range[rect.Address]);
                    }
                    object raw=region.Value2,rawFormula=region.Formula;var values=raw as object[,];var formulas=rawFormula as object[,];
                    object hasFormula=region.HasFormula;
                    for(int r=0;r<rect.Rows;r++)for(int c=0;c<rect.Columns;c++)
                    {
                        if((c&255)==0)cancel.ThrowIfCancellationRequested();
                        object value=values==null?raw:values[r+1,c+1],f=formulas==null?rawFormula:formulas[r+1,c+1];
                        if(value==null && (f==null || Equals(f,"")))continue;
                        int row=rect.Row+r,column=rect.Column+c;
                        if(value is ErrorWrapper || value is int)throw new InvalidOperationException("S"+(sheetIndex+1)+"!"+DataCell.Address(row,column)+"：数据含 Excel 错误值，请先处理。");
                        string formula=null;
                        if(f is string text && text.StartsWith("=",StringComparison.Ordinal) && !Equals(hasFormula,false))
                        {
                            dynamic cell=hold(sheet.Range[DataCell.Address(row,column)]);
                            if((bool)cell.HasFormula)
                            { if((bool)cell.HasArray)throw new InvalidOperationException("数组公式不在支持范围，未导出。");formula=text; }
                        }
                        visit(new DataCell {Sheet=sheetIndex,Row=row,Column=column,Value=value,Formula=formula});count++;
                    }
                    progress?.Invoke(count);
                }
            }
        }
        private static int CompareMaskCell(DataCell a,DataCell b)
        {int s=a.Sheet.CompareTo(b.Sheet);if(s!=0)return s;int r=a.Row.CompareTo(b.Row);return r!=0?r:a.Column.CompareTo(b.Column);}
        private static DataCell FindMaskCell(List<DataCell> cells,DataCell cell)
        {int index=cells.BinarySearch(cell,MaskCellComparer.Instance);return index<0?null:cells[index];}
        private sealed class MaskCellComparer : IComparer<DataCell>
        { internal static readonly MaskCellComparer Instance=new MaskCellComparer();public int Compare(DataCell a,DataCell b)=>CompareMaskCell(a,b); }
    }
}
