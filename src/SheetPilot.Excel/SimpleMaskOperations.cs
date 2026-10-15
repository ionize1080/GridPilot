using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using Newtonsoft.Json;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class ExcelGateway
    {
        public SimpleMaskSnapshot CaptureSimpleMask(CancellationToken token,Action<string,int,long,long> progress=null)
        {
            CheckThread();dynamic app=application(),selection=app.Selection;
            if(selection==null||Microsoft.VisualBasic.Information.TypeName((object)selection).TrimStart('_')!="Range")throw new InvalidOperationException("请先选择单元格、整行或整列。");
            dynamic sheet=selection.Worksheet,book=sheet.Parent;
            if((int)app.ActiveWindow.SelectedSheets.Count!=1)throw new InvalidOperationException("请先取消工作表组合，再选择一个工作表。");
            var snapshot=new SimpleMaskSnapshot{BookKey=WorkbookKey(book),SheetKey=SheetKey(sheet),BookName=(string)book.Name,SheetName=(string)sheet.Name};
            var seen=new HashSet<string>();int read=0;
            if((int)selection.Areas.Count>100)throw new InvalidOperationException("最多支持 100 个不连续区域，请分批选择。");
            for(int a=1;a<=(int)selection.Areas.Count;a++)
            {
                token.ThrowIfCancellationRequested();dynamic area=selection.Areas.Item[a];var blockCells=new List<SimpleMaskCell>();
                // SpecialCells can expand a single-cell selection: always intersect back with the original area.
                foreach(int type in (bool)sheet.ProtectContents?new[]{0}:new[]{2,-4123})
                {
                    dynamic occupied=null;try{occupied=type==0?app.Intersect(area,sheet.UsedRange):app.Intersect(area,area.SpecialCells(type));}catch(COMException){continue;}
                    if(occupied==null)continue;
                    if(type==0&&Convert.ToDouble(occupied.CountLarge,CultureInfo.InvariantCulture)>1000000)throw new InvalidOperationException("受保护工作表的选区超过 100 万格，请缩小选区或解除保护后重试。");
                    if(type!=0&&Convert.ToDouble(occupied.CountLarge,CultureInfo.InvariantCulture)+snapshot.Cells.Count+blockCells.Count>100000)throw new InvalidOperationException("选区超过 100000 个有内容单元格，请分批处理；整行整列的空格不计入。");
                    for(int region=1;region<=(int)occupied.Areas.Count;region++)
                    {
                        dynamic part=occupied.Areas.Item[region];object raw=(object)part.Value2;var matrix=raw as object[,];
                        int width=(int)part.Columns.Count,height=(int)part.Rows.Count,startRow=(int)part.Row,startColumn=(int)part.Column;
                        object format=part.NumberFormat,merged=part.MergeCells,hasFormula=type==0?(object)part.HasFormula:type==-4123;
                        object locked=type==0?(object)part.Locked:false;
                        for(int row=0;row<height;row++)for(int col=0;col<width;col++)
                        {
                            token.ThrowIfCancellationRequested();object value=matrix==null?raw:matrix[row+1,col+1];
                            if(value==null&&Equals(hasFormula,false))continue;
                            int r=startRow+row,cx=startColumn+col;string columnName=DataCell.Address(1,cx);columnName=columnName.Substring(0,columnName.Length-1);
                            string address="$"+columnName+"$"+r;if(!seen.Add(address))continue;
                            if(snapshot.Cells.Count+blockCells.Count>=100000)throw new InvalidOperationException("选区超过 100000 个有内容单元格，请分批处理。");
                            var c=new SimpleMaskCell{Address=address,Row=r,Column=cx,Block=snapshot.Blocks.Count,Before=value};
                            try
                            {
                                // Uniform ranges need only a handful of COM calls; inspect individual cells only for mixed metadata.
                                dynamic cell=null;if(!(format is string)||!(merged is bool)||!(hasFormula is bool)||!(locked is bool))cell=sheet.Range[address];
                                bool formula=hasFormula is bool f?f:(bool)cell.HasFormula;
                                if(value==null&&!formula)continue;
                                c.NumberFormat=format as string??(string)cell.NumberFormat;
                                if(formula)c.Skip="公式";
                                else if(merged is bool m?m:(bool)cell.MergeCells)c.Skip="合并单元格";
                                else if(value is ErrorWrapper || value is int)c.Skip="Excel 错误值";
                                else if(value is bool)c.Skip="布尔值";
                                else if(string.IsNullOrWhiteSpace(SimpleMask.Text(value)))c.Skip="空白";
                                else if(type==0&&(locked is bool l?l:(bool)cell.Locked))c.Skip="受保护单元格";
                            }
                            catch(COMException){c.Skip="读取失败";c.Before=null;}
                            blockCells.Add(c);read++;if(read%128==0)progress?.Invoke("读取选区，可取消",0,read,0);
                        }
                    }
                }
                if(blockCells.Count==0)continue;
                int top=blockCells.Min(c=>c.Row),left=blockCells.Min(c=>c.Column),bottom=blockCells.Max(c=>c.Row),right=blockCells.Max(c=>c.Column);
                snapshot.Blocks.Add(new SimpleMaskBlock{Address=(string)area.Address,Row=top,Column=left,Rows=bottom-top+1,Columns=right-left+1});snapshot.Cells.AddRange(blockCells.OrderBy(c=>c.Row).ThenBy(c=>c.Column));
            }
            progress?.Invoke("读取完成",100,read,read);return snapshot;
        }
        public string ApplySimpleMask(SimpleMaskSnapshot source,SimpleMaskResult result,bool backupConfirmed,CancellationToken token,Action<string,int,long,long> progress=null)
        {
            CheckThread();if(!backupConfirmed)throw new InvalidOperationException("写回原表前请备份，并勾选备份确认。");
            dynamic app=application(),book=FindBook(source.BookKey),sheet=FindPage(book,source.SheetKey);
            if((bool)book.ReadOnly)throw new InvalidOperationException("工作簿只读，请复制或导出结果。");
            if((int)book.Windows.Count>0 && (int)book.Windows[1].SelectedSheets.Count>1)throw new InvalidOperationException("请先取消工作表组合，再写回原表。");
            bool events=app.EnableEvents,screen=app.ScreenUpdating;int applied=0,failed=0;bool canceled=false;
            var changes=new List<SimpleMaskCell>();
            try
            {
                app.EnableEvents=false;app.ScreenUpdating=false;
                foreach(var c in result.Cells.Where(c=>c.Skip==null))
                {
                    if(token.IsCancellationRequested){canceled=true;break;}
                    if(c.After==SimpleMask.Text(c.Before)){c.WriteStatus="无需修改";continue;}
                    dynamic cell=null;bool started=false;
                    try
                    {
                        cell=sheet.Range[c.Address];
                        if((bool)cell.HasFormula||(bool)cell.MergeCells||((bool)sheet.ProtectContents&&(bool)cell.Locked))throw new InvalidOperationException("公式、合并或保护状态已改变");
                        if(!Equals((object)cell.Value2,c.Before)||!Equals((string)cell.NumberFormat,c.NumberFormat))throw new InvalidOperationException("预览后单元格已改变");
                        started=true;cell.NumberFormat="@";cell.Value2=c.After;
                        if((bool)cell.HasFormula||!SimpleMaskValueEquals((object)cell.Value2,c.After))throw new InvalidOperationException("写入验证失败");
                        c.WriteStatus="已写入";changes.Add(c);applied++;
                    }
                    catch(Exception ex) when(ex is COMException || ex is InvalidOperationException)
                    {
                        c.WriteStatus=ex is InvalidOperationException?ex.Message:"Excel 拒绝写入";failed++;
                        if(started)try{cell.NumberFormat="@";cell.Value2=c.Before;cell.NumberFormat=c.NumberFormat;}catch(COMException){c.WriteStatus+="；恢复失败，请从备份核对";}
                    }
                    if((applied+failed)%128==0)progress?.Invoke("写回原表，可取消（已写入项保留）",0,applied+failed,0);
                }
                if(canceled)foreach(var c in result.Cells.Where(c=>c.Skip==null&&c.WriteStatus==null))c.WriteStatus="取消后未写入";
            }
            finally{try{app.EnableEvents=events;}finally{app.ScreenUpdating=screen;}}
            if(changes.Count>0)
            {
                WatchUndo();undoInvalidated=false;
                undo=()=>RestoreSimpleMask(source,changes);
            }
            progress?.Invoke("写回结束",100,applied+failed,applied+failed);
            return (canceled?"已取消，":"")+"已写入 "+applied+" 格；写入跳过 "+failed+" 格；规则跳过 "+result.Skipped+" 格。未自动保存工作簿；可用 GridPilot 撤销。";
        }
        private static bool SimpleMaskValueEquals(object actual,string expected) => expected.Length==0?actual==null||Equals(actual,""):Equals(actual,expected);
        private string RestoreSimpleMask(SimpleMaskSnapshot source,List<SimpleMaskCell> changes)
        {
            CheckThread();dynamic app=application(),book=FindBook(source.BookKey),sheet=FindPage(book,source.SheetKey);
            if((bool)book.ReadOnly || ((int)book.Windows.Count>0&&(int)book.Windows[1].SelectedSheets.Count>1))throw new InvalidOperationException("请取消工作表组合并确认工作簿可写。");
            // Verify every target before the first restoration to avoid overwriting intervening edits.
            foreach(var c in changes){dynamic cell=sheet.Range[c.Address];if((bool)cell.HasFormula||(bool)cell.MergeCells||((bool)sheet.ProtectContents&&(bool)cell.Locked)||!SimpleMaskValueEquals((object)cell.Value2,c.After)||(string)cell.NumberFormat!="@")throw new InvalidOperationException("脱敏后的单元格已改变，无法自动撤销。请使用备份。");}
            bool events=app.EnableEvents;int restored=0;try{app.EnableEvents=false;foreach(var c in changes){dynamic cell=sheet.Range[c.Address];cell.NumberFormat="@";cell.Value2=c.Before;cell.NumberFormat=c.NumberFormat;restored++;}}finally{app.EnableEvents=events;}
            return "已还原 "+restored+" 格及数字格式；未自动保存。";
        }
        public string ExportSimpleMask(SimpleMaskSnapshot source,SimpleMaskResult result,string path,CancellationToken token)
        {
            CheckThread();path=Path.GetFullPath(path);if(Path.GetExtension(path).ToLowerInvariant()!=".xlsx"||File.Exists(path))throw new InvalidOperationException("请选择尚不存在的 .xlsx 文件，避免覆盖已有数据。");
            foreach(var b in source.Blocks)if((long)b.Rows*b.Columns>1000000)throw new InvalidOperationException("结果区域含超过 100 万格，请缩小选择后导出。");
            dynamic app=application(),book=null;bool events=app.EnableEvents,screen=app.ScreenUpdating;
            string staging=Path.Combine(Path.GetDirectoryName(path),".GridPilot-"+Guid.NewGuid().ToString("N")+".xlsx");
            try
            {
                app.EnableEvents=false;app.ScreenUpdating=false;book=app.Workbooks.Add(-4167);
                for(int i=0;i<source.Blocks.Count;i++)
                {
                    token.ThrowIfCancellationRequested();var block=source.Blocks[i];dynamic sheet=i==0?book.Worksheets[1]:book.Worksheets.Add(Type.Missing,book.Worksheets[book.Worksheets.Count]);sheet.Name="脱敏结果"+(i+1);
                    var values=new object[block.Rows,block.Columns];foreach(var c in result.Cells.Where(c=>c.Block==i&&c.Skip==null))values[c.Row-block.Row,c.Column-block.Column]=c.After;
                    dynamic range=sheet.Range[sheet.Cells[1,1],sheet.Cells[block.Rows,block.Columns]];range.NumberFormat="@";range.Value2=values;sheet.Columns.ColumnWidth=20;
                }
                token.ThrowIfCancellationRequested();book.SaveAs(staging,51);book.Close(false);book=null;token.ThrowIfCancellationRequested();File.Move(staging,path);return path;
            }
            finally{try{if(book!=null)book.Close(false);}finally{try{if(File.Exists(staging))File.Delete(staging);}finally{app.EnableEvents=events;app.ScreenUpdating=screen;}}}
        }
        public string ExportSimpleRecovery(SimpleMaskSnapshot source,SimpleMaskResult result,string path)
        {
            path=Path.GetFullPath(path);if(Path.GetExtension(path).ToLowerInvariant()!=".json")throw new InvalidOperationException("还原记录应保存为 .json。");
            var record=new {schema="GridPilot.SimpleMask.Recovery/1",mappingVersion=SimpleMask.MappingVersion,createdUtc=DateTime.UtcNow,source.BookName,source.SheetName,source.Blocks,options=result.Options,letters=SimpleMask.Letters,digits=SimpleMask.Digits,cells=result.Cells.Select(c=>new{c.Address,c.Block,c.Row,c.Column,beforeType=c.Before is string?"text":c.Before is double?"number":"other",c.Before,c.After,c.NumberFormat,c.Skip,c.WriteStatus})};
            // CreateNew avoids replacing any user file, even if the selected path is the source workbook.
            using(var stream=new FileStream(path,FileMode.CreateNew,FileAccess.Write,FileShare.None))using(var writer=new StreamWriter(stream,new System.Text.UTF8Encoding(false)))writer.Write(JsonConvert.SerializeObject(record,Formatting.Indented));
            return path;
        }
    }
}
