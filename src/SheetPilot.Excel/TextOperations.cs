using System;
using System.Collections.Generic;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class ExcelGateway
    {
        public IReadOnlyList<TextInput> CaptureTextSelection()
        {
            CheckThread(); dynamic app=application(), selected=app.Selection;
            if(selected==null || Microsoft.VisualBasic.Information.TypeName((object)selected).TrimStart('_')!="Range") throw new InvalidOperationException("请先选择需要分词的单元格。");
            var result=new List<TextInput>(); long total=0; var seen=new HashSet<string>();
            for(int a=1;a<=(int)selected.Areas.Count;a++)
            {
                dynamic range=selected.Areas.Item[a]; total+=(long)(double)range.CountLarge;
                if(total>20000) throw new InvalidOperationException("一次最多读取 20000 个单元格，请缩小选区。");
                object raw=range.Value2; var values=raw as object[,];int rowCount=range.Rows.Count,columnCount=range.Columns.Count;
                for(int r=1;r<=rowCount;r++) for(int c=1;c<=columnCount;c++)
                {
                    object value=values==null?raw:values[r,c]; if(!(value is string text) || string.IsNullOrWhiteSpace(text)) continue;
                    dynamic cell=range.Cells[r,c]; if((bool)cell.MergeCells) continue;
                    string address=cell.Address; if(!seen.Add(address)) continue;
                    result.Add(new TextInput { Sheet=(string)range.Worksheet.Name,Address=address,Text=text });
                }
            }
            if(result.Count==0) throw new InvalidOperationException("所选区域没有可分词文本。");
            return result;
        }
        public void WriteTextOutput(object[,] values)
        {
            CheckThread(); dynamic app=application(), book=null; bool events=app.EnableEvents;
            try
            {
                app.EnableEvents=false; book=app.Workbooks.Add(-4167); dynamic sheet=book.Worksheets[1]; sheet.Name="分词结果";
                dynamic range=sheet.Range[sheet.Cells[1,1],sheet.Cells[values.GetLength(0),values.GetLength(1)]];
                range.NumberFormat="@"; range.Value2=values;
                sheet.Rows[1].Font.Bold=true; sheet.Columns.ColumnWidth=20; sheet.Columns[4].ColumnWidth=55;
            }
            catch { if(book!=null) book.Close(false); throw; }
            finally { app.EnableEvents=events; }
        }
    }
}
