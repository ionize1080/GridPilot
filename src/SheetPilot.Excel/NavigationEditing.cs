using System;
using System.IO;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class ExcelGateway
    {
        public object OpenReportSource(string path,string sheetName,int row)
        {
            CheckThread();path=Path.GetFullPath(path);if(!File.Exists(path))throw new InvalidOperationException("源文件不存在，请重新扫描。");
            if(row<1||row>1048576)throw new InvalidOperationException("行号无效。");
            dynamic app=application(),book=null,books=app.Workbooks;
            for(int i=1;i<=books.Count;i++){dynamic candidate=books[i];if(string.Equals((string)candidate.FullName,path,StringComparison.OrdinalIgnoreCase)){book=candidate;break;}}
            if(book==null)
            {
                int previous=app.AutomationSecurity;bool events=app.EnableEvents;
                try{app.AutomationSecurity=3;app.EnableEvents=false;book=books.Open(path,0,false);}
                finally{app.AutomationSecurity=previous;app.EnableEvents=events;}
            }
            book.Activate();dynamic sheet=book.Worksheets[sheetName];
            if((int)sheet.Visible!=-1)throw new InvalidOperationException("源工作表已隐藏，请在 Excel 中取消隐藏。");
            sheet.Activate();sheet.Range["A"+row].Select();return new {message="已定位源 Excel；请修改并保存，再扫描更新。"};
        }
        public void EditSheet(SheetEdit request)
        {
            CheckThread();if(request?.Source==null)throw new InvalidOperationException("请先选择工作表。");
            dynamic app=application(),book=null,sheet=null,target=null;
            for(int i=1;i<=app.Workbooks.Count;i++){dynamic b=app.Workbooks[i];if(WorkbookKey(b)==request.Source.WorkbookKey){book=b;break;}}
            if(book==null)throw new InvalidOperationException("工作簿已关闭或另存，请刷新目录。");
            if((bool)book.ProtectStructure)throw new InvalidOperationException("工作簿结构已保护，请先在 Excel 解除保护。");
            if((bool)book.ReadOnly)throw new InvalidOperationException("工作簿为只读，请另存可编辑副本后重试。");
            dynamic sheets=book.Sheets;
            for(int i=1;i<=sheets.Count;i++){dynamic s=sheets[i];if(SheetKey(s)==request.Source.Key)sheet=s;if(request.Target!=null&&SheetKey(s)==request.Target.Key)target=s;}
            if(sheet==null)throw new InvalidOperationException("工作表已变化，请刷新目录。");
            switch(request.Action)
            {
                case "rename":
                    string error=RenamePlanner.ValidateName(request.Name);if(error!=null)throw new InvalidOperationException(error);
                    for(int i=1;i<=sheets.Count;i++){dynamic s=sheets[i];if(SheetKey(s)!=request.Source.Key&&string.Equals((string)s.Name,request.Name,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("同一工作簿已存在此名称。");}
                    sheet.Name=request.Name;break;
                case "add":book.Worksheets.Add(Type.Missing,sheet);break;
                case "hide":
                    int visible=0;for(int i=1;i<=sheets.Count;i++)if((int)sheets[i].Visible==-1)visible++;
                    if((int)sheet.Visible==-1&&visible<=1)throw new InvalidOperationException("至少保留一个可见工作表。");sheet.Visible=0;break;
                case "show":sheet.Visible=-1;break;
                case "first":if((int)sheet.Index!=1)sheet.Move(sheets[1]);break;
                case "last":if((int)sheet.Index!=sheets.Count)sheet.Move(Type.Missing,sheets[sheets.Count]);break;
                case "up":if((int)sheet.Index>1)sheet.Move(sheets[(int)sheet.Index-1]);break;
                case "down":if((int)sheet.Index<sheets.Count)sheet.Move(Type.Missing,sheets[(int)sheet.Index+1]);break;
                case "move":
                    if(request.Target==null||request.Target.WorkbookKey!=request.Source.WorkbookKey||target==null)throw new InvalidOperationException("只支持同一工作簿内调整顺序。");
                    if(request.Source.Key!=request.Target.Key){if(request.After)sheet.Move(Type.Missing,target);else sheet.Move(target);}break;
                default:throw new InvalidOperationException("未知工作表操作。");
            }
        }
    }
}
