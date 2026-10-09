using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed class ExcelGateway : IWorkbookGateway
    {
        private readonly Func<object> application;
        private readonly int threadId = Thread.CurrentThread.ManagedThreadId;
        public ExcelGateway(Func<object> application) { this.application = application; }
        private void CheckThread()
        { if (Thread.CurrentThread.ManagedThreadId != threadId) throw new InvalidOperationException("Excel COM access must stay on its owning thread."); }

        public WorkbookSnapshot Capture(bool allWorkbooks = false)
        {
            CheckThread();
            dynamic app = application();
            dynamic active = app.ActiveWorkbook;
            if (active == null) return WorkbookSnapshot.Empty;
            string activeKey = WorkbookKey(active);
            string activeSheetName = app.ActiveSheet == null ? "" : (string)app.ActiveSheet.Name;
            var result = new List<SheetItem>();
            int bookCount = 0;
            if (!allWorkbooks) { AddSheets(active, activeKey, activeSheetName, result); bookCount = 1; }
            else
            {
                dynamic books = app.Workbooks;
                int count = books.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic book = books.Item[i];
                    if ((bool)book.IsAddin) continue;
                    AddSheets(book, activeKey, activeSheetName, result); bookCount++;
                }
            }
            return new WorkbookSnapshot(activeKey, allWorkbooks ? "全部 " + bookCount + " 个工作簿 · 当前 " + (string)active.Name : (string)active.Name, result);
        }

        private static void AddSheets(dynamic book, string activeKey, string activeSheetName, List<SheetItem> result)
        {
            string bookKey = WorkbookKey(book), bookName = book.Name;
            dynamic sheets = book.Sheets;
            int count = sheets.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic sheet = sheets.Item[i];
                string name = sheet.Name;
                int visibility = (int)sheet.Visible;
                // Chart.Type is a chart style, not XlSheetType. Inspect COM type information first.
                bool isChart = Microsoft.VisualBasic.Information.TypeName((object)sheet).TrimStart('_') == "Chart";
                int type = isChart ? -4109 : (int)sheet.Type;
                result.Add(new SheetItem(SheetKey(sheet), name, i,
                    visibility == -1 ? SheetVisibility.Visible : visibility == 2 ? SheetVisibility.VeryHidden : SheetVisibility.Hidden,
                    bookKey == activeKey && name == activeSheetName, type == -4167 ? "工作表" : type == -4109 ? "图表页" : "特殊页面", bookKey, bookName));
            }
        }

        public void Navigate(string workbookKey, SheetItem target)
        {
            CheckThread();
            dynamic app = application();
            dynamic books = app.Workbooks;
            int count = books.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic book = books.Item[i];
                if (WorkbookKey(book) != workbookKey) continue;
                dynamic sheets = book.Sheets;
                int sheetCount = sheets.Count;
                for (int j = 1; j <= sheetCount; j++)
                {
                    dynamic sheet = sheets.Item[j];
                    if (SheetKey(sheet) != target.Key) continue;
                    // Recheck at the moment of execution, not just from the cached UI row.
                    if ((int)sheet.Visible != -1) throw new InvalidOperationException("此页已隐藏，请先在 Excel 中取消隐藏。");
                    book.Activate();
                    sheet.Activate();
                    return;
                }
                throw new InvalidOperationException("此页已删除或标识已改变，请刷新后重试。");
            }
            throw new InvalidOperationException("工作簿已关闭或另存为，请刷新后重试。");
        }

        private static string SheetKey(dynamic sheet)
        {
            try { string codeName = sheet.CodeName; if (!string.IsNullOrEmpty(codeName)) return "code:" + codeName; }
            catch (COMException) { }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
            // A fresh unsaved workbook can return an empty CodeName. COM identity survives a tab rename.
            IntPtr identity = Marshal.GetIUnknownForObject((object)sheet);
            try { return "com:" + identity.ToInt64().ToString("X"); }
            finally { Marshal.Release(identity); }
        }
        private static string WorkbookKey(dynamic workbook)
        {
            // Balance ONLY this explicit IUnknown AddRef; never FinalRelease a shared Office RCW.
            IntPtr pointer = Marshal.GetIUnknownForObject((object)workbook);
            try { return pointer.ToInt64().ToString("X") + ":" + (string)workbook.Name; }
            finally { Marshal.Release(pointer); }
        }
    }
}
