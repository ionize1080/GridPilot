using System;
using System.IO;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        // Called only by the workspace UI after validating a managed result ID.
        // This is deliberately absent from the MCP operation catalog.
        public object OpenWorkspacePaper(string file)
        {
            if (string.IsNullOrWhiteSpace(file) || !Path.IsPathRooted(file) ||
                !string.Equals(Path.GetExtension(file), ".xlsx", StringComparison.OrdinalIgnoreCase) || !File.Exists(file))
                throw new InvalidOperationException("分析底稿不存在，请刷新工作区。");
            dynamic app = application();
            object books = null, opened = null;
            int security = (int)app.AutomationSecurity;
            bool events = (bool)app.EnableEvents;
            try
            {
                app.AutomationSecurity = 3;
                app.EnableEvents = false;
                books = app.Workbooks;
                opened = ((dynamic)books).Open(Filename: Path.GetFullPath(file), UpdateLinks: 0,
                    ReadOnly: false, IgnoreReadOnlyRecommended: true, AddToMru: false);
                ((dynamic)opened).Activate();
                return new { opened = true, book = (string)((dynamic)opened).Name, message = "底稿已在当前 Excel 中打开" };
            }
            finally
            {
                app.AutomationSecurity = security;
                app.EnableEvents = events;
                Release(opened); Release(books);
            }
        }
    }
}
