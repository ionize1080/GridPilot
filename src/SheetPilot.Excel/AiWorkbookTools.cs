using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private object WorkbookOperation(Snapshot s, string op, JObject o)
        {
            dynamic book = s.book, app = application();
            if (op == "workbook.describe")
            {
                var sheets = new JArray(); foreach (dynamic sh in book.Sheets) { try { sheets.Add(new JObject { ["name"] = (string)sh.Name, ["visible"] = (int)sh.Visible, ["index"] = (int)sh.Index }); } finally { Release(sh); } }
                return new { name = (string)book.Name, path = (string)book.FullName, fileFormat = (int)book.FileFormat, hasVba = (bool)book.HasVBProject, readOnly = (bool)book.ReadOnly, protectedStructure = (bool)book.ProtectStructure, date1904 = (bool)book.Date1904, calculation = (int)app.Calculation, sheets };
            }
            if (op == "workbook.list")
            {
                if (o["directory"] != null) return new { files = Directory.EnumerateFiles(LocalPath((string)o["directory"])).Where(p => new[] { ".xlsx", ".xlsm", ".xlsb", ".xls" }.Contains(Path.GetExtension(p).ToLowerInvariant())).Take(500).ToArray() };
                var names = new List<string>(); foreach (dynamic b in app.Workbooks) { try { names.Add((string)b.Name); } finally { Release(b); } } return new { workbooks = names };
            }
            if (op == "workbook.create") { object b = app.Workbooks.Add(-4167); try { return new { name = (string)((dynamic)b).Name, message = "已创建，请重新获取上下文。" }; } finally { Release(b); } }
            if (op == "workbook.open" || op == "workbook.import")
            {
                string path;
                if (op == "workbook.import") { string ext = (string)o["extension"] ?? ".xlsx"; if (!new[] { ".xlsx", ".xlsm", ".xlsb", ".xls" }.Contains(ext)) throw new ArgumentException("不支持的文件格式。"); byte[] bytes = Convert.FromBase64String(Required(o, "base64")); if (bytes.Length > 32 * 1024 * 1024) throw new ArgumentException("导入最多 32 MB。"); path = Path.Combine(History.Root, "import-" + Guid.NewGuid().ToString("N") + ext); File.WriteAllBytes(path, bytes); }
                else path = LocalPath(Required(o, "path"));
                int security = app.AutomationSecurity; object opened = null;
                try { app.AutomationSecurity = 3; opened = app.Workbooks.Open(path, UpdateLinks: 0); return new { name = (string)((dynamic)opened).Name, message = "已打开（未运行宏或更新外部链接），请重新获取上下文。" }; }
                finally { app.AutomationSecurity = security; Release(opened); }
            }
            if (op == "workbook.export")
            {
                string version = SaveVersion(s.book, "导出"); string path = History.Materialize(version, Path.Combine(History.Root, "exports"));
                if (o["path"] != null) { string target = LocalPath((string)o["path"]); File.Copy(path, target, false); return new { path = target }; }
                if (new FileInfo(path).Length > 4 * 1024 * 1024) throw new InvalidOperationException("base64 导出限 4 MB，请指定本地 path。");
                return new { base64 = Convert.ToBase64String(File.ReadAllBytes(path)), extension = Path.GetExtension(path) };
            }
            if (op == "workbook.settings")
            {
                if (o["properties"] is JObject properties) foreach (var property in properties.Properties()) { dynamic item = book.BuiltinDocumentProperties[property.Name]; try { item.Value = ((JValue)property.Value).Value; } finally { Release(item); } }
                if (o["calculation"] != null) app.Calculation = (int)o["calculation"];
                if (o["date1904"] != null) book.Date1904 = (bool)o["date1904"];
                if (o["protect"] != null) { if ((bool)o["protect"]) book.Protect((string)o["password"] ?? "", true, false); else book.Unprotect((string)o["password"] ?? ""); }
                return new { updated = true };
            }
            if (op.StartsWith("vba."))
            {
                object project = null, components = null, module = null, code = null;
                try
                {
                    try { project = book.VBProject; components = ((dynamic)project).VBComponents; }
                    catch (Exception ex) { throw new InvalidOperationException("Excel 未允许访问 VBA 项目，或项目已锁定。请用户自行检查信任中心设置；工具不会修改安全设置。", ex); }
                    if (op == "vba.read")
                    {
                        var result = new JArray(); int offset = Math.Max(0, (int?)o["offset"] ?? 0), limit = Math.Min(500, Math.Max(1, (int?)o["limit"] ?? 200));
                        for (int i = 1; i <= (int)((dynamic)components).Count; i++) { dynamic c = ((dynamic)components).Item(i); object cm = null; try { if (o["module"] != null && (string)c.Name != (string)o["module"]) continue; cm = c.CodeModule; int count = ((dynamic)cm).CountOfLines, take = Math.Max(0, Math.Min(limit, count - offset)); result.Add(new JObject { ["module"] = (string)c.Name, ["type"] = (int)c.Type, ["totalLines"] = count, ["offset"] = offset, ["code"] = take == 0 ? "" : (string)((dynamic)cm).Lines[offset + 1, take], ["hasMore"] = offset + take < count }); } finally { Release(cm); Release(c); } }
                        return new { modules = result };
                    }
                    string name = Required(o, "module");
                    for (int i = 1; i <= (int)((dynamic)components).Count; i++) { object c = ((dynamic)components).Item(i); if (string.Equals((string)((dynamic)c).Name, name, StringComparison.OrdinalIgnoreCase)) { module = c; break; } Release(c); }
                    if (op == "vba.delete") { if (module == null) throw new ArgumentException("模块不存在。"); if ((int)((dynamic)module).Type == 100) throw new ArgumentException("文档模块不能删除，请写入空代码清空。"); ((dynamic)components).Remove(module); return new { deleted = name }; }
                    if ((int)book.FileFormat == 51) throw new InvalidOperationException("请先将工作簿另存为 xlsm、xlsb 或 xls 后写入 VBA，以确保快照保留宏。");
                    string source = (string)o["code"] ?? throw new ArgumentException("缺少 code。"); VbaRisk.Inspect(source);
                    if (module == null) { int type = (int?)o["type"] ?? 1; if (type != 1 && type != 2) throw new ArgumentException("新建模块仅支持标准模块 1、类模块 2。"); module = ((dynamic)components).Add(type); ((dynamic)module).Name = name; }
                    code = ((dynamic)module).CodeModule; int old = ((dynamic)code).CountOfLines; if (old > 0) ((dynamic)code).DeleteLines(1, old); if (source.Length > 0) ((dynamic)code).AddFromString(source);
                    return new { written = name, executed = false };
                }
                finally { Release(code); Release(module); Release(components); Release(project); }
            }
            object sheet = LiveSheet(s, o), range = null;
            try
            {
                if (op == "csv.import")
                {
                    var rows = new List<string[]>(); using (var parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(LocalPath(Required(o, "path")), System.Text.Encoding.UTF8)) { parser.SetDelimiters((string)o["delimiter"] ?? ","); parser.HasFieldsEnclosedInQuotes = true; while (!parser.EndOfData) { rows.Add(parser.ReadFields()); if (rows.Count > 100000) throw new ArgumentException("CSV 最多 100000 行。"); } }
                    int width = rows.Count == 0 ? 0 : rows.Max(r => r.Length); if ((long)rows.Count * width > 2000000) throw new ArgumentException("CSV 最多 200 万个单元格。");
                    var start = AiRange.Parse((string)o["address"] ?? "A1"); if (rows.Count > 0) { range = ((dynamic)sheet).Range[Address(start.Row, start.Column) + ":" + Address(start.Row + rows.Count - 1, start.Column + width - 1)]; object[,] values = new object[rows.Count, width]; for (int r = 0; r < rows.Count; r++) for (int c = 0; c < width; c++) values[r, c] = c < rows[r].Length ? rows[r][c] : ""; ((dynamic)range).NumberFormat = "@"; ((dynamic)range).Value2 = values; } return new { rows = rows.Count, columns = width };
                }
                range = ((dynamic)sheet).Range[Required(o, "address")]; string target = LocalPath(Required(o, "path"));
                using (var writer = new StreamWriter(new FileStream(target, FileMode.CreateNew), new System.Text.UTF8Encoding(true))) { object values = ((dynamic)range).Value2; int nr = ((dynamic)range).Rows.Count, nc = ((dynamic)range).Columns.Count; for (int r = 0; r < nr; r++) writer.WriteLine(string.Join(",", Enumerable.Range(0, nc).Select(c => "\"" + Convert.ToString(At(values, r, c)).Replace("\"", "\"\"") + "\""))); }
                return new { path = target };
            }
            finally { Release(range); Release(sheet); }
        }
        private object SheetOperation(Snapshot s, string op, JObject o)
        {
            dynamic book = s.book; object sheets = book.Worksheets, sheet = null;
            try
            {
                if (op == "sheet.create") { sheet = ((dynamic)sheets).Add(After: ((dynamic)sheets).Item[((dynamic)sheets).Count]); ((dynamic)sheet).Name = Required(o, "name"); return new { name = (string)((dynamic)sheet).Name }; }
                if (op == "sheet.batch_rename")
                {
                    Release(sheets);sheets=book.Sheets;
                    var changes = o["changes"] as JArray ?? throw new ArgumentException("缺少 changes。"); var targets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var change in changes) { string old = (string)change["from"], name = (string)change["to"]; if (targets.ContainsKey(old)) throw new ArgumentException("源工作表重复。"); ValidateSheetName(name); targets.Add(old, name); }
                    if (targets.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count() != targets.Count) throw new ArgumentException("目标名称重复。");
                    var refs = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    try
                    {
                        foreach (dynamic sh in (dynamic)sheets) { if (targets.ContainsKey((string)sh.Name)) refs.Add((string)sh.Name, (object)sh); else { bool collision = targets.Values.Contains((string)sh.Name, StringComparer.OrdinalIgnoreCase); Release(sh); if (collision) throw new ArgumentException("目标名称与未改名工作表冲突。"); } }
                        if (refs.Count != targets.Count) throw new ArgumentException("源工作表不存在。");
                        foreach (var pair in refs) ((dynamic)pair.Value).Name = "gp_" + Guid.NewGuid().ToString("N").Substring(0, 24);
                        foreach (var pair in refs) ((dynamic)pair.Value).Name = targets[pair.Key];
                    }
                    catch { foreach (var pair in refs) try { ((dynamic)pair.Value).Name = "gp_" + Guid.NewGuid().ToString("N").Substring(0, 24); } catch { } foreach (var pair in refs) try { ((dynamic)pair.Value).Name = pair.Key; } catch { } throw; }
                    finally { foreach (var sh in refs.Values) Release(sh); }
                    return new { renamed = targets.Count };
                }
                sheet = LiveSheet(s, o); dynamic w = sheet;
                switch (op)
                {
                    case "sheet.describe": return new { name = (string)w.Name, index = (int)w.Index, visible = (int)w.Visible, usedRange = (string)w.UsedRange.Address[false, false], protectedContents = (bool)w.ProtectContents, autoFilter = (bool)w.AutoFilterMode, merges = (object)w.UsedRange.MergeCells, conditionalFormats = (int)w.UsedRange.FormatConditions.Count, shapes = (int)w.Shapes.Count, notes = (int)w.Comments.Count, tables = (int)w.ListObjects.Count, charts = (int)w.ChartObjects().Count, pivots = (int)w.PivotTables().Count };
                    case "sheet.rename": ValidateSheetName(Required(o, "name")); w.Name = (string)o["name"]; break;
                    case "sheet.copy": w.Copy(After: w); if (o["name"] != null) ((dynamic)application()).ActiveSheet.Name = (string)o["name"]; break;
                    case "sheet.delete": dynamic app = application(); bool alerts = app.DisplayAlerts; try { app.DisplayAlerts = false; w.Delete(); } finally { app.DisplayAlerts = alerts; } break;
                    case "sheet.layout":
                        if (o["visible"] != null) w.Visible = (int)o["visible"];
                        if (o["tabColor"] != null) w.Tab.Color = AiRange.Color((string)o["tabColor"]);
                        if (o["protect"] != null) { if ((bool)o["protect"]) w.Protect((string)o["password"] ?? ""); else w.Unprotect((string)o["password"] ?? ""); }
                        if (o["freezeRows"] != null || o["freezeColumns"] != null) { w.Activate(); dynamic win = ((dynamic)application()).ActiveWindow; try { win.FreezePanes = false; win.SplitRow = (int?)o["freezeRows"] ?? 0; win.SplitColumn = (int?)o["freezeColumns"] ?? 0; win.FreezePanes = true; } finally { Release(win); } }
                        if (o["print"] is JObject print) { object setup = w.PageSetup; try { SetProperties(setup, print, "Orientation PaperSize Zoom FitToPagesWide FitToPagesTall PrintArea PrintTitleRows PrintTitleColumns LeftMargin RightMargin TopMargin BottomMargin CenterHorizontally CenterVertically LeftHeader CenterHeader RightHeader LeftFooter CenterFooter RightFooter"); } finally { Release(setup); } }
                        if (o["address"] != null) { object r = w.Range[(string)o["address"]]; try { if (o["rowHeight"] != null) ((dynamic)r).RowHeight = (double)o["rowHeight"]; if (o["columnWidth"] != null) ((dynamic)r).ColumnWidth = (double)o["columnWidth"]; if (o["hidden"] != null) { if ((string)o["axis"] == "columns") ((dynamic)r).EntireColumn.Hidden = (bool)o["hidden"]; else ((dynamic)r).EntireRow.Hidden = (bool)o["hidden"]; } if (o["group"] != null) { dynamic axis = (string)o["axis"] == "columns" ? ((dynamic)r).EntireColumn : ((dynamic)r).EntireRow; try { if ((bool)o["group"]) axis.Group(); else axis.Ungroup(); } finally { Release(axis); } } if (o["autoFilter"] != null) { if ((bool)o["autoFilter"]) ((dynamic)r).AutoFilter(); else w.AutoFilterMode = false; } } finally { Release(r); } }
                        break;
                }
                return new { updated = true };
            }
            finally { Release(sheet); Release(sheets); }
        }
        private static void ValidateSheetName(string name) { if (string.IsNullOrWhiteSpace(name) || name.Length > 31 || name.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0 || name.StartsWith("'") || name.EndsWith("'")) throw new ArgumentException("工作表名称不合法。"); }
        private static object ComGet(object target, string property) => target.GetType().InvokeMember(property, BindingFlags.GetProperty, null, target, null);
        private static void ComSet(object target, string property, object value) => target.GetType().InvokeMember(property, BindingFlags.SetProperty, null, target, new[] { value });
        private static void SetProperties(object target, JObject values, string allowed)
        {
            var names = new HashSet<string>(allowed.Split(' '), StringComparer.OrdinalIgnoreCase);
            foreach (var p in values.Properties()) { if (!names.Contains(p.Name) || !(p.Value is JValue)) throw new ArgumentException("不支持的属性：" + p.Name); ComSet(target, p.Name, ((JValue)p.Value).Value); }
        }
    }
}
