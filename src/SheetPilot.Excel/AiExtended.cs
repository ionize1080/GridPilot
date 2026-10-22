using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private WorkbookSnapshotStore historyStore;
        public int SnapshotBudgetMB { get; set; } = 500;
        public void ApplySnapshotBudget() { History.BudgetBytes=(long)Math.Max(100,SnapshotBudgetMB)*1024*1024; History.Prune(); }
        private WorkbookSnapshotStore History => historyStore ?? (historyStore = new WorkbookSnapshotStore(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GridPilot", "Recovery", "history-v2")));
        
        private static readonly HashSet<string> ReadOperations = new HashSet<string> { "capability.list", "range.spill", "comment.list", "query.list", "connection.list", "model.inspect", "slicer.list", "plugin.rename_preview", "workbook.describe", "workbook.list", "sheet.describe", "range.read", "range.trace", "vba.read", "chart.list", "chart.inspect", "pivot.list", "pivot.inspect", "snapshot.list", "table.list", "name.list", "note.list", "image.list" };
        private string SaveVersion(object book, string label)
        {
            History.BudgetBytes=(long)Math.Max(100,SnapshotBudgetMB)*1024*1024;
            dynamic b = book; int format = Convert.ToInt32(b.FileFormat);
            string ext = format == 52 ? ".xlsm" : format == 50 ? ".xlsb" : format == 56 || format == -4143 ? ".xls" : format == 51 ? ".xlsx" : null;
            if (ext == null) throw new InvalidOperationException("请先另存为 xlsx、xlsm、xlsb 或 xls。");
            string temp = Path.Combine(History.Root, Guid.NewGuid().ToString("N") + ext);
            dynamic app = application(); bool events = app.EnableEvents;
            try { app.EnableEvents = false; b.SaveCopyAs(temp); string id = History.Save(temp, (string)b.Name, label); return id; }
            finally { app.EnableEvents = events; if (File.Exists(temp)) File.Delete(temp); }
        }
        public object SnapshotHistory() => new { versions = History.List(), folder = History.Root, usedBytes=History.UsedBytes, budgetBytes=History.BudgetBytes };
        public object OpenVersion(string id)
        {
            // Restore as a complete independent workbook: preserve unsaved edits and
            // avoid copying sheets (which loses VBA, charts, names and pivot caches).
            string file = History.Materialize(id, Path.Combine(History.Root, "restored"));
            dynamic app = application(); int security = app.AutomationSecurity; bool events = app.EnableEvents;
            object books = null, book = null;
            try { app.AutomationSecurity = 3; app.EnableEvents = false; books = app.Workbooks; book = ((dynamic)books).Open(file, UpdateLinks: 0, ReadOnly: false); ((dynamic)book).Activate(); return new { message = "已打开完整快照副本，原工作簿及未保存编辑仍保留。请重新发送问题以绑定恢复副本。", book = (string)((dynamic)book).Name, versionId = id }; }
            finally { app.EnableEvents = events; app.AutomationSecurity = security; Release(book); Release(books); }
        }
        public object Extended(JObject request)
        {
            var snapshot = Get((string)request["snapshotId"]);
            string operation = (string)request["operation"];
            var options = request["options"] as JObject ?? new JObject();
            CheckDiscoveryBoundary(snapshot,operation,options);
            CheckModernCapability(snapshot,operation,options);
            if (ReadOperations.Contains(operation)) return ExecuteExtended(snapshot, operation, options);
            if (!ExtendedOperations.Contains(operation)) throw new InvalidOperationException("未知操作，请先查询 excel_help。");
            if ((bool)((dynamic)snapshot.book).ReadOnly) throw new InvalidOperationException("工作簿只读。");
            string[] risks = operation == "vba.write" ? VbaRisk.Inspect((string)options["code"]) : new string[0];
            var p = new Proposal { id = Guid.NewGuid().ToString("N"), sequence = ++proposalSequence, snapshot = snapshot.id, kind = "extended", sheet = (string)options["sheet"] ?? "", address = (string)options["address"] ?? "A1", options = new JObject { ["operation"] = operation, ["options"] = options.DeepClone(), ["risks"] = JArray.FromObject(risks) } };
            // Bind approval to the exact immutable request; never accept approval from tool arguments.
            bool applied = risks.Length == 0 && AutomaticFor(snapshot.book);
            if (!applied) p.options["baseline"] = History.Signature(SaveVersion(snapshot.book, "待审批 · " + operation));
            proposals.Add(p.id, p);
            object result = null;
            if (applied) result = ApplyExtended(new[] { p }, false);
            return new { proposalId = p.id, snapshotId = p.snapshot, book = snapshot.bookName, sheet = p.sheet, address = p.address, kind = operation, data = options, requiresManualApproval = risks.Length > 0, risks, applied, result, checkpoint = CheckpointState(), status = applied ? "已应用并保存快照。" : risks.Length > 0 ? "检测到网络或执行能力，必须手动批准此 VBA；自动批准不适用。" : "待手动应用。" };
        }
        private void CheckExtended(Proposal[] batch, bool manual)
        {
            foreach (var p in batch) { Get(p.snapshot); if (!manual && ((JArray)p.options["risks"]).Count > 0) throw new InvalidOperationException("危险 VBA 必须由用户手动批准。"); }
            foreach(var group in batch.Where(p=>p.options["baseline"]!=null).GroupBy(p=>Get(p.snapshot).book))
            {
                string signature=History.Signature(SaveVersion(group.Key,"审批前检查"));
                if(group.Any(p=>(string)p.options["baseline"]!=signature))throw new InvalidOperationException("待审批期间工作簿已变化，整批未执行。请丢弃旧建议并重新生成。");
            }
        }
        private object ApplyExtended(Proposal[] batch, bool manual, bool checkedBeforeBatch=false)
        {
            if(!checkedBeforeBatch)CheckExtended(batch,manual);
            var ids = new List<string>(); var results = new List<object>();
            foreach (var p in batch)
            {
                var s = Get(p.snapshot); string operation = (string)p.options["operation"];
                string before = SaveVersion(s.book, "变更前 · " + operation);
                try
                {
                    dynamic app = application(); bool events = app.EnableEvents, screen = app.ScreenUpdating;
                    try { app.EnableEvents = false; app.ScreenUpdating = false; results.Add(ExecuteExtended(s, operation, (JObject)p.options["options"])); }
                    finally { app.EnableEvents = events; app.ScreenUpdating = screen; }
                    SaveVersion(s.book, "变更后 · " + operation); undo=null;checkpointFolder=History.Root;newCheckpoint=true; proposals.Remove(p.id); ids.Add(p.id);
                }
                catch (Exception ex) { throw new InvalidOperationException("操作未完成，可能已有部分变更。完整变更前快照 " + before + " 已保留，可从历史快照打开；已完成项：" + string.Join(",", ids) + "。原因：" + ex.Message, ex); }
            }
            return new { message = "已应用并保留完整历史快照。", appliedIds = ids, retiredIds = ids, results, checkpoint = CheckpointState() };
        }
        private static readonly HashSet<string> ExtendedOperations = new HashSet<string>((
            "range.formula2 range.calculate comment.add comment.reply comment.delete query.set query.delete slicer.filter slicer.clear slicer.timeline table.resize workbook.create workbook.open workbook.export workbook.import workbook.settings sheet.create sheet.rename sheet.copy sheet.delete sheet.layout sheet.batch_rename " +
            "range.write range.clear range.copy range.insert range.delete range.sort range.replace range.transform range.format range.merge range.conditional_format range.validation range.goal_seek " +
            "csv.import csv.export vba.write vba.delete table.create table.update table.delete name.set name.delete note.set note.delete image.insert image.delete sparkline.add sparkline.delete " +
            "plugin.rename plugin.sort_sheets plugin.convert plugin.copy plugin.segment plugin.simple_mask chart.create chart.edit chart.delete chart.series chart.trendline chart.errorbars pivot.create pivot.edit pivot.field pivot.filter pivot.sort pivot.group pivot.calculated pivot.refresh pivot.delete slicer.add slicer.delete snapshot.create").Split(' '));
        private object ExecuteExtended(Snapshot s, string operation, JObject o)
        {
            CheckDiscoveryBoundary(s,operation,o);
            CheckModernCapability(s,operation,o);
            if(IsModernOperation(operation))return ModernOperation(s,operation,o);
            if (operation.StartsWith("plugin.")) return PluginOperation(s, operation, o);
            if (operation == "snapshot.list") { var versions=History.List().Where(v=>(string)v["book"]==s.bookName).ToArray();int offset=Math.Max(0,(int?)o["offset"]??0),limit=Math.Min(200,Math.Max(1,(int?)o["limit"]??50));return new{versions=versions.Skip(offset).Take(limit).ToArray(),total=versions.Length,nextOffset=offset+limit<versions.Length?(int?)(offset+limit):null}; }
            if (operation == "snapshot.create") return new { versionId = SaveVersion(s.book, (string)o["label"] ?? "手动快照") };
            if (operation.StartsWith("workbook.") || operation.StartsWith("vba.") || operation.StartsWith("csv.")) return WorkbookOperation(s, operation, o);
            if (operation.StartsWith("sheet.")) return SheetOperation(s, operation, o);
            if (operation.StartsWith("range.")) return RangeOperation(s, operation, o);
            if (operation.StartsWith("chart.")) return ChartOperation(s, operation, o);
            if (operation.StartsWith("pivot.") || operation.StartsWith("slicer.")) return PivotOperation(s, operation, o);
            return ObjectOperation(s, operation, o);
        }
        private object LiveSheet(Snapshot s, JObject o)
        {
            string name = (string)o["sheet"]; if (string.IsNullOrEmpty(name)) throw new ArgumentException("缺少 sheet。");
            object sheets = ((dynamic)s.book).Worksheets;
            try { return ((dynamic)sheets).Item[name]; } finally { Release(sheets); }
        }
        private static string Required(JObject o, string key) { string value = (string)o[key]; if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("缺少 " + key); return value; }
        private static string LocalPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.StartsWith(@"\\") || path.Contains("://")) throw new ArgumentException("只允许本地文件路径。");
            string full = Path.GetFullPath(path); if (new DriveInfo(Path.GetPathRoot(full)).DriveType == DriveType.Network) throw new ArgumentException("不支持网络驱动器。"); return full;
        }
    }
}
