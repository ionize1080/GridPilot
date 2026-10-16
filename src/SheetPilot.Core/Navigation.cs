using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace SheetPilot.Core
{
    public enum SheetVisibility { Visible, Hidden, VeryHidden }

    // Detached values only. Never put a COM object in a snapshot or a worker task.
    public sealed class SheetItem
    {
        public SheetItem(string key, string name, int index, SheetVisibility visibility, bool active, string kind, string workbookKey = "", string workbookName = "")
        { Key = key; Name = name; Index = index; Visibility = visibility; IsActive = active; Kind = kind; WorkbookKey = workbookKey; WorkbookName = workbookName; }
        public string Key { get; }
        public string Name { get; }
        public int Index { get; }
        public SheetVisibility Visibility { get; }
        public bool IsActive { get; }
        public string Kind { get; }
        public string WorkbookKey { get; }
        public string WorkbookName { get; }
        public string Subtitle => string.IsNullOrEmpty(WorkbookName) ? Detail : WorkbookName + " · " + Detail;
        public bool CanNavigate => Visibility == SheetVisibility.Visible;
        public string Number => Index.ToString("00");
        public string Detail => Visibility == SheetVisibility.VeryHidden ? "深度隐藏" : Visibility == SheetVisibility.Hidden ? "隐藏" : Kind;
        public string Hint => CanNavigate ? "单击跳转 · Enter 打开" : "此页已隐藏；请先在 Excel 中取消隐藏";
    }

    public sealed class WorkbookSnapshot
    {
        public static readonly WorkbookSnapshot Empty = new WorkbookSnapshot("", "未打开工作簿", new SheetItem[0]);
        public WorkbookSnapshot(string key, string name, IEnumerable<SheetItem> sheets)
        { Key = key; Name = name; Sheets = Array.AsReadOnly(sheets.ToArray()); }
        public string Key { get; }
        public string Name { get; }
        public IReadOnlyList<SheetItem> Sheets { get; }
        public bool SameContent(WorkbookSnapshot other) => other != null && Key == other.Key && Name == other.Name &&
            Sheets.Count == other.Sheets.Count && Sheets.Zip(other.Sheets, (a,b) => a.Key == b.Key && a.Name == b.Name &&
                a.Index == b.Index && a.Visibility == b.Visibility && a.IsActive == b.IsActive && a.Kind == b.Kind && a.WorkbookKey == b.WorkbookKey && a.WorkbookName == b.WorkbookName).All(x => x);
    }

    public sealed class SearchOptions
    {
        public string Query { get; set; } = "";
        public bool UseRegex { get; set; }
        public bool MatchCase { get; set; }
        public bool IncludeHidden { get; set; } = true;
    }

    public sealed class SearchResult
    {
        public SearchResult(IEnumerable<SheetItem> items, string error = null)
        { Items = Array.AsReadOnly(items.ToArray()); Error = error; }
        public IReadOnlyList<SheetItem> Items { get; }
        public string Error { get; }
    }

    public sealed class SheetSearch
    {
        // A per-match timeout AND an overall budget protect against pathological expressions / large books.
        public SearchResult Filter(IReadOnlyList<SheetItem> sheets, SearchOptions options, CancellationToken cancellation)
        {
            var found = new List<SheetItem>();
            var query = options.Query ?? "";
            if (query.Length > 512) return Error("搜索表达式最多 512 个字符。");
            Regex regex = null;
            try
            {
                if (options.UseRegex && query.Length != 0)
                    regex = new Regex(query, RegexOptions.CultureInvariant | (options.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase), TimeSpan.FromMilliseconds(40));
            }
            catch (ArgumentException) { return Error("正则表达式无效，请检查括号、转义或字符组。"); }
            var clock = Stopwatch.StartNew();
            try
            {
                foreach (var sheet in sheets)
                {
                    cancellation.ThrowIfCancellationRequested();
                    if (clock.ElapsedMilliseconds > 200) return Error("搜索耗时过长，请简化表达式。");
                    if (!options.IncludeHidden && !sheet.CanNavigate) continue;
                    bool match = query.Length == 0 || (regex != null ? regex.IsMatch(sheet.Name) :
                        sheet.Name.IndexOf(query, options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) >= 0);
                    if (match) found.Add(sheet);
                }
            }
            catch (RegexMatchTimeoutException) { return Error("正则匹配超时，请简化表达式。"); }
            return new SearchResult(found);
        }
        private static SearchResult Error(string message) => new SearchResult(new SheetItem[0], message);
    }

    public interface IWorkbookGateway
    {
        WorkbookSnapshot Capture(bool allWorkbooks = false);
        void Navigate(string workbookKey, SheetItem sheet);
    }

    public sealed class FeatureCommand
    {
        public FeatureCommand(string id, string label, string description, string imageMso, Action execute)
        { Id = id; Label = label; Description = description; ImageMso = imageMso; Execute = execute; }
        public string Id { get; }
        public string Label { get; }
        public string Description { get; }
        public string ImageMso { get; }
        public Action Execute { get; }
    }

    public interface IFeatureModule : IDisposable
    {
        string Id { get; }
        string Label { get; }
        IEnumerable<FeatureCommand> Commands { get; }
    }

    // Modules are explicitly registered, not arbitrary DLLs loaded from writable directories.
    public sealed class FeatureRegistry : IDisposable
    {
        private readonly List<IFeatureModule> modules = new List<IFeatureModule>();
        private readonly Dictionary<string, FeatureCommand> commands = new Dictionary<string, FeatureCommand>();
        public IReadOnlyList<IFeatureModule> Modules => modules.AsReadOnly();
        public void Add(IFeatureModule module)
        {
            if (modules.Any(m => m.Id == module.Id)) throw new ArgumentException("Duplicate module id.");
            var additions = module.Commands.ToArray();
            if (additions.Select(c => c.Id).Distinct().Count() != additions.Length || additions.Any(c => commands.ContainsKey(c.Id)))
                throw new ArgumentException("Duplicate command id.");
            foreach (var command in additions) commands.Add(command.Id, command);
            modules.Add(module);
        }
        public void Execute(string id)
        {
            if (!commands.TryGetValue(id, out var command)) throw new ArgumentException("Unknown command.");
            command.Execute();
        }
        public void Dispose()
        {
            var errors=new List<Exception>();
            try { for (int i = modules.Count - 1; i >= 0; i--) { try { modules[i].Dispose(); } catch(Exception ex) { errors.Add(ex); } } }
            finally { modules.Clear(); commands.Clear(); }
            if(errors.Count>0)throw new AggregateException("Feature cleanup failed.",errors);
        }
    }
}
