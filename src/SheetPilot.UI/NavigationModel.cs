using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public sealed class NavigationModel : INotifyPropertyChanged, IDisposable
    {
        private readonly Dispatcher dispatcher;
        private readonly DispatcherTimer debounce;
        private readonly SheetSearch search = new SheetSearch();
        private CancellationTokenSource pending;
        private WorkbookSnapshot snapshot = WorkbookSnapshot.Empty;
        private long revision;
        private bool disposed;
        private bool searchPending;
        private string query = "", message = "", status = "准备就绪";
        private bool settingsOpen;
        private bool settingsSaveFailed;
        private string settingsStatus = "设置更改后自动保存到 XLL 同目录。";
        public NavigationSettings Settings { get; }
        private IReadOnlyList<SheetItem> results = new SheetItem[0];
        private SheetItem selected;
        public NavigationModel(NavigationSettings settings = null)
        {
            Settings = settings ?? new NavigationSettings();
            Settings.PropertyChanged += SettingsChanged;
            dispatcher = Dispatcher.CurrentDispatcher;
            debounce = new DispatcherTimer(TimeSpan.FromMilliseconds(180), DispatcherPriority.Background, (s,e) => { debounce.Stop(); StartSearch(); }, dispatcher);
            debounce.Stop();
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public event Action<bool> ScopeChanged;
        public event Action RefreshRequested;
        public event Action<SheetItem> NavigateRequested;
        public string Query { get => query; set { if (Set(ref query, value)) ScheduleSearch(); } }
        public bool UseRegex { get => Settings.UseRegex; set => Settings.UseRegex = value; }
        public bool MatchCase { get => Settings.MatchCase; set => Settings.MatchCase = value; }
        public bool IncludeHidden { get => Settings.IncludeHidden; set => Settings.IncludeHidden = value; }
        public int ScopeIndex { get => Settings.ScopeIndex; set => Settings.ScopeIndex = value; }
        public bool IsSettingsOpen { get => settingsOpen; set { Set(ref settingsOpen, value); Notify(nameof(IsNavigationOpen)); } }
        public bool IsNavigationOpen => !IsSettingsOpen;
        public string SettingsStatus { get => settingsStatus; set => Set(ref settingsStatus, value); }
        public bool SettingsSaveFailed { get => settingsSaveFailed; set => Set(ref settingsSaveFailed, value); }
        public bool ShowWorkbookNames => Settings.WorkbookNameMode == 1 || (Settings.WorkbookNameMode == 0 && ScopeIndex == 1);
        public double EffectiveRowHeight => Math.Max(Settings.RowHeight, ShowWorkbookNames ? Settings.FontSize * 2 + 14 : Settings.FontSize + 12);
        public System.Windows.Thickness RowMargin => new System.Windows.Thickness(0, 0, 0, Settings.RowGap);
        private void SettingsChanged(object sender, PropertyChangedEventArgs e)
        {
            if (disposed) return;
            if (e.PropertyName == nameof(Settings.WorkbookNameMode) || e.PropertyName == nameof(Settings.ScopeIndex))
            { Notify(nameof(ShowWorkbookNames)); Notify(nameof(EffectiveRowHeight)); }
            else if (e.PropertyName == nameof(Settings.RowHeight) || e.PropertyName == nameof(Settings.FontSize)) Notify(nameof(EffectiveRowHeight));
            else if (e.PropertyName == nameof(Settings.RowGap)) Notify(nameof(RowMargin));
            if (e.PropertyName == nameof(Settings.ScopeIndex) || e.PropertyName == nameof(Settings.UseRegex) || e.PropertyName == nameof(Settings.MatchCase) || e.PropertyName == nameof(Settings.IncludeHidden)) Notify(e.PropertyName);
            if (e.PropertyName == nameof(Settings.ScopeIndex)) { ClearForScope(); ScopeChanged?.Invoke(ScopeIndex == 1); }
            else if (e.PropertyName == nameof(Settings.UseRegex) || e.PropertyName == nameof(Settings.MatchCase) || e.PropertyName == nameof(Settings.IncludeHidden)) ScheduleSearch();
        }
        public string WorkbookName => snapshot.Name;
        public IReadOnlyList<SheetItem> Results { get => results; private set => Set(ref results, value); }
        public SheetItem Selected { get => selected; set => Set(ref selected, value); }
        public string Message { get => message; private set { Set(ref message, value); Notify(nameof(HasMessage)); } }
        public bool HasMessage => !string.IsNullOrEmpty(Message);
        public string Status { get => status; private set => Set(ref status, value); }
        public bool IsEmpty => Results.Count == 0;
        public string EmptyTitle => HasMessage ? "调整搜索后重试" : snapshot.Sheets.Count == 0 ? "没有可导航的工作表" : "没有找到匹配项";
        public string EmptyHint => snapshot.Sheets.Count == 0 ? "打开工作簿后，页面将显示在这里。" : "试试更短的关键词，或切换搜索范围。";
        public void Update(WorkbookSnapshot value)
        {
            if (disposed || snapshot.SameContent(value)) return;
            bool activityOnly = !searchPending && snapshot.Sheets.Count == value.Sheets.Count &&
                snapshot.Sheets.Zip(value.Sheets, (a,b) => a.Key == b.Key && a.WorkbookKey == b.WorkbookKey && a.Name == b.Name &&
                    a.WorkbookName == b.WorkbookName && a.Index == b.Index && a.Visibility == b.Visibility && a.Kind == b.Kind).All(equal => equal);
            snapshot = value;
            Notify(nameof(WorkbookName));
            if (activityOnly)
            {
                var lookup = value.Sheets.ToDictionary(s => RenamePlanner.Key(s));
                Results = Results.Select(s => lookup[RenamePlanner.Key(s)]).ToArray();
                Selected = Results.FirstOrDefault(s => s.IsActive);
                return;
            }
            ScheduleSearch();
        }
        public void Report(string value) { Message = value; Notify(nameof(EmptyTitle)); }
        public void Refresh() => RefreshRequested?.Invoke();
        public void Navigate(SheetItem item)
        {
            if (item == null || disposed || searchPending) return;
            if (!item.CanNavigate) { Report("此页已隐藏。请先在 Excel 中取消隐藏，再跳转。"); return; }
            Message = "";
            NavigateRequested?.Invoke(item);
        }
        private void ClearForScope()
        {
            snapshot = WorkbookSnapshot.Empty;
            Results = new SheetItem[0];
            Selected = null;
            Notify(nameof(WorkbookName));
            ScheduleSearch();
        }
        private void ScheduleSearch()
        {
            if (disposed) return;
            revision++;
            searchPending = true;
            Status = "正在搜索…";
            pending?.Cancel();
            debounce.Stop(); debounce.Start();
        }
        private async void StartSearch()
        {
            if (disposed) return;
            var version = revision;
            var source = new CancellationTokenSource();
            pending = source;
            var input = snapshot;
            var options = new SearchOptions { Query = Query, UseRegex = UseRegex, MatchCase = MatchCase, IncludeHidden = IncludeHidden };
            Status = "正在搜索…";
            try
            {
                var output = await Task.Run(() => search.Filter(input.Sheets, options, source.Token));
                if (disposed || source.IsCancellationRequested || version != revision) return;
                searchPending = false;
                var old = Selected;
                Results = output.Items;
                Selected = Results.FirstOrDefault(x => x.IsActive) ?? Results.FirstOrDefault(x => old != null && x.Key == old.Key && x.WorkbookKey == old.WorkbookKey) ?? Results.FirstOrDefault();
                Message = output.Error ?? "";
                Status = Results.Count + " / " + input.Sheets.Count + " 个页面";
                Notify(nameof(IsEmpty)); Notify(nameof(EmptyTitle)); Notify(nameof(EmptyHint));
            }
            catch (OperationCanceledException) { }
            catch (Exception) { if (!disposed && version == revision) { searchPending = false; Report("搜索暂时不可用，请刷新重试。"); } }
            finally { if (ReferenceEquals(pending, source)) pending = null; source.Dispose(); }
        }
        public void Dispose() { disposed = true; Settings.PropertyChanged -= SettingsChanged; debounce.Stop(); revision++; pending?.Cancel(); }
        private bool Set<T>(ref T field, T value, [CallerMemberName] string name = null)
        { if (EqualityComparer<T>.Default.Equals(field, value)) return false; field = value; Notify(name); return true; }
        private void Notify(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
