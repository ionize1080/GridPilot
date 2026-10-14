using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Timer = System.Windows.Forms.Timer;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Forms;
using ExcelDna.Integration;
using SheetPilot.Core;
using SheetPilot.UI;

namespace SheetPilot.Excel
{
    public sealed class ToolsModule : IFeatureModule
    {
        private readonly ExcelGateway gateway;
        private readonly NavigationModule navigation;
        private readonly ToolPreferencesStore store;
        private readonly ToolPreferences settings;
        private FocusOverlay focus;
        private readonly HashSet<PortableWindow> portableWindows=new HashSet<PortableWindow>();
        public int IconTheme => settings.IconTheme;
        private readonly HashSet<Window> windows = new HashSet<Window>();
        private readonly Timer statusTimer = new Timer { Interval = 9000 };
        private object previousStatus;
        private string status;
        private bool disposed;
        public ToolsModule(ExcelGateway gateway, NavigationModule navigation)
        {
            this.gateway = gateway; this.navigation = navigation;
            store = new ToolPreferencesStore(Path.Combine(Path.GetDirectoryName(ExcelDnaUtil.XllPath), "GridPilot.tools.json")); settings = store.Load();

            statusTimer.Tick += (s,e) => ClearStatus();
        }
        public string Id => "GridPilot_Tools";
        public string Label => "常用工具";
        public bool FocusEnabled => focus?.Enabled == true;
        public bool CanUndo => gateway.CanUndo;
        internal string PerformanceSummary => focus?.PerformanceSummary ?? "Focus not started";
        public IEnumerable<FeatureCommand> Commands => new[] {
            new FeatureCommand("GridPilot_Database", "报表数据库", "文件夹报表归档、快照与累计去重、可视化和 SQL 查询；在源 Excel 修改后扫描更新。", "", () => Run(OpenDatabase)),
            new FeatureCommand("GridPilot_Mask", "数据脱敏", "选择字段、统一跨表映射；验证支持的公式关系后另存独立副本。", "", () => Run(OpenMask)),
            new FeatureCommand("GridPilot_Segment", "中英文分词", "离线分词、专业词表和三种输出；预览后生成新工作簿。", "", () => Run(()=>Show(new SegmentationWindow(navigation.Settings,gateway.CaptureTextSelection(),Path.Combine(Path.GetDirectoryName(ExcelDnaUtil.XllPath),"Resources"),gateway.WriteTextOutput)))),
            new FeatureCommand("GridPilot_CopyInternal", "原样复制·内部公式", "复制到新工作簿，保留排版和范围内公式；外部及无法确认依赖的公式按设置转为当前值或停止。Esc 可取消。", "", () => Run(() => CopySelection(true))),
            new FeatureCommand("GridPilot_CopyValues", "原样复制·仅数值", "复制到新工作簿，所有公式转为当前值，保留文本、错误值与排版。复制选项位于统一设置。Esc 可取消。", "", () => Run(() => CopySelection(false))),
            new FeatureCommand("GridPilot_Rename", "批量重命名", "选择范围、查找替换与编号；检查预览后应用。", "", () => Run(OpenRename)),
            new FeatureCommand("GridPilot_Focus", "聚焦", "临时高亮活动单元格所在行列；再次点击关闭，不修改单元格格式。", "", () => { if(focus==null)focus=new FocusOverlay(()=>ExcelDnaUtil.Application,action=>ExcelAsyncUtil.QueueAsMacro(()=>action()),settings); focus.Toggle(); Announce(focus.Enabled ? "聚焦已开启，在普通工作表视图中跟随活动单元格。" : "聚焦已关闭。"); Ribbon.Refresh(); }),
            new FeatureCommand("GridPilot_Text", "转为文本", "将选中的数值常量转为文本，保留视觉位置与其他格式。", "", () => Run(() => Announce(gateway.ConvertSelection(true, settings)))),
            new FeatureCommand("GridPilot_Number", "转为数值", "将选中的数字文本转为数值，默认保护前导零和长编号。", "", () => Run(() => Announce(gateway.ConvertSelection(false, settings)))),
            new FeatureCommand("GridPilot_Undo", "撤销上次操作", "恢复最近一次成功的 GridPilot 重命名或内容转换；不覆盖已被再次修改的目标。", "", () => Run(() => { Announce(gateway.UndoLast()); navigation.Refresh(); })),
            new FeatureCommand("GridPilot_Settings", "设置", "原样复制、聚焦、转换与重命名默认规则；设置随加载项文件夹保存。", "", () => Run(OpenSettings)) };
        private void Run(Action action) { if (disposed) return; ExcelAsyncUtil.QueueAsMacro(() => { if (disposed) return; AddIn.Guard(action); Ribbon.Refresh(); }); }
        private void CopySelection(bool formulas)
        {
            try
            {
                var result = gateway.CopySelection(formulas, settings.Copy, (stage, done, total) =>
                { if ((GetAsyncKeyState(27) & 0x8000) != 0) throw new OperationCanceledException(); });
                Announce(result.Summary); Log.Write("Copy: cells=" + result.Cells + "; kept=" + result.KeptFormulas + "; values=" + result.ValueFormulas + "; ms=" + result.ElapsedMilliseconds);
                navigation.Refresh();
            }
            catch (OperationCanceledException) { Announce("原样复制已取消，未保留未完成的结果工作簿。"); }
        }
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        public void Execute(string id) { foreach (var command in Commands) if (command.Id == id) { command.Execute(); return; } }
        private void OpenRename()
        {
            var working = new ToolPreferences { Rename = settings.Rename.Clone() };
            Show(new RenameWindow(working, navigation.Settings, all => gateway.Capture(all), plan => { var result = gateway.ApplyRename(plan); navigation.Refresh(); Ribbon.Refresh(); return result; }, () => { settings.Rename = working.Rename.Clone(); store.Save(settings); return store.Status; }));
        }
        private void OpenMask()
        {
            MaskSnapshot snapshot=null;MaskResult preview=null;string fingerprint=null;
            var lifetime=new CancellationTokenSource();CancellationTokenSource operation=null;bool busy=false;
            PortableWindow window=null;
            Action<string,int,long,long> progress=(phase,percent,current,total)=>window?.SendProgress(phase,percent,current,total);
            window=OpenPortable(async message=>
            {
                lifetime.Token.ThrowIfCancellationRequested();
                if(window.IsDisconnected)throw new OperationCanceledException();
                string action=(string)message["action"];
                if(action!="init" && action!="preview" && action!="previewPage" && action!="copyMask" && action!="export")throw new InvalidOperationException("未知操作。");
                if(busy)throw new InvalidOperationException("请等待当前操作完成。");
                busy=true;operation?.Dispose();operation=CancellationTokenSource.CreateLinkedTokenSource(lifetime.Token);var token=operation.Token;
                try
                {
                    if(action=="init")
                    {
                        snapshot=gateway.CaptureMaskWorkbook(token,progress);
                        progress("识别字段与统计词频",0,0,snapshot.Cells.Count);
                        var suggested=await Task.Run(()=>MaskInference.Suggest(snapshot,token),token);
                        return new {mode="mask",sheets=snapshot.Sheets.Select((s,i)=>new {id=i,name=s.Name,nonEmptyRows=s.NonEmptyRows,nonEmptyColumns=s.NonEmptyColumns,rows=s.Rows,columns=s.Columns}),cells=snapshot.Cells.Count,rules=suggested,settings};
                    }
                    if(snapshot==null)throw new InvalidOperationException("请先完成工作簿扫描。");
                    var selected=message["rules"].ToObject<MaskRule[]>();bool mentions=(bool?)message["replaceMentions"]??settings.MaskReplaceMentions;string key=JsonConvert.SerializeObject(selected)+mentions;
                    if((action=="export" || action=="copyMask" || action=="previewPage") && (preview==null || key!=fingerprint))throw new InvalidOperationException("字段规则已改变，请重新预览后导出。");
                    if(preview==null || key!=fingerprint)
                    {preview=null;fingerprint=null;preview=await Task.Run(()=>MaskPlanner.Build(snapshot,selected,token,mentions,progress),token);fingerprint=key;}
                    token.ThrowIfCancellationRequested();
                    if(action=="preview" || action=="previewPage")
                        return await Task.Run(()=>MaskPreview.Page(snapshot,preview,(string)message["search"],(int?)message["offset"]??0,(int?)message["pageSize"]??100,token),token);
                    if(action=="copyMask")
                    {
                        if((bool?)message["confirmed"]!=true)throw new InvalidOperationException("请先确认预览和保留字段。");
                        gateway.VerifyMaskSource(snapshot,token,progress);
                        return await Task.Run(()=>MaskPreview.Clipboard(preview,(int?)message["sheet"]??0,(int?)message["firstRow"]??1,(int?)message["lastRow"]??1),token);
                    }
                    if((bool?)message["confirmed"]!=true)throw new InvalidOperationException("请先确认预览和保留字段。");
                    string path=(string)message["path"];if(string.IsNullOrWhiteSpace(path))throw new InvalidOperationException("请选择新的保存位置。");
                    string saved=await gateway.ExportMaskAsync(snapshot,preview,path,token,(bool?)message["preserveFormatting"]??settings.MaskPreserveFormatting,progress);
                    return new {saved,message="已验证并保存："+saved};
                }
                finally{busy=false;}
            },()=>{try{operation?.Cancel();}catch(ObjectDisposedException){}});
            window.Disconnected+=()=>{lifetime.Cancel();};
        }
        private void OpenDatabase()
        {
            var lifetime=new CancellationTokenSource();
            var window=OpenPortable(async message=>
            {
                string action=(string)message["action"];
                if(action=="init")return new {mode="database"};
                if(action=="dbReadFile")return await Task.Run(()=>ReportReader.Read((string)message["path"],(string)message["sheet"],(int?)message["maxRows"]??0,lifetime.Token),lifetime.Token);
                if(action=="dbOpenSource")return gateway.OpenReportSource((string)message["path"],(string)message["sheet"],(int?)message["row"]??1);
                throw new InvalidOperationException("未知数据库操作。");
            });
            window.Disconnected+=()=>lifetime.Cancel();
        }
        private void OpenSettings()
        {
            OpenPortable(message=>
            {
                string action=(string)message["action"];
                if(action=="init")return Task.FromResult<object>(new {mode="settings",settings,navigation=navigation.Settings});
                if(action=="saveSettings")
                {
                    var candidate=message["settings"].ToObject<ToolPreferences>();candidate.Normalize();
                    JsonConvert.PopulateObject(JsonConvert.SerializeObject(candidate),settings);bool saved=store.Save(settings);Ribbon.Refresh();
                    return Task.FromResult<object>(new {message=store.Status,saved,settings});
                }
                if(action=="navigationSettings"){navigation.ShowSettings();return Task.FromResult<object>(new {message="已打开导航布局设置。"});}
                throw new InvalidOperationException("未知操作。");
            });
        }
        private PortableWindow OpenPortable(Func<JObject,Task<object>> handler,Action close=null)
        {
            var window=new PortableWindow(Path.GetDirectoryName(ExcelDnaUtil.XllPath),handler);
            portableWindows.Add(window);
            window.CancelRequested+=()=>close?.Invoke();
            window.Closed+=()=>{portableWindows.Remove(window);close?.Invoke();if(!disposed)WorksheetFocus.Restore(ExcelDnaUtil.WindowHandle);};
            return window;
        }
        private void Show(Window window)
        {
            new WindowInteropHelper(window).Owner = ExcelDnaUtil.WindowHandle;
            windows.Add(window); try { window.ShowDialog(); } finally { windows.Remove(window); }
        }
        private void Announce(string text)
        {
            dynamic app = ExcelDnaUtil.Application;
            if (status == null) previousStatus = app.StatusBar;
            status = "GridPilot · " + text; app.StatusBar = status; statusTimer.Stop(); statusTimer.Start();
        }
        private void ClearStatus()
        {
            statusTimer.Stop(); if (status == null) return;
            try { dynamic app = ExcelDnaUtil.Application; if (Equals((object)app.StatusBar, status)) app.StatusBar = previousStatus ?? false; } catch { }
            status = null;
        }
        public void Dispose() { if (disposed) return; disposed = true;gateway.ReleaseUndoEvents(); focus?.Dispose(); foreach(var browser in portableWindows)browser.Dispose();portableWindows.Clear(); ClearStatus(); statusTimer.Dispose(); foreach (var window in new List<Window>(windows)) window.Close(); windows.Clear(); }
    }
}
