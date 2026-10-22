using System;
using System.IO;
using System.Linq;
using System.Threading;
using Newtonsoft.Json.Linq;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private object PluginOperation(Snapshot s, string op, JObject o)
        {
            ((dynamic)s.book).Activate(); var gateway = new ExcelGateway(application); object sheet = null, range = null;
            try
            {
                if (op == "plugin.rename" || op == "plugin.rename_preview")
                {
                    var settings = o["settings"]?.ToObject<RenameOptions>() ?? new RenameOptions(); settings.AllWorkbooks = false;
                    var source = gateway.Capture(false); var selected = o["sheets"] is JArray names ? new System.Collections.Generic.HashSet<string>(names.Values<string>()) : null;
                    var excluded = new System.Collections.Generic.HashSet<string>(source.Sheets.Where(sh => selected != null && !selected.Contains(sh.Name)).Select(RenamePlanner.Key));
                    var plan = RenamePlanner.Build(source.Sheets, settings, excluded);
                    if (op == "plugin.rename_preview") return new { changes = plan.Select(p => new { from = p.Source.Name, to = p.NewName, selected = p.Selected, error = p.Error }) };
                    if(plan.Any(p=>p.Error!=null))throw new ArgumentException("重命名预览存在冲突："+string.Join("；",plan.Where(p=>p.Error!=null).Select(p=>p.Error)));
                    return SheetOperation(s,"sheet.batch_rename",new JObject{["changes"]=new JArray(plan.Where(p=>p.Changed).Select(p=>new JObject{["from"]=p.Source.Name,["to"]=p.NewName}))});
                }
                if (op == "plugin.sort_sheets") return new { message = gateway.SortSheets(gateway.Capture(false).Sheets, o["settings"]?.ToObject<SortOptions>() ?? new SortOptions { Natural = true }) };
                sheet = LiveSheet(s, o); ((dynamic)sheet).Activate(); string address = Required(o, "address"); AiRange.Parse(address); range = ((dynamic)sheet).Range[address]; ((dynamic)range).Select();
                if (op == "plugin.convert") return new { message = gateway.ConvertSelection((string)o["to"] == "text", o["settings"]?.ToObject<ToolPreferences>() ?? new ToolPreferences()) };
                if (op == "plugin.copy") { var result = gateway.CopySelection((bool?)o["keepInternalFormulas"] ?? true, o["settings"]?.ToObject<CopyOptions>() ?? new CopyOptions()); SaveVersion(result.Workbook, "原样复制结果"); return new { message = result.Summary }; }
                if (op == "plugin.segment")
                {
                    var settings = o["settings"]?.ToObject<SegmentationOptions>() ?? new SegmentationOptions();
                    var segmenter = new TextSegmenter(Path.Combine(Path.GetDirectoryName(typeof(AiOperations).Assembly.Location), "Resources"));
                    var result = gateway.CaptureTextSelection().Select(input => new TextOutput { Source = input, Words = segmenter.Split(input.Text, settings, CancellationToken.None) }).ToArray();
                    gateway.WriteTextOutput(TextBatch.Layout(result, settings)); object output = ((dynamic)application()).ActiveWorkbook; try { SaveVersion(output, "分词结果"); } finally { Release(output); }
                    return new { message = "已生成分词结果工作簿。", rows = result.Length };
                }
                if (op == "plugin.simple_mask")
                {
                    var snapshot = gateway.CaptureSimpleMask(CancellationToken.None); var result = SimpleMask.Build(snapshot, o["settings"]?.ToObject<SimpleMaskOptions>() ?? new SimpleMaskOptions(), CancellationToken.None);
                    if (o["path"] != null) return new { path = gateway.ExportSimpleMask(snapshot, result, LocalPath((string)o["path"]), CancellationToken.None) };
                    return new { message = gateway.ApplySimpleMask(snapshot, result, true, CancellationToken.None) };
                }
                throw new ArgumentException("未知插件操作。");
            }
            finally { gateway.ReleaseUndoEvents(); Release(range); Release(sheet); }
        }
    }
}
