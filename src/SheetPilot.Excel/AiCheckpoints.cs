using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private bool automatic, checkpointBefore, newCheckpoint = true;
        private string checkpointFolder;
        public object Approval(bool enabled)
        {
            if (enabled && !automatic) newCheckpoint = true;
            automatic = enabled;
            return CheckpointState();
        }
        public object CheckpointState() => new { automatic, available = undo != null, position = checkpointBefore ? "before" : "after", folder = checkpointFolder,
            message = automatic ? "自动批准已开启；首次修改前保存完整备份。" : "逐次批准；修改前后均保留快照。" };
        private string SaveCheckpoint(Proposal[] batch, string label, string folder = null)
        {
            folder = folder ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "GridPilot", "Recovery", DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            var books = batch.Select(p => Get(p.snapshot).book).Distinct().ToArray();
            for (int i = 0; i < books.Length; i++)
            {
                dynamic book = books[i];
                int format = Convert.ToInt32(book.FileFormat);
                string extension = format == 52 ? ".xlsm" : format == 50 ? ".xlsb" : format == 56 || format == -4143 ? ".xls" : format == 51 ? ".xlsx" : null;
                if (extension == null) throw new InvalidOperationException("请先把工作簿另存为 xlsx、xlsm、xlsb 或 xls，再使用 AI 写入，以便完整备份。");
                string file = Path.Combine(folder, label + "-" + (i + 1) + extension);
                dynamic app=application(); bool events=app.EnableEvents;
                try { app.EnableEvents=false; book.SaveCopyAs(file); } finally { app.EnableEvents=events; }
                if (!File.Exists(file) || new FileInfo(file).Length == 0) throw new IOException("工作簿备份未完成，未继续修改。");
            }
            File.WriteAllText(Path.Combine(folder,"README.txt"),"GridPilot 本地完整工作簿备份\r\nbefore-* 为本次修改前；after-* 为各次成功修改后的完整副本，按修改时间选择。\r\n界面切换仅还原 AI 触及的范围或对象；关闭助手后可直接打开这些文件手工恢复。\r\n备份含工作簿原始数据，请自行保管；不会上传至模型或 GitHub。\r\n");
            return folder;
        }
        private Proposal[] CombineCheckpoint(Proposal[] batch)
        {
            var result = new List<Proposal>();
            var cells = new Dictionary<Tuple<object,string,string>,Proposal>();
            foreach (var p in batch)
            {
                if (p.edits == null) { result.Add(new Proposal{id=p.id,snapshot=p.snapshot,sheet=p.sheet,address=p.address,kind=p.kind,chartName=p.chartName,style=p.style,beforeStyle=p.beforeStyle,afterStyle=p.afterStyle,sourceFormula=p.sourceFormula,afterSourceFormula=p.afterSourceFormula,options=p.options,chartState=p.chartState}); continue; }
                foreach (var e in p.edits)
                {
                    var key=Tuple.Create(Get(p.snapshot).book,p.sheet,e.address);
                    cells.TryGetValue(key,out var old);
                    var first = old?.edits[0] ?? e;
                    cells[key]=new Proposal { id = p.id, snapshot = p.snapshot, sheet = p.sheet, address = e.address, kind = p.kind,
                        edits = new[] { new Edit { address = e.address, formula = e.formula, value = e.value, beforeFormula = first.beforeFormula, beforeValue = first.beforeValue, hadFormula = first.hadFormula, afterFormula = e.afterFormula, afterHasFormula=e.afterHasFormula } } };
                }
            }
            return cells.Values.Concat(result.Where(p=>p.kind=="format")).Concat(result.Where(p=>p.kind!="format")).ToArray();
        }
        private void CompleteCheckpointSources(Proposal[] checkpoint)
        {
            foreach(var p in checkpoint.Where(p=>p.kind=="chart"||p.kind=="pivot"))
            {
                object sheet=Worksheet(Get(p.snapshot),p.sheet),range=null;
                try
                {
                    range=((dynamic)sheet).Range[p.address];p.afterSourceFormula=((dynamic)range).Formula;
                    var bounds=AiRange.Parse(p.address);var array=p.afterSourceFormula as Array;
                    object baseline=array==null?p.afterSourceFormula:array.Clone();
                    foreach(var q in checkpoint.Where(q=>q.edits!=null&&q.sheet==p.sheet&&ReferenceEquals(Get(q.snapshot).book,Get(p.snapshot).book)))
                    foreach(var e in q.edits)
                    {
                        var cell=AiRange.Parse(e.address);if(cell.Row<bounds.Row||cell.Row>bounds.LastRow||cell.Column<bounds.Column||cell.Column>bounds.LastColumn)continue;
                        if(baseline is Array values)values.SetValue(e.beforeFormula,cell.Row-bounds.Row+values.GetLowerBound(0),cell.Column-bounds.Column+values.GetLowerBound(1));else baseline=e.beforeFormula;
                    }
                    p.sourceFormula=baseline;
                }finally{Release(range);Release(sheet);}
            }
        }
        private void CheckCheckpointExtension(Proposal[] batch)
        {
            if (!automatic || newCheckpoint || undo == null) return;
            if (checkpointBefore) throw new InvalidOperationException("当前为变更前快照，请先切换至变更后再继续自动批准。");
            if (batch.Any(p => !ReferenceEquals(Get(p.snapshot).book, Get(undo[0].snapshot).book)))
                throw new InvalidOperationException("自动批准绑定当前工作簿。请关闭再开启自动批准，为另一工作簿建立新快照。");
            foreach (var p in undo) CheckProposal(p, true);
            // Content can be coalesced cell-by-cell. Overlapping format patches need
            // separate checkpoints so mixed styles are never flattened on restore.
            foreach (var p in batch.Where(p => p.kind == "format")) foreach (var old in undo.Where(x => x.kind == "format"))
            {
                if (old.sheet != p.sheet) continue;
                var a = AiRange.Parse(p.address); var b = AiRange.Parse(old.address);
                if (a.Row <= b.LastRow && b.Row <= a.LastRow && a.Column <= b.LastColumn && b.Column <= a.LastColumn)
                    throw new InvalidOperationException("此格式与当前快照中的格式修改重叠。请关闭自动批准后单独应用，保留独立恢复点。");
            }
        }
        public object RestoreCheckpoint(bool before)
        {
            if (undo == null) throw new InvalidOperationException("还没有变更快照。");
            if (checkpointBefore == before) return CheckpointState();
            foreach (var p in undo) CheckProposal(p, before);
            dynamic app = application(); bool events = app.EnableEvents, screen = app.ScreenUpdating;
            var written = new List<Proposal>();
            try
            {
                app.EnableEvents = false; app.ScreenUpdating = false;
                foreach (var p in before ? undo.Reverse() : undo) { written.Add(p); WriteProposal(p, before); }
            }
            catch (Exception ex)
            {
                var failures = new List<string>();
                foreach (var p in written.AsEnumerable().Reverse()) try { WriteProposal(p, !before); } catch (Exception restore) { failures.Add(restore.Message); }
                throw new InvalidOperationException(failures.Count == 0 ? "切换失败，已恢复当前状态：" + ex.Message : "切换未完全恢复，请使用本地完整备份：" + checkpointFolder);
            }
            finally { app.EnableEvents = events; app.ScreenUpdating = screen; }
            checkpointBefore = before; automatic = false;
            return new { message = before ? "已还原变更前；自动批准已关闭。" : "已恢复变更后。", checkpoint = CheckpointState() };
        }
    }
}
