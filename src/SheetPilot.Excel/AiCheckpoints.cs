using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private bool checkpointBefore, newCheckpoint = true;
        private sealed class ApprovalGrant { public object book; public string scope; }
        private readonly List<ApprovalGrant> approvals = new List<ApprovalGrant>();
        private ApprovalGrant CurrentApproval()
        {
            object book=null,books=null;
            try {
                books=((dynamic)application()).Workbooks;var opened=new List<object>();
                try { for(int i=1;i<=(int)((dynamic)books).Count;i++)opened.Add(((dynamic)books).Item[i]);
                    foreach(var grant in approvals.ToArray())if(!opened.Any(b=>ReferenceEquals(b,grant.book))){approvals.Remove(grant);Release(grant.book);}
                } finally { foreach(var b in opened)Release(b); }
                book=((dynamic)application()).ActiveWorkbook;return approvals.FirstOrDefault(g=>ReferenceEquals(g.book,book));
            } finally { Release(book);Release(books); }
        }
        private bool automatic { get => CurrentApproval()!=null; set { if(!value)Approval(false); } }
        private bool AutomaticFor(object book) => ReferenceEquals(CurrentApproval()?.book,book);
        private string checkpointFolder;
        public object Approval(bool enabled, string scope = "session")
        {
            if(scope!="session"&&scope!="workbook")throw new InvalidOperationException("请选择本次对话或当前工作簿。");
            var current=CurrentApproval();
            if(current!=null){approvals.Remove(current);Release(current.book);}
            if(enabled){object book=((dynamic)application()).ActiveWorkbook;if(book==null)throw new InvalidOperationException("请先打开工作簿。");approvals.Add(new ApprovalGrant{book=book,scope=scope});if(current==null)newCheckpoint=true;}
            return CheckpointState();
        }
        public object ConversationChanged()
        {
            foreach(var grant in approvals.Where(g=>g.scope=="session").ToArray()){approvals.Remove(grant);Release(grant.book);}
            return CheckpointState();
        }
        public object CheckpointState() { var grant=CurrentApproval();return new { automatic=grant!=null, approvalScope=grant?.scope, approvalBook=grant==null?null:(string)((dynamic)grant.book).Name, available = undo != null, position = checkpointBefore ? "before" : "after", folder = checkpointFolder,
            message = grant!=null ? "自动批准已开启；修改前后保存完整备份。" : "逐次批准；修改前后均保留快照。" }; }
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
        private void CheckCheckpointExtension(Proposal[] batch,bool batchAutomatic)
        {
            if (!batchAutomatic || newCheckpoint || undo == null) return;
            if (checkpointBefore) { newCheckpoint=true;return; }
            if (batch.Any(p => !ReferenceEquals(Get(p.snapshot).book, Get(undo[0].snapshot).book)))
                {newCheckpoint=true;return;}
            try { foreach (var p in undo) CheckProposal(p, true); }
            catch { newCheckpoint=true;return; } // Keep older full backups; start from the current live workbook.
            // Content can be coalesced cell-by-cell. Overlapping format patches need
            // separate checkpoints so mixed styles are never flattened on restore.
            foreach (var p in batch.Where(p => p.kind == "format")) foreach (var old in undo.Where(x => x.kind == "format"))
            {
                if (old.sheet != p.sheet) continue;
                var a = AiRange.Parse(p.address); var b = AiRange.Parse(old.address);
                if (a.Row <= b.LastRow && b.Row <= a.LastRow && a.Column <= b.LastColumn && b.Column <= a.LastColumn)
                    {newCheckpoint=true;return;}
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
            checkpointBefore = before;if(CurrentApproval()?.scope!="workbook")automatic = false;
            return new { message = before ? "已还原变更前。" : "已恢复变更后。", checkpoint = CheckpointState() };
        }
    }
}
