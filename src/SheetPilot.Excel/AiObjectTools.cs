using System;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private object ObjectOperation(Snapshot s, string op, JObject o)
        {
            object sheet = null, collection = null, item = null, range = null;
            try
            {
                dynamic book = s.book;
                if (op.StartsWith("name."))
                {
                    collection = book.Names;
                    if (op == "name.list") { var list = new JArray(); foreach (dynamic name in (dynamic)collection) { try { list.Add(new JObject { ["name"] = (string)name.Name, ["refersTo"] = (string)name.RefersTo, ["visible"] = (bool)name.Visible }); } finally { Release(name); } } return new { names = list }; }
                    if (op == "name.delete") { item = ((dynamic)collection).Item(Required(o, "name")); ((dynamic)item).Delete(); }
                    else { string formula = Required(o, "refersTo"); ValidateFormula(formula); item = ((dynamic)collection).Add(Required(o, "name"), formula, (bool?)o["visible"] ?? true); }
                    return new { updated = true };
                }
                sheet = LiveSheet(s, o); dynamic w = sheet;
                if (op == "table.list" || op == "note.list" || op == "image.list")
                {
                    collection = op == "table.list" ? w.ListObjects : op == "note.list" ? w.Comments : w.Shapes;
                    var list = new JArray(); int offset = Math.Max(0, (int?)o["offset"] ?? 0), total = ((dynamic)collection).Count;
                    for (int i = offset + 1; i <= Math.Min(total, offset + 200); i++) { object obj = ((dynamic)collection).Item(i); try { if (op == "note.list") list.Add(new JObject { ["address"] = (string)((dynamic)obj).Parent.Address[false, false], ["text"] = (string)((dynamic)obj).Text() }); else list.Add(new JObject { ["name"] = (string)((dynamic)obj).Name, ["address"] = op == "table.list" ? (string)((dynamic)obj).Range.Address[false, false] : null }); } finally { Release(obj); } }
                    return new { items = list, total, nextOffset = offset + list.Count < total ? (int?)(offset + list.Count) : null };
                }
                if (op.StartsWith("table."))
                {
                    collection = w.ListObjects;
                    if (op == "table.create") { string address = Required(o, "address"); AiRange.Parse(address); range = w.Range[address]; item = ((dynamic)collection).Add(1, range, Type.Missing, (bool?)o["header"] == false ? 2 : 1); if (o["name"] != null) ((dynamic)item).Name = (string)o["name"]; }
                    else item = ((dynamic)collection).Item(Required(o, "name"));
                    if (op == "table.delete") { if ((bool?)o["keepData"] ?? true) ((dynamic)item).Unlist(); else ((dynamic)item).Delete(); return new { deleted = true }; }
                    if (o["address"] != null && op != "table.create") { range = w.Range[(string)o["address"]]; ((dynamic)item).Resize(range); }
                    if (o["properties"] is JObject props) SetProperties(item, props, "Name TableStyle ShowHeaders ShowTotals ShowTableStyleFirstColumn ShowTableStyleLastColumn ShowTableStyleRowStripes ShowTableStyleColumnStripes");
                    if (o["column"] != null) { object column = ((dynamic)item).ListColumns[(int)o["column"]]; try { if (o["formula"] != null) { ValidateFormula((string)o["formula"]); ((dynamic)column).DataBodyRange.Formula = (string)o["formula"]; } if (o["totals"] != null) ((dynamic)column).TotalsCalculation = (int)o["totals"]; if (o["columnName"] != null) ((dynamic)column).Name = (string)o["columnName"]; } finally { Release(column); } }
                    return new { name = (string)((dynamic)item).Name };
                }
                if (op == "image.insert") { string path = LocalPath(Required(o, "path")); item = w.Shapes.AddPicture(path, 0, -1, (double?)o["left"] ?? 10, (double?)o["top"] ?? 10, (double?)o["width"] ?? -1, (double?)o["height"] ?? -1); if (o["name"] != null) ((dynamic)item).Name = (string)o["name"]; return new { name = (string)((dynamic)item).Name }; }
                if (op == "image.delete") { item = w.Shapes[Required(o, "name")]; ((dynamic)item).Delete(); return new { deleted = true }; }
                string addr = Required(o, "address"); AiRange.Parse(addr); range = w.Range[addr];
                if (op.StartsWith("note.")) { ((dynamic)range).ClearComments(); if (op == "note.set") item = ((dynamic)range).AddComment((string)o["text"] ?? ""); }
                else if (op == "sparkline.delete") ((dynamic)range).SparklineGroups.Clear();
                else if (op == "sparkline.add") { string source = Required(o, "source"); AiRange.Parse(source); item = ((dynamic)range).SparklineGroups.Add((int?)o["type"] ?? 1, "'" + ((string)o["sheet"]).Replace("'", "''") + "'!" + source); }
                else throw new ArgumentException("未知对象操作。");
                return new { updated = true };
            }
            finally { Release(item); Release(collection); Release(range); Release(sheet); }
        }
    }
}
