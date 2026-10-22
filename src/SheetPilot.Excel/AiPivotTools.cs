using System;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private object PivotOperation(Snapshot s, string op, JObject o)
        {
            object sheet = LiveSheet(s, o), pivot = null, field = null; var owned = new List<object>();
            try
            {
                dynamic book = s.book, w = sheet;
                if (op == "slicer.delete") { object cache = book.SlicerCaches[Required(o, "cache")]; owned.Add(cache); if (o["name"] == null) ((dynamic)cache).Delete(); else { object slicer = ((dynamic)cache).Slicers[(string)o["name"]]; owned.Add(slicer); ((dynamic)slicer).Delete(); } return new { deleted = true }; }
                if (op == "pivot.list") { object pivots = w.PivotTables(); owned.Add(pivots); var list = new JArray(); for (int i = 1; i <= (int)((dynamic)pivots).Count; i++) { object p = ((dynamic)pivots).Item(i); try { list.Add(new JObject { ["name"] = (string)((dynamic)p).Name, ["address"] = (string)((dynamic)p).TableRange2.Address[false, false] }); } finally { Release(p); } } return new { pivots = list }; }
                if (op == "pivot.create")
                {
                    string address = Required(o, "address"); AiRange.Parse(address); object range = w.Range[address]; owned.Add(range);
                    object destinationSheet = book.Worksheets[Required(o, "destinationSheet")]; owned.Add(destinationSheet); string dest = (string)o["destination"] ?? "A3"; AiRange.Parse(dest); object destination = ((dynamic)destinationSheet).Range[dest]; owned.Add(destination);
                    object caches = book.PivotCaches(); owned.Add(caches); object cache = ((dynamic)caches).Create(1, ((dynamic)range).Address[true, true, 1, true], 3); owned.Add(cache);
                    pivot = ((dynamic)cache).CreatePivotTable(destination, (string)o["name"] ?? "GP_Pivot_" + Guid.NewGuid().ToString("N").Substring(0, 8));
                    foreach (var pair in new[] { Tuple.Create("rows", 1), Tuple.Create("columns", 2), Tuple.Create("filters", 3) }) { int index = 0; foreach (string name in (o[pair.Item1] as JArray ?? new JArray()).Values<string>()) { object f = ((dynamic)pivot).PivotFields(name); try { ((dynamic)f).Orientation = pair.Item2; ((dynamic)f).Position = ++index; } finally { Release(f); } } }
                    int vi = 0; foreach (JObject value in o["values"] as JArray ?? new JArray()) { object f = ((dynamic)pivot).PivotFields(Required(value, "field")); try { object added = ((dynamic)pivot).AddDataField(f, (string)value["caption"] ?? "Value_" + (++vi), PivotFunction(value["function"])); Release(added); } finally { Release(f); } }
                    return new { name = (string)((dynamic)pivot).Name };
                }
                pivot = w.PivotTables(Required(o, "pivot")); dynamic pt = pivot;
                if (op == "slicer.add")
                {
                    object caches = book.SlicerCaches; owned.Add(caches);
                    bool timeline=(bool?)o["timeline"]==true;
                    if(timeline && double.Parse((string)((dynamic)application()).Version,System.Globalization.CultureInfo.InvariantCulture)<15)throw new InvalidOperationException("时间线需要 Excel 2013 或更新版本；Excel 2010 可使用普通切片器。");
                    object cache = timeline ? ((dynamic)caches).Add2(pivot, Required(o,"field"),Type.Missing,2) : ((dynamic)caches).Add(pivot, Required(o, "field")); owned.Add(cache);
                    object slicers = ((dynamic)cache).Slicers; owned.Add(slicers); object slicer = ((dynamic)slicers).Add(sheet, Type.Missing, (string)o["name"] ?? "GP_Slicer_" + Guid.NewGuid().ToString("N").Substring(0, 6), (string)o["caption"] ?? (string)o["field"], (double?)o["top"] ?? 20, (double?)o["left"] ?? 520, (double?)o["width"] ?? 150, (double?)o["height"] ?? 200); owned.Add(slicer);
                    return new { cache = (string)((dynamic)cache).Name, name = (string)((dynamic)slicer).Name };
                }
                if (op == "pivot.inspect")
                {
                    object fields = pt.PivotFields(); owned.Add(fields); var result = new JArray();
                    for (int i = 1; i <= (int)((dynamic)fields).Count; i++) { object f = ((dynamic)fields).Item(i); try { var info = new JObject { ["name"] = (string)((dynamic)f).Name, ["orientation"] = (int)((dynamic)f).Orientation, ["caption"] = (string)((dynamic)f).Caption }; if (o["field"] != null && (string)((dynamic)f).Name == (string)o["field"]) { object items = ((dynamic)f).PivotItems(); try { int offset = Math.Max(0, (int?)o["offset"] ?? 0), total = ((dynamic)items).Count; var page = new JArray(); for (int j = offset + 1; j <= Math.Min(total, offset + 200); j++) { object item = ((dynamic)items).Item(j); try { page.Add(new JObject { ["name"] = (string)((dynamic)item).Name, ["visible"] = (bool)((dynamic)item).Visible }); } finally { Release(item); } } info["items"] = page; info["totalItems"] = total; info["nextOffset"] = offset + page.Count < total ? (JToken)(offset + page.Count) : JValue.CreateNull(); } finally { Release(items); } } result.Add(info); } finally { Release(f); } }
                    return new { name = (string)pt.Name, address = (string)pt.TableRange2.Address[false, false], fields = result };
                }
                if (op == "pivot.delete") { pt.TableRange2.Clear(); return new { deleted = true }; }
                if (op == "pivot.refresh") { object cache = pt.PivotCache(); owned.Add(cache); if ((int)((dynamic)cache).SourceType != 1) throw new InvalidOperationException("只刷新本地工作表透视缓存，不允许外部数据连接。"); return new { refreshed = (bool)pt.RefreshTable() }; }
                if (op == "pivot.edit")
                {
                    if (o["layout"] != null) pt.RowAxisLayout((int)o["layout"]);
                    if (o["repeatLabels"] != null) pt.RepeatAllLabels((bool)o["repeatLabels"] ? 2 : 1);
                    if (o["properties"] is JObject props) SetProperties(pivot, props, "Name RowGrand ColumnGrand TableStyle2 ShowTableStyleRowHeaders ShowTableStyleColumnHeaders ShowTableStyleRowStripes ShowTableStyleColumnStripes DisplayFieldCaptions ShowDrillIndicators InGridDropZones DisplayNullString NullString DisplayErrorString ErrorString PreserveFormatting HasAutoFormat MergeLabels CompactLayoutRowHeader CompactLayoutColumnHeader SubtotalLocation");
                    if (o["source"] != null) { string source = (string)o["source"]; AiRange.Parse(source); object range = w.Range[source]; owned.Add(range); object caches = book.PivotCaches(); owned.Add(caches); object cache = ((dynamic)caches).Create(1, ((dynamic)range).Address[true, true, 1, true], 3); owned.Add(cache); pt.ChangePivotCache(cache); }
                    return new { updated = true };
                }
                if (op == "pivot.calculated")
                {
                    object calculated;
                    if (o["field"] != null) { field = pt.PivotFields((string)o["field"]); calculated = ((dynamic)field).CalculatedItems(); }
                    else calculated = pt.CalculatedFields(); owned.Add(calculated);
                    if ((bool?)o["delete"] == true) { object item = ((dynamic)calculated).Item(Required(o, "name")); owned.Add(item); ((dynamic)item).Delete(); }
                    else { string formula = Required(o, "formula"); ValidateFormula(formula); object item = ((dynamic)calculated).Add(Required(o, "name"), formula); owned.Add(item); }
                    return new { updated = true };
                }
                field = pt.PivotFields(Required(o, "field")); dynamic pf = field;
                switch (op)
                {
                    case "pivot.field":
                        if ((bool?)o["addValue"] == true) { object value = pt.AddDataField(field, (string)o["caption"] ?? "Value_" + Guid.NewGuid().ToString("N").Substring(0, 6), PivotFunction(o["function"])); owned.Add(field); field = value; pf = value; }
                        if (o["orientation"] != null) pf.Orientation = (int)o["orientation"];
                        if (o["position"] != null) pf.Position = (int)o["position"];
                        if (o["function"] != null && (bool?)o["addValue"] != true) pf.Function = PivotFunction(o["function"]);
                        if (o["properties"] is JObject fp) SetProperties(field, fp, "Caption NumberFormat Calculation BaseField BaseItem ShowAllItems LayoutForm LayoutBlankLine LayoutSubtotalLocation RepeatLabels IncludeNewItemsInFilter EnableMultiplePageItems");
                        if (o["subtotals"] is JArray subs) { if (subs.Count != 12) throw new ArgumentException("subtotals 需要 12 个布尔值。"); pf.Subtotals = subs.Values<bool>().Cast<object>().ToArray(); }
                        break;
                    case "pivot.filter":
                        if ((bool?)o["clear"] == true) pf.ClearAllFilters();
                        if (o["page"] != null) pf.CurrentPage = (string)o["page"];
                        if (o["visibleItems"] is JArray visible)
                        {
                            var chosen = new HashSet<string>(visible.Values<string>()); if (chosen.Count == 0) throw new ArgumentException("至少保留一个项目。"); object items = pf.PivotItems(); owned.Add(items);
                            var all = new List<object>(); for (int i = 1; i <= (int)((dynamic)items).Count; i++) { object item = ((dynamic)items).Item(i); all.Add(item); owned.Add(item); }
                            if (chosen.Any(name => !all.Any(item => (string)((dynamic)item).Name == name))) throw new ArgumentException("筛选项目不存在。");
                            if ((int)pf.Orientation == 3) pf.EnableMultiplePageItems = true;
                            foreach (var item in all.Where(item => chosen.Contains((string)((dynamic)item).Name))) ((dynamic)item).Visible = true;
                            foreach (var item in all.Where(item => !chosen.Contains((string)((dynamic)item).Name))) ((dynamic)item).Visible = false;
                        }
                        if (o["type"] != null) { object filters = pf.PivotFilters; owned.Add(filters); object dataField = o["dataField"] == null ? null : pt.DataFields((string)o["dataField"]); if (dataField != null) owned.Add(dataField); object filter = ((dynamic)filters).Add(Type: (int)o["type"], DataField: dataField ?? Type.Missing, Value1: (o["value1"] as JValue)?.Value ?? Type.Missing, Value2: (o["value2"] as JValue)?.Value ?? Type.Missing); owned.Add(filter); }
                        break;
                    case "pivot.sort": pf.AutoSort((bool?)o["descending"] == true ? 2 : 1, (string)o["by"] ?? (string)o["field"]); break;
                    case "pivot.group":
                        object selection = o["address"] == null ? pf.DataRange.Cells[1, 1] : w.Range[(string)o["address"]]; owned.Add(selection);
                        if ((bool?)o["ungroup"] == true) ((dynamic)selection).Ungroup();
                        else ((dynamic)selection).Group((o["start"] as JValue)?.Value ?? true, (o["end"] as JValue)?.Value ?? true, (double?)o["by"] ?? 1, o["periods"] is JArray periods ? (object)periods.Values<bool>().Cast<object>().ToArray() : Type.Missing); break;
                    default: throw new ArgumentException("未知透视表操作。");
                }
                return new { updated = true };
            }
            finally { if (!owned.Contains(field)) Release(field); foreach (object value in owned.AsEnumerable().Reverse()) Release(value); Release(pivot); Release(sheet); }
        }
        private static int PivotFunction(JToken value)
        {
            if (value?.Type == JTokenType.Integer) return (int)value;
            var functions = new Dictionary<string, int> { { "sum", -4157 }, { "count", -4112 }, { "average", -4106 }, { "min", -4139 }, { "max", -4136 }, { "product", -4149 }, { "countNumbers", -4113 }, { "stdev", -4155 }, { "stdevp", -4156 }, { "var", -4164 }, { "varp", -4165 } };
            if (!functions.TryGetValue((string)value ?? "sum", out int result)) throw new ArgumentException("未知汇总函数。"); return result;
        }
    }
}
