using System;
using System.Linq;
using Newtonsoft.Json.Linq;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        private object RangeOperation(Snapshot s, string op, JObject o)
        {
            object sheet = LiveSheet(s, o), range = null;
            try
            {
                string address = Required(o, "address"); var bounds = AiRange.Parse(address); range = ((dynamic)sheet).Range[address]; dynamic r = range;
                if (op == "range.read")
                {
                    int offset = Math.Max(0, (int?)o["offset"] ?? 0), limit = Math.Min(2000, Math.Max(1, (int?)o["limit"] ?? 200));
                    int width = bounds.LastColumn - bounds.Column + 1; var cells = new JArray();
                    for (long index = offset; index < Math.Min(bounds.Size, (long)offset + limit); index++) { int row = bounds.Row + (int)(index / width), column = bounds.Column + (int)(index % width); object c = ((dynamic)sheet).Cells[row, column]; try { cells.Add(new JObject { ["address"] = Address(row, column), ["value"] = JToken.FromObject(((dynamic)c).Value2 ?? ""), ["formula"] = (bool)((dynamic)c).HasFormula ? (string)((dynamic)c).Formula : null, ["numberFormat"] = (string)((dynamic)c).NumberFormat }); } finally { Release(c); } }
                    return new { cells, total = bounds.Size, nextOffset = offset + cells.Count < bounds.Size ? (int?)(offset + cells.Count) : null, live = true };
                }
                if (op == "range.trace") { object traced = null; try { traced = (string)o["direction"] == "dependents" ? r.DirectDependents : r.DirectPrecedents; return new { address = (string)((dynamic)traced).Address[false, false, 1, true] }; } finally { Release(traced); } }
                if ((bool)((dynamic)s.book).ReadOnly || (bool)((dynamic)sheet).ProtectContents) throw new InvalidOperationException("目标只读或受保护。");
                switch (op)
                {
                    case "range.write":
                        var data = o["data"] as JArray ?? throw new ArgumentException("缺少二维 data。"); int nr = bounds.LastRow - bounds.Row + 1, nc = bounds.LastColumn - bounds.Column + 1;
                        if (data.Count != nr || data.Any(row => !(row is JArray) || row.Count() != nc)) throw new ArgumentException("data 大小与 address 不匹配。");
                        bool formulas = (bool?)o["formulas"] == true;
                        foreach (var row in data) foreach (var token in row) { if (!(token is JValue)) throw new ArgumentException("单元格只能是基本值。"); if (formulas) ValidateFormula((string)token); }
                        for (int i = 0; i < nr; i++) for (int j = 0; j < nc; j++) { object cell = r.Cells[i + 1, j + 1]; try { if (formulas) ((dynamic)cell).Formula = (string)data[i][j]; else SetValue(cell, ((JValue)data[i][j]).Value); } finally { Release(cell); } }
                        break;
                    case "range.clear": switch ((string)o["what"] ?? "contents") { case "all": r.Clear(); break; case "formats": r.ClearFormats(); break; case "notes": r.ClearComments(); break; case "validation": r.Validation.Delete(); break; case "conditionalFormats": r.FormatConditions.Delete(); break; case "contents": r.ClearContents(); break; default: throw new ArgumentException("未知清空类型。"); } break;
                    case "range.insert": if ((string)o["axis"] == "rows") r.EntireRow.Insert(); else if ((string)o["axis"] == "columns") r.EntireColumn.Insert(); else r.Insert((string)o["shift"] == "right" ? -4161 : -4121); break;
                    case "range.delete": if ((string)o["axis"] == "rows") r.EntireRow.Delete(); else if ((string)o["axis"] == "columns") r.EntireColumn.Delete(); else r.Delete((string)o["shift"] == "left" ? -4159 : -4162); break;
                    case "range.copy":
                        object targetSheet = ((dynamic)s.book).Worksheets[(string)o["destinationSheet"] ?? (string)o["sheet"]], target = null;
                        try { string dest = Required(o, "destination"); AiRange.Parse(dest); target = ((dynamic)targetSheet).Range[dest]; r.Copy(); try { ((dynamic)target).PasteSpecial((int?)o["pasteType"] ?? -4104, (int?)o["pasteOperation"] ?? -4142, (bool?)o["skipBlanks"] ?? false, (bool?)o["transpose"] ?? false); } finally { ((dynamic)application()).CutCopyMode = false; } } finally { Release(target); Release(targetSheet); } break;
                    case "range.sort":
                        ((dynamic)s.book).Activate(); ((dynamic)sheet).Activate(); r.Select(); var gateway = new ExcelGateway(application); var captured = gateway.CaptureSort();
                        var options = new MultiSortOptions { HasHeader = (bool?)o["header"] ?? true, MatchCase = (bool?)o["matchCase"] ?? false, StrokeOrder = (bool?)o["strokeOrder"] ?? false, Levels = new System.Collections.Generic.List<SortLevel>() };
                        foreach (JObject level in o["levels"] as JArray ?? throw new ArgumentException("缺少 levels。")) options.Levels.Add(new SortLevel { Column = (int)level["column"] - 1, Mode = (string)level["mode"] ?? "values", Descending = (bool?)level["descending"] ?? false, TextAsNumbers = (bool?)level["textAsNumbers"] ?? false, Rules = (string)level["rules"] ?? "[中文数字]", Sequences = (string)level["sequences"] ?? "" });
                        return new { message = gateway.ApplySort(captured, gateway.PlanSort(captured, options)) };
                    case "range.replace":
                        string replacement = (string)o["replacement"] ?? ""; if (replacement.StartsWith("=")) ValidateFormula(replacement);
                        return new { replaced = (bool)r.Replace(Required(o, "find"), replacement, (bool?)o["whole"] == true ? 1 : 2, 1, (bool?)o["matchCase"] ?? false, false, false, false) };
                    case "range.merge": if ((bool?)o["merge"] ?? true) r.Merge((bool?)o["across"] ?? false); else r.UnMerge(); break;
                    case "range.transform":
                        switch (Required(o, "transform"))
                        {
                            case "deduplicate": r.RemoveDuplicates((o["columns"] as JArray ?? new JArray(1)).Values<int>().Cast<object>().ToArray(), (bool?)o["header"] == true ? 1 : 2); break;
                            case "fillDown": r.FillDown(); break; case "fillRight": r.FillRight(); break;
                            case "series": r.DataSeries((int?)o["rowcol"] ?? 2, (int?)o["type"] ?? -4132, (int?)o["date"] ?? 1, (double?)o["step"] ?? 1, o["stop"] == null ? Type.Missing : (object)(double)o["stop"], (bool?)o["trend"] ?? false); break;
                            case "textToColumns": r.TextToColumns(Destination: r.Cells[1, 1], DataType: 1, TextQualifier: 1, ConsecutiveDelimiter: false, Tab: (string)o["delimiter"] == "\t", Semicolon: (string)o["delimiter"] == ";", Comma: (string)o["delimiter"] == ",", Space: (string)o["delimiter"] == " ", Other: !new[] { "\t", ";", ",", " " }.Contains((string)o["delimiter"]), OtherChar: (string)o["delimiter"] ?? "|"); break;
                            default: throw new ArgumentException("未知 transform。");
                        } break;
                    case "range.format":
                        if (o["properties"] is JObject props) SetProperties(range, props, "NumberFormat HorizontalAlignment VerticalAlignment WrapText ShrinkToFit Orientation IndentLevel Locked FormulaHidden RowHeight ColumnWidth");
                        if (o["font"] is JObject font) { object f = r.Font; try { SetProperties(f, font, "Name Size Bold Italic Underline Strikethrough Color Superscript Subscript"); } finally { Release(f); } }
                        if (o["fill"] is JObject fill) { object f = r.Interior; try { SetProperties(f, fill, "Color Pattern PatternColor TintAndShade"); } finally { Release(f); } }
                        if (o["borders"] is JArray borders) foreach (JObject b in borders) { object border = r.Borders[(int)b["index"]]; try { SetProperties(border, (JObject)b["properties"], "LineStyle Weight Color TintAndShade"); } finally { Release(border); } }
                        if ((bool?)o["autoFit"] == true) r.Columns.AutoFit(); break;
                    case "range.validation":
                        foreach(string key in new[]{"formula1","formula2"}) if (((string)o[key] ?? "").StartsWith("=")) ValidateFormula((string)o[key]);
                        object validation = r.Validation; try { ((dynamic)validation).Delete(); if ((bool?)o["delete"] != true) { if (((string)o["formula1"] ?? "").StartsWith("=")) ValidateFormula((string)o["formula1"]); ((dynamic)validation).Add((int?)o["type"] ?? 3, (int?)o["alertStyle"] ?? 1, (int?)o["operator"] ?? 1, (string)o["formula1"] ?? "", (string)o["formula2"] ?? ""); if (o["properties"] is JObject vp) SetProperties(validation, vp, "IgnoreBlank InCellDropdown ShowInput ShowError InputTitle InputMessage ErrorTitle ErrorMessage"); } } finally { Release(validation); } break;
                    case "range.conditional_format":
                        foreach(string key in new[]{"formula1","formula2"}) if (((string)o[key] ?? "").StartsWith("=")) ValidateFormula((string)o[key]);
                        object conditions = r.FormatConditions, condition = null;
                        try
                        {
                            if ((bool?)o["delete"] == true) { ((dynamic)conditions).Delete(); break; }
                            string kind = (string)o["kind"] ?? "formula";
                            switch (kind)
                            {
                                case "colorScale": condition = ((dynamic)conditions).AddColorScale((int?)o["colors"] ?? 3); break;
                                case "dataBar": condition = ((dynamic)conditions).AddDatabar(); break;
                                case "iconSet": condition = ((dynamic)conditions).AddIconSetCondition(); if (o["iconSet"] != null) ((dynamic)condition).IconSet = ((dynamic)s.book).IconSets[(int)o["iconSet"]]; break;
                                case "top": condition = ((dynamic)conditions).AddTop10(); ((dynamic)condition).Rank = (int?)o["rank"] ?? 10; ((dynamic)condition).Percent = (bool?)o["percent"] ?? false; break;
                                case "duplicate": condition = ((dynamic)conditions).AddUniqueValues(); ((dynamic)condition).DupeUnique = (bool?)o["unique"] == true ? 0 : 1; break;
                                case "formula": ValidateFormula(Required(o, "formula1")); condition = ((dynamic)conditions).Add(2, Type.Missing, (string)o["formula1"]); break;
                                case "cellValue": condition = ((dynamic)conditions).Add(1, (int?)o["operator"] ?? 1, (string)o["formula1"], (string)o["formula2"]); break;
                                default: throw new ArgumentException("未知条件格式类型。");
                            }
                            if (o["fillColor"] != null) ((dynamic)condition).Interior.Color = AiRange.Color((string)o["fillColor"]);
                            if (o["fontColor"] != null) ((dynamic)condition).Font.Color = AiRange.Color((string)o["fontColor"]);
                            if (o["barColor"] != null) ((dynamic)condition).BarColor.Color = AiRange.Color((string)o["barColor"]);
                            if (o["criteria"] is JArray criteria) foreach (JObject c in criteria) { dynamic item = ((dynamic)condition).ColorScaleCriteria[(int)c["index"]]; try { item.Type = (int)c["type"]; if (c["value"] != null) item.Value = ((JValue)c["value"]).Value; item.FormatColor.Color = AiRange.Color((string)c["color"]); } finally { Release(item); } }
                        }
                        finally { Release(condition); Release(conditions); } break;
                    case "range.goal_seek": object changing = ((dynamic)sheet).Range[Required(o, "changingCell")]; try { return new { converged = (bool)r.GoalSeek((double)o["goal"], changing) }; } finally { Release(changing); }
                    default: throw new ArgumentException("未知范围操作。");
                }
                return new { updated = true, sheet = (string)o["sheet"], address };
            }
            finally { Release(range); Release(sheet); }
        }
    }
}
