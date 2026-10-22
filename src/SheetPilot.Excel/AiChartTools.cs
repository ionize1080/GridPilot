using System;
using System.Linq;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class AiOperations
    {
        public static readonly Dictionary<string, int> Excel2010Charts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) {
            {"column",51},{"bar",57},{"line",4},{"pie",5},{"scatter",-4169},{"area",1},
            {"xl3DArea",-4098},{"xl3DColumn",-4100},{"xl3DLine",-4101},{"xl3DPie",-4102},{"xlDoughnut",-4120},{"xlRadar",-4151},{"xlXYScatter",-4169},{"xlArea",1},{"xlLine",4},{"xlPie",5},{"xlBubble",15},
            {"xlColumnClustered",51},{"xlColumnStacked",52},{"xlColumnStacked100",53},{"xl3DColumnClustered",54},{"xl3DColumnStacked",55},{"xl3DColumnStacked100",56},
            {"xlBarClustered",57},{"xlBarStacked",58},{"xlBarStacked100",59},{"xl3DBarClustered",60},{"xl3DBarStacked",61},{"xl3DBarStacked100",62},
            {"xlLineStacked",63},{"xlLineStacked100",64},{"xlLineMarkers",65},{"xlLineMarkersStacked",66},{"xlLineMarkersStacked100",67},{"xlPieOfPie",68},{"xlPieExploded",69},{"xl3DPieExploded",70},{"xlBarOfPie",71},
            {"xlXYScatterSmooth",72},{"xlXYScatterSmoothNoMarkers",73},{"xlXYScatterLines",74},{"xlXYScatterLinesNoMarkers",75},{"xlAreaStacked",76},{"xlAreaStacked100",77},{"xl3DAreaStacked",78},{"xl3DAreaStacked100",79},
            {"xlDoughnutExploded",80},{"xlRadarMarkers",81},{"xlRadarFilled",82},{"xlSurface",83},{"xlSurfaceWireframe",84},{"xlSurfaceTopView",85},{"xlSurfaceTopViewWireframe",86},{"xlBubble3DEffect",87},
            {"xlStockHLC",88},{"xlStockOHLC",89},{"xlStockVHLC",90},{"xlStockVOHLC",91},{"xlCylinderColClustered",92},{"xlCylinderColStacked",93},{"xlCylinderColStacked100",94},{"xlCylinderBarClustered",95},{"xlCylinderBarStacked",96},{"xlCylinderBarStacked100",97},{"xlCylinderCol",98},
            {"xlConeColClustered",99},{"xlConeColStacked",100},{"xlConeColStacked100",101},{"xlConeBarClustered",102},{"xlConeBarStacked",103},{"xlConeBarStacked100",104},{"xlConeCol",105},{"xlPyramidColClustered",106},{"xlPyramidColStacked",107},{"xlPyramidColStacked100",108},{"xlPyramidBarClustered",109},{"xlPyramidBarStacked",110},{"xlPyramidBarStacked100",111},{"xlPyramidCol",112}
        };
        private static int ChartType(JToken token)
        {
            int type; if (token?.Type == JTokenType.Integer) type = (int)token; else if (!Excel2010Charts.TryGetValue((string)token ?? "column", out type)) throw new ArgumentException("未知图表类型，请查询 chart.create 帮助。");
            if (!Excel2010Charts.Values.Contains(type)) throw new ArgumentException("图表类型不属于 Excel 2010。"); return type;
        }
        private const string ChartProperties = "Name Text Caption Formula Left Top Width Height Position IncludeInLayout Visible Color ColorIndex RGB ThemeColor TintAndShade Transparency Weight DashStyle Style Size Bold Italic Underline Strikethrough Orientation HorizontalAlignment VerticalAlignment AutoScaleFont FontStyle Shadow GlowRadius SoftEdgeRadius HasTitle HasLegend ChartType ChartStyle PlotBy DisplayBlanksAs PlotVisibleOnly Rotation Elevation Perspective RightAngleAxes DepthPercent HeightPercent GapDepth AutoScaling RoundedCorners HasAxis AxisGroup HasDataLabels HasErrorBars Smooth InvertIfNegative Explosion MarkerStyle MarkerSize MarkerForegroundColor MarkerBackgroundColor MarkerForegroundColorIndex MarkerBackgroundColorIndex BorderColor SchemeColor Pattern PatternColor LineStyle MinimumScale MaximumScale MinimumScaleIsAuto MaximumScaleIsAuto MajorUnit MinorUnit MajorUnitIsAuto MinorUnitIsAuto ScaleType LogBase Crosses CrossesAt ReversePlotOrder TickLabelPosition MajorTickMark MinorTickMark HasMajorGridlines HasMinorGridlines CategoryType BaseUnit MajorUnitScale MinorUnitScale BaseUnitIsAuto DisplayUnit DisplayUnitCustom HasDisplayUnitLabel NumberFormat NumberFormatLinked Offset Spacing ReadingOrder MultiLevel TickLabelSpacing TickMarkSpacing HasMultiLevelLabels ShowValue ShowCategoryName ShowSeriesName ShowLegendKey ShowPercentage ShowBubbleSize Separator Type Order Period Forward Backward Intercept InterceptIsAuto DisplayEquation DisplayRSquared LegendKey BubbleScale GapWidth Overlap VaryByCategories HasDropLines HasHiLoLines HasUpDownBars SplitType SplitValue SecondPlotSize DoughnutHoleSize FirstSliceAngle EndStyle CapStyle JoinStyle BeginArrowheadStyle EndArrowheadStyle BeginArrowheadLength EndArrowheadLength BeginArrowheadWidth EndArrowheadWidth";
        private static readonly HashSet<string> ChartChildren = new HashSet<string>(("ChartArea PlotArea ChartTitle Legend Font Format Fill Line ForeColor BackColor Border Interior TickLabels AxisTitle MajorGridlines MinorGridlines DataLabel DisplayUnitLabel ErrorBars LeaderLines DataTable Walls Floor UpBars DownBars DropLines HiLoLines Shadow Glow SoftEdge TextFrame2 TextRange ParagraphFormat").Split(' '), StringComparer.OrdinalIgnoreCase);
        private object ChartOperation(Snapshot s, string op, JObject o)
        {
            object sheet = LiveSheet(s, o), charts = null, container = null, chart = null;
            var owned = new List<object>();
            try
            {
                charts = ((dynamic)sheet).ChartObjects();
                if (op == "chart.list") { var list = new JArray(); for (int i = 1; i <= (int)((dynamic)charts).Count; i++) { object co = ((dynamic)charts).Item(i); try { list.Add(ChartState(co)); } finally { Release(co); } } return new { charts = list }; }
                if (op == "chart.create")
                {
                    int type = ChartType(o["type"]); string address = Required(o, "address"); AiRange.Parse(address);
                    container = ((dynamic)charts).Add((double?)o["left"] ?? 30, (double?)o["top"] ?? 30, (double?)o["width"] ?? 480, (double?)o["height"] ?? 280);
                    if (o["name"] != null) ((dynamic)container).Name = (string)o["name"];
                    chart = ((dynamic)container).Chart; object source = ((dynamic)sheet).Range[address]; try { ((dynamic)chart).SetSourceData(source); ((dynamic)chart).ChartType = type; } finally { Release(source); }
                    if (o["title"] != null) { ((dynamic)chart).HasTitle = true; ((dynamic)chart).ChartTitle.Text = (string)o["title"]; }
                    return ChartState(container);
                }
                container = ((dynamic)charts).Item(Required(o, "chart")); chart = ((dynamic)container).Chart;
                if (op == "chart.delete") { ((dynamic)container).Delete(); return new { deleted = true }; }
                if (op == "chart.series")
                {
                    object series = null, collection = ((dynamic)chart).SeriesCollection();
                    try
                    {
                        series = (string)o["action"] == "add" ? ((dynamic)collection).NewSeries() : ((dynamic)collection).Item((int?)o["index"] ?? 1);
                        if ((string)o["action"] == "delete") { ((dynamic)series).Delete(); return new { deleted = true }; }
                        if (o["name"] != null) { if(((string)o["name"]).StartsWith("="))ValidateFormula((string)o["name"]); ((dynamic)series).Name = (string)o["name"]; }
                        foreach (string key in new[] { "values", "xValues", "bubbleSizes" }) if (o[key] != null) { object value; if (o[key] is JArray array) value = array.Select(t => ((JValue)t).Value).ToArray(); else { string addr = (string)o[key]; AiRange.Parse(addr); value = ((dynamic)sheet).Range[addr]; owned.Add(value); } ComSet(series, key == "values" ? "Values" : key == "xValues" ? "XValues" : "BubbleSizes", value); }
                        if (o["type"] != null) ((dynamic)series).ChartType = ChartType(o["type"]);
                        if (o["axisGroup"] != null) ((dynamic)series).AxisGroup = (int)o["axisGroup"];
                        return new { updated = true };
                    }
                    finally { Release(series); Release(collection); }
                }
                if (op == "chart.trendline" || op == "chart.errorbars")
                {
                    object series = ((dynamic)chart).SeriesCollection((int?)o["series"] ?? 1); owned.Add(series);
                    if (op == "chart.errorbars") { if ((bool?)o["delete"] == true) ((dynamic)series).HasErrorBars = false; else ((dynamic)series).ErrorBar((int?)o["direction"] ?? -4168, (int?)o["include"] ?? 1, (int?)o["type"] ?? 2, (double?)o["amount"] ?? 1, (double?)o["minus"] ?? 1); }
                    else { object trends = ((dynamic)series).Trendlines(); owned.Add(trends); if ((bool?)o["delete"] == true) { object trend = ((dynamic)trends).Item((int?)o["index"] ?? 1); owned.Add(trend); ((dynamic)trend).Delete(); } else { object trend = ((dynamic)trends).Add((int?)o["type"] ?? -4132); owned.Add(trend); if (o["properties"] is JObject props) SetProperties(trend, props, ChartProperties); } }
                    return new { updated = true };
                }
                object target = chart;
                foreach (string segment in ((string)o["path"] ?? "").Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var match = Regex.Match(segment, @"^(Axes|SeriesCollection|Points|DataLabels|Trendlines|ChartGroups|LegendEntries)\(\s*(\w+)\s*(?:,\s*(\w+)\s*)?\)$", RegexOptions.IgnoreCase);
                    if (match.Success) { object[] args = match.Groups[3].Success ? new object[] { ChartIndex(match.Groups[2].Value), ChartIndex(match.Groups[3].Value) } : new object[] { ChartIndex(match.Groups[2].Value) }; target = target.GetType().InvokeMember(match.Groups[1].Value, BindingFlags.GetProperty | BindingFlags.InvokeMethod, null, target, args); }
                    else { if (!ChartChildren.Contains(segment)) throw new ArgumentException("未知图表元素路径：" + segment); target = ComGet(target, segment); }
                    owned.Add(target);
                }
                if (op == "chart.inspect")
                {
                    if (o["properties"] is JArray properties) { var result = new JObject(); foreach (string property in properties.Values<string>()) { if (!ChartProperties.Split(' ').Contains(property, StringComparer.OrdinalIgnoreCase)) throw new ArgumentException("未知图表属性。"); object value = ComGet(target, property); if (System.Runtime.InteropServices.Marshal.IsComObject(value)) { Release(value); throw new ArgumentException("该属性是对象，请指定元素路径。"); } result[property] = value == null ? JValue.CreateNull() : JToken.FromObject(value); } return result; }
                    return ChartState(container);
                }
                var edits = o["properties"] as JObject ?? throw new ArgumentException("缺少 properties。");
                if (edits["Formula"] != null) ValidateFormula((string)edits["Formula"]);
                if (edits["ChartType"] != null) edits["ChartType"] = ChartType(edits["ChartType"]);
                SetProperties(target, edits, ChartProperties); return new { updated = true };
            }
            finally { foreach (object value in owned.AsEnumerable().Reverse()) Release(value); Release(chart); Release(container); Release(charts); Release(sheet); }
        }
        private static int ChartIndex(string text) { switch(text.ToLowerInvariant()){case "xlcategory":case "xlprimary":return 1;case "xlvalue":case "xlsecondary":return 2;case "xlseriesaxis":return 3;default:return int.Parse(text,System.Globalization.CultureInfo.InvariantCulture);} }
    }
}
