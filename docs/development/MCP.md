# Excel MCP：精简入口，按需发现

[文档首页](../README.md) · v2.1.0 · 2026-10-22

13 个公开工具：原有 10 个读取／建议工具，加 `excel_context`、`excel_help` 和 `excel_operation`。高级目录按需返回，不放进每轮工具描述；当前包含 99 个操作。目录不随每轮发送。

## 调用约定

1. 自动/所有工作簿模式先调用 `excel_context({snapshotId:目录标识,bookId:目录中的标识,mode:"capture",sheet?:表名,address?:范围})`，后续数据和修改工具必须使用返回的子 `snapshotId`。省略表/区域时完整扫描含隐藏表的工作簿；mode=bind 只绑定元数据/VBA。目录标识不能当作已扫描数据。
2. `excel_help({operation:"pivot.field"})` 查询参数，或 `excel_help({topic:"pivot"})` 列出操作。
3. `excel_operation({snapshotId,operation,options})` 读取或生成建议。修改成功必须检查 `applied=true`；否则等待用户应用。
4. 分页读取，使用 1-based 索引，只有 offset 为 0-based。值不执行公式；公式以 = 开头。
5. 审批、恢复、运行宏不在模型工具列表。网络／Shell VBA 即使开启自动批准也必须用户逐项手动批准。

MCP 使用官方 SDK stdio，提供 `gridpilot://guide` Resource 和 `excel-workflow` Prompt。没有 HTTP/SSE 或网络文件能力。本地启动环境 `GRIDPILOT_MCP_READ_ONLY=1` 拒绝所有写操作；`GRIDPILOT_MCP_TOOLS=excel_read,excel_help,...` 过滤入口，列出与调用两端均检查。

## 操作目录

以下参数与源码 catalog 同步；详细例子可通过 excel_help 获取。Excel 枚举使用本机 Excel 的数字常量；chart type 可用 Excel 2010 的名称。

### workbook

- `workbook.describe`：{}.
- `workbook.list`：directory?: local directory; omit to list open books.
- `workbook.create`：{}. Creates an empty workbook. Recapture context for subsequent calls.
- `workbook.open`：path: local workbook path. Macros disabled; external links not updated. Recapture afterward.
- `workbook.import`：base64: string, extension: .xlsx|.xlsm|.xlsb|.xls. Max 32 MB. No network paths.
- `workbook.export`：path?: new local output path; omit for base64 up to 4 MB. Never overwrites an existing file.
- `workbook.settings`：properties?: {Title,Subject,Author,Keywords,Comments,...}; calculation?: -4105(auto)|-4135(manual)|2; date1904?: bool; protect?: bool; password?: string.

### sheet

- `sheet.describe`：sheet: string.
- `sheet.create`：name: string.
- `sheet.rename`：sheet: string, name: string.
- `sheet.copy`：sheet: string, name: string.
- `sheet.delete`：sheet: string. Full workbook snapshot first; cannot delete the last visible worksheet.
- `sheet.batch_rename`：changes: [{from,to},...]. Validates all names; supports cycles and swaps.
- `sheet.layout`：sheet; visible?: -1|0|2; tabColor?: #RRGGBB; protect?,password?; freezeRows?,freezeColumns?; address?,rowHeight?,columnWidth?,hidden?,axis?: rows|columns,group?,autoFilter?; print?: Excel PageSetup properties.

### range

- `range.read`：sheet,address; offset?: 0-based cell offset; limit?: 1..2000. Live read includes blank cells and nextOffset; use frozen excel_read for stable captured analysis.
- `range.trace`：sheet,address; direction?: precedents|dependents. Native direct dependencies on active worksheet; Excel may reject cross-sheet or unavailable precedents.
- `range.write`：sheet,address,data: rectangular 2D values; formulas?: bool. Formula strings start =, English functions/commas. Values are literal. Supports tables; Excel protected/array/merged constraints still apply.
- `range.clear`：sheet,address; what?: contents|formats|notes|validation|conditionalFormats|all.
- `range.copy`：sheet,address,destination; destinationSheet?; pasteType?: -4104(all)|-4163(values)|-4123(formulas)|-4122(formats); pasteOperation?: -4142(none); skipBlanks?,transpose?.
- `range.insert`：sheet,address; axis?: rows|columns; shift?: right|down (insert), left|up (delete).
- `range.delete`：sheet,address; axis?: rows|columns; shift?: right|down (insert), left|up (delete).
- `range.sort`：sheet,address; header?: true; matchCase?,strokeOrder?; levels: [{column: 1-based relative column,mode: values|natural|custom,descending?: bool,textAsNumbers?: bool,rules?: string,sequences?: string}]. Uses the existing GridPilot multi-level sort engine; up to 20 levels. rules/sequences use the same syntax as the UI.
- `range.replace`：sheet,address,find,replacement; whole?,matchCase?.
- `range.transform`：sheet,address,transform: deduplicate|fillDown|fillRight|series|textToColumns; columns?: [1-based indexes],header?; rowcol?:1|2,type?:Excel XlDataSeriesType,date?,step?,stop?,trend?; delimiter?.
- `range.format`：sheet,address; properties?: {NumberFormat,HorizontalAlignment,VerticalAlignment,WrapText,ShrinkToFit,Orientation,IndentLevel,Locked,FormulaHidden,RowHeight,ColumnWidth}; font?: {Name,Size,Bold,Italic,Underline,Strikethrough,Color}; fill?: {Color,Pattern,PatternColor}; borders?: [{index:7..12,properties:{LineStyle,Weight,Color}}]; autoFit?. Excel COM colors are OLE BGR integers (red=255). Enum values are native Excel constants. Mixed formats supported.
- `range.merge`：sheet,address; merge?: true; across?.
- `range.validation`：sheet,address; delete?; type?: 1(whole)|2(decimal)|3(list)|4(date)|5(time)|6(length)|7(custom); operator?: Excel XlFormatConditionOperator; formula1?,formula2?,alertStyle?; properties?: {IgnoreBlank,InCellDropdown,ShowInput,ShowError,InputTitle,InputMessage,ErrorTitle,ErrorMessage}.
- `range.conditional_format`：sheet,address; delete?; kind: formula|cellValue|colorScale|dataBar|iconSet|top|duplicate; formula1?,formula2?,operator?; colors?:2|3; iconSet?: native ID; rank?,percent?,unique?; fillColor?,fontColor?,barColor?: #RRGGBB; criteria?: [{index,type,value?,color}].
- `range.goal_seek`：sheet,address: formula cell; changingCell,goal: number.

### csv

- `csv.import`：sheet,path: local UTF-8 CSV,address?:A1,delimiter?: comma. Imports as literal text to prevent formula execution. Up to 100k rows/2m cells.
- `csv.export`：sheet,address,path: new local CSV file.

### vba

- `vba.read`：module?: name; offset?:0-based line; limit?:1..500. Requires Excel Trust access to VBA object model enabled by the user; never changes security settings.
- `vba.write`：module: name,code: VBA source; type?:1(standard)|2(class). Replaces whole module; existing document modules supported. Requires macro-enabled workbook. Network/shell/dynamic execution requires MANUAL approval even in auto mode. Never runs the code.
- `vba.delete`：module: name. Document modules cannot be removed; use empty code to clear.

### table

- `table.list`：sheet; offset?.
- `table.create`：sheet,name; address?; header?; properties?: {Name,TableStyle,ShowHeaders,ShowTotals,ShowTableStyleFirstColumn,ShowTableStyleLastColumn,ShowTableStyleRowStripes,ShowTableStyleColumnStripes}; column?:1-based,formula?,totals?:XlTotalsCalculation,columnName?.
- `table.update`：sheet,name; address?; header?; properties?: {Name,TableStyle,ShowHeaders,ShowTotals,ShowTableStyleFirstColumn,ShowTableStyleLastColumn,ShowTableStyleRowStripes,ShowTableStyleColumnStripes}; column?:1-based,formula?,totals?:XlTotalsCalculation,columnName?.
- `table.delete`：sheet,name; keepData?:true.

### note

- `note.list`：sheet; offset?.
- `note.set`：sheet,address,text.
- `note.delete`：sheet,address.

### image

- `image.list`：sheet; offset?.
- `image.insert`：sheet,path: local image; name?,left?,top?,width?,height? (points).
- `image.delete`：sheet,name.

### name

- `name.list`：{}.
- `name.set`：name,refersTo: formula beginning =,visible?:true.
- `name.delete`：name.

### sparkline

- `sparkline.add`：sheet,address: destination,source: A1 source; type?:1(line)|2(column)|3(winloss).
- `sparkline.delete`：sheet,address.

### chart

- `chart.list`：sheet.
- `chart.create`：sheet,address; name?,title?,left?,top?,width?,height?; type: Excel 2010 XlChartType name or numeric enum. Types: xl3DArea xl3DAreaStacked xl3DAreaStacked100 xl3DBarClustered xl3DBarStacked xl3DBarStacked100 xl3DColumn xl3DColumnClustered xl3DColumnStacked xl3DColumnStacked100 xl3DLine xl3DPie xl3DPieExploded xlArea xlAreaStacked xlAreaStacked100 xlBarClustered xlBarOfPie xlBarStacked xlBarStacked100 xlBubble xlBubble3DEffect xlColumnClustered xlColumnStacked xlColumnStacked100 xlConeBarClustered xlConeBarStacked xlConeBarStacked100 xlConeCol xlConeColClustered xlConeColStacked xlConeColStacked100 xlCylinderBarClustered xlCylinderBarStacked xlCylinderBarStacked100 xlCylinderCol xlCylinderColClustered xlCylinderColStacked xlCylinderColStacked100 xlDoughnut xlDoughnutExploded xlLine xlLineMarkers xlLineMarkersStacked xlLineMarkersStacked100 xlLineStacked xlLineStacked100 xlPie xlPieExploded xlPieOfPie xlPyramidBarClustered xlPyramidBarStacked xlPyramidBarStacked100 xlPyramidCol xlPyramidColClustered xlPyramidColStacked xlPyramidColStacked100 xlRadar xlRadarFilled xlRadarMarkers xlStockHLC xlStockOHLC xlStockVHLC xlStockVOHLC xlSurface xlSurfaceTopView xlSurfaceTopViewWireframe xlSurfaceWireframe xlXYScatter xlXYScatterLines xlXYScatterLinesNoMarkers xlXYScatterSmooth xlXYScatterSmoothNoMarkers. Combo charts: set each series type/axisGroup separately. Stock/surface/bubble require correctly shaped source data.
- `chart.inspect`：sheet,chart; path?: element path; properties?: array of property names.
- `chart.edit`：sheet,chart,path?: element path; properties: Excel property/value object. Paths: ChartArea,PlotArea,ChartTitle,Legend,Axes(type,group),SeriesCollection(i),Points(i),DataLabels(i),Trendlines(i),ChartGroups(i),LegendEntries(i),Font,Format.Fill/Line.ForeColor,TickLabels,AxisTitle,MajorGridlines,MinorGridlines,DataLabel,Border,Interior,ErrorBars,DataTable,Walls,Floor. Read first. Set HasTitle/HasDataLabels/HasMajorGridlines before accessing absent objects. Set properties incl Text,Name,Color,RGB,Size,Bold,LineStyle,Weight,DashStyle,Transparency,MarkerStyle,MarkerSize,MinimumScale,MaximumScale,MajorUnit,MinorUnit,ScaleType,NumberFormat,Position,Left,Top,Width,Height,HasLegend,ChartType,AxisGroup,ShowValue,ShowCategoryName,ShowPercentage. COM colors are OLE BGR. Cannot navigate Parent/Application or execute arbitrary methods.
- `chart.delete`：sheet,chart.
- `chart.series`：sheet,chart,action:add|edit|delete; index?:1-based; name?,values?,xValues?,bubbleSizes?: arrays or local A1 range; type?,axisGroup?:1|2.
- `chart.trendline`：sheet,chart,series?:1; delete?,index?; type?: XlTrendlineType; properties?:{Order,Period,Forward,Backward,Intercept,DisplayEquation,DisplayRSquared}.
- `chart.errorbars`：sheet,chart,series?:1; delete?; direction?:-4168(Y)|-4167(X); include?:1; type?:2(fixed)|3(percent)|-4155(stdev)|4(stderr); amount?,minus?.

### pivot

- `pivot.list`：sheet.
- `pivot.create`：sheet,address: source with headers; destinationSheet: existing sheet,destination?:A3,name?; rows?,columns?,filters?: exact header name arrays; values?:[{field,function,caption?}]. Native local pivot cache. Create destination sheet first. Aggregations: sum/count/average/min/max/product/countNumbers/stdev/stdevp/var/varp.
- `pivot.inspect`：sheet,pivot; field?: field whose items to list; offset?:0-based.
- `pivot.edit`：sheet,pivot; layout?:0(compact)|1(tabular)|2(outline); repeatLabels?; source?:local A1; properties?: {Name,RowGrand,ColumnGrand,TableStyle2,ShowTableStyleRowHeaders,ShowTableStyleColumnHeaders,ShowTableStyleRowStripes,DisplayFieldCaptions,ShowDrillIndicators,InGridDropZones,DisplayNullString,NullString,DisplayErrorString,ErrorString,PreserveFormatting,HasAutoFormat,MergeLabels}.
- `pivot.field`：sheet,pivot,field; orientation?:0(remove)|1(row)|2(column)|3(filter)|4(data); position?; addValue?,caption?,function?; subtotals?:12 bools; properties?:{Caption,NumberFormat,Calculation,BaseField,BaseItem,ShowAllItems,LayoutForm,LayoutBlankLine,LayoutSubtotalLocation,RepeatLabels,IncludeNewItemsInFilter,EnableMultiplePageItems}. Calculation uses XlPivotFieldCalculation for show-as percent/running total/difference/rank. BaseField/BaseItem configure comparison. Inspect data field caption before editing it.
- `pivot.filter`：sheet,pivot,field; clear?; page?:string; visibleItems?:names; type?:XlPivotFilterType,dataField?,value1?,value2?. Supports manual items, report pages, label/value/date/top filters using native Excel enum. Cannot hide every item.
- `pivot.sort`：sheet,pivot,field; descending?,by?: field/caption.
- `pivot.group`：sheet,pivot,field; address?: specific item cells; ungroup?; start?,end?,by?; periods?:7 booleans(seconds,minutes,hours,days,months,quarters,years).
- `pivot.calculated`：sheet,pivot,name; field?:for calculated item, omit for calculated field; formula?:starts =; delete?.
- `pivot.refresh`：sheet,pivot. Refresh only local worksheet sources; no network refresh.
- `pivot.delete`：sheet,pivot. Refresh only local worksheet sources; no network refresh.

### slicer

- `slicer.add`：sheet,pivot,field; name?,caption?,top?,left?,width?,height?,timeline?. Excel 2010 slicers supported; timeline:true requires a date field and Excel 2013+.
- `slicer.delete`：sheet,cache; name?:specific slicer (omit to delete cache).

### snapshot

- `snapshot.list`：offset?:0-based; limit?:1..200 (default50). Local history metadata only. Restore is user UI only, opens an independent complete workbook preserving the current workbook.
- `snapshot.create`：label?:string.

### plugin

- `plugin.rename`：sheets?: selected sheet names; settings?: {Find,Replacement,Regex,MatchCase,FirstOnly,MatchesOnly,Prefix,Suffix,Trim,LetterCase,Numbering,Start,Step,Digits,RestartPerBook,NumberAtEnd,Separator,Truncate,IncludeHidden,IncludeCharts}. Uses existing GridPilot batch-rename planner, including regex and numbering; preview returns exact changes. Bound workbook only.
- `plugin.rename_preview`：sheets?: selected sheet names; settings?: {Find,Replacement,Regex,MatchCase,FirstOnly,MatchesOnly,Prefix,Suffix,Trim,LetterCase,Numbering,Start,Step,Digits,RestartPerBook,NumberAtEnd,Separator,Truncate,IncludeHidden,IncludeCharts}. Uses existing GridPilot batch-rename planner, including regex and numbering; preview returns exact changes. Bound workbook only.
- `plugin.sort_sheets`：settings?: {Natural,Alphabetical,Descending,MatchCase,Rules,Sequences}. Uses existing worksheet ordering engine.
- `plugin.convert`：sheet,address,to:text|number; settings?: existing ToolPreferences.
- `plugin.copy`：sheet,address,keepInternalFormulas?:true; settings?: existing CopyOptions. Creates a new workbook preserving appearance and optionally internal formulas.
- `plugin.segment`：sheet,address; settings?:{OutputMode:0(join)|1(columns)|2(rows),KeepPunctuation,Distinct,Separator,Terms}. Offline Chinese/English segmentation to a new workbook.
- `plugin.simple_mask`：sheet,address; path?:new local output workbook; settings?:{Mode,HeaderMode,HeaderRows,MaskHeaders,PreservePreamble,Prefix,Suffix,EdgeCount,Pattern,Replacement,Operator,Operand}. Existing simple masking engine; writes source if path omitted, otherwise exports.

## 对象与安全边界

图表路径支持 `Axes(2,1)`（也接受 `Axes(xlValue, xlPrimary)`）、`SeriesCollection(1).Points(2).Format.Fill.ForeColor`、ChartArea／PlotArea／ChartTitle／Legend／Font／DataLabels／Trendlines 等。属性清单位于 AiChartTools.cs；不同原生图表可编辑的属性由 Excel 决定，不是任意 COM 反射入口。不能沿 Application／Parent 跳出图表。

高级操作受保护区域、数组与合并规则、格式及版本限制约束。只刷新本地透视缓存，不更新网络连接。Excel 2013+ 时间线不在 Excel 2010 中可用。公式追踪返回原生直接依赖，不提供任意公式逐步执行器。工具错误可更正参数后重试，但修改错误可能已部分执行，必须先检查结果／恢复快照。

批量重命名与多级排序直接复用插件规则引擎。复杂数据库、监管核对及关系脱敏工作流仍由独立界面处理；不把已有界面自动算作 MCP 支持。

静态 VBA 风险检查聚焦进程启动、网络 API、动态自动化目标。不是沙箱，不执行代码，不修改信任中心。规则和测试见 VbaRisk.cs / AiExtendedTests.cs。

## 快照实现

不可变 JSON 清单 + SHA-256 内容寻址 + 256 KiB gzip 块。OOXML 先拆 ZIP 部件再去重，.xls 按原始文件分块。原子写入 flush+rename；恢复验证每块哈希。所有原始字节保留；仅审批比较时规范化保存时间、工作簿 revision 指针及 VBA 工程随机元数据。CFB 比较逻辑规范化目录／流，保留 VBA 保护字段的解码值，不清除保护。

规范参考：[Microsoft MS-OLEPS 保存时间](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/47f4bebf-9513-4d40-a4bc-1eda5ce6732a)、[Microsoft MS-OVBA 数据解码](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/7e9d84fe-86e3-46d6-aaff-8388e72c0168)、[Microsoft 图表枚举](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype)、[excel-mcp-server](https://github.com/haris-musa/excel-mcp-server)。

恢复打开完整独立工作簿，保留原文档的未保存编辑。高级写入的失败信息包含变更前版本编号；历史预算自动回收已实现，SaveCopyAs 仍有完整保存成本；不具备数据库 WAL 或瞬时 COW。

## v2.1 新版 Office 操作

`capability.list` 返回 Version/Build、接口探测和可用操作；`excel_help({snapshotId,topic})` 按宿主能力过滤。无需每次重复读取能力清单；同一工作簿和宿主会话可以复用结果，遇到原生错误再重新核验。

| 组 | 操作 | 启用依据 |
| --- | --- | --- |
| 动态数组 | range.formula2 / range.spill | 16+ 且 Formula2、HasSpill 实际可读 |
| 线程批注 | comment.list / add / reply / delete | 16+ 且 CommentThreaded 实际可读 |
| Power Query | query.list / set / delete | 16+ 且 Workbook.Queries 可读；不刷新来源 |
| 数据模型 | model.inspect | 15+ 且 Workbook.Model 可读 |
| 连接 | connection.list | 名称、类型与说明；不返回连接串，不刷新 |
| 切片器 | slicer.list / filter / clear / timeline | 基础筛选；时间线 15+ 并由原生缓存校验 |
| 表格及计算 | table.resize / range.calculate | 原生 Excel 表格/区域接口 |
| 现代图表 | chart.create / edit / series | 组合类型 15+、现代类型 16+；具体数据形状/Build 原生校验 |

现代图表使用 Shapes.AddChart2，创建前临时选定源范围，并恢复原活动位置；不调用部分现代类型未实现的 SetSourceData。旧图表继续原有路径，趋势线/误差线的 type 与图表 type 分开验证。新工具仍走完整快照和批准链路。

官方核验：[Formula2](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2)、[HasSpill](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hasspill)、[Queries](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.queries)、[CommentThreaded](https://learn.microsoft.com/en-us/office/vba/api/excel.range.commentthreaded)、[AddChart2](https://learn.microsoft.com/en-us/office/vba/api/excel.shapes.addchart2)、[XlChartType](https://learn.microsoft.com/en-us/office/vba/api/excel.xlcharttype)。

自动/所有工作簿模式的最大读取范围为目录创建时当前 Excel 进程已打开的工作簿及 VBA。workbook.open/import、目录文件枚举、CSV 导入和磁盘图片读取在这些模式中被拒绝；新打开的文件需重新发送以创建目录。恢复、批准与宏执行始终不属于模型工具。
