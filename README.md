# GridPilot

Windows 桌面 Excel 的便携工具集。当前版本 **1.4.0**。

[下载完整便携包](https://github.com/ionize1080/GridPilot/releases/tag/v1.4.0) · [更新说明](RELEASE-v1.4.0.md) · [验证记录](VALIDATION.txt) · [数据库设计](docs/v1.1/DESIGN.md)

新版脱敏与总设置采用独立 Chromium 窗口，免安装、离线运行、按需启动。导航保留轻量原生任务面板。XLL、业务 DLL、词典和浏览器已拆分，**必须完整解压，不能单独复制 XLL**。

![数据脱敏界面，合成演示数据](docs/v1.0/mask.png)

## 使用

1. 便携包要求 64 位 Windows 10/11 和已有 .NET Framework 4.6.2+；无需 WebView2、Chrome、Node.js 或管理员安装。
2. 按 **Excel 位数**加载 `GridPilot-x86.xll` 或 `GridPilot-x64.xll`。32 位 Excel 可以配合独立 64 位 Chromium 界面运行。
3. 升级时退出 Excel，保留 `GridPilot.navigation.json` 和 `GridPilot.tools.json`。新旧版本不要同时加载。

本机验证环境是 Windows 11 + Excel 16 x64。**Win10 21H1 + Excel 2010、32 位 Excel 尚未实机验收**。构建基线与兼容路径保留，但不能把新版 Excel 的测试结果视为旧版验收。

## v1.4 监管核对与变动报告

Ribbon → **监管核对**：连接两期余额、发放回收、专项表及监管报表，保存取数规则后每期重复运行。支持财务/试算平衡多项加减汇总、错误跳过、Top N 变动、来源行定位与报告草稿。

两期金额、差额与已配置的现金流分别核对；异常记录保留为“部分结果”。自动识别底稿只生成待复核映射，不推断未知监管口径。导出 Markdown / CSV / 完整 JSON 底稿，源文件只读。

[使用说明](RECONCILIATION.txt) · [功能与技术设计](docs/v1.4/DESIGN.md)

![监管核对工作区，合成演示数据](docs/v1.4/reconciliation.png)

## v1.3 智能表头与简易脱敏

Ribbon → **简易脱敏**：处理选区、整行、整列和不连续区域。支持固定拼音首字母/字符替换、首尾交换、倒序、自定义正则及纯数字加减乘除，可选前后缀。逐格预览和跳过错误；可复制、导出 XLSX，或确认备份后写回原表并撤销。

默认智能识别并保留表头，可选择表头也脱敏或手动指定表头行。区分合并标题、时间与导出者等前置信息，支持多层表头；预览可核对识别字段。复杂布局请手动修正。

首字母不能单独还原中文。需要将来恢复时，请单独导出含原文的还原记录 JSON；固定映射不会随运行随机改变。详见 [使用与边界](SIMPLE-MASKING.txt)。

![简易脱敏界面，合成演示数据](docs/v1.3/simple-mask.png)

## 功能

- 多报表数据库：时点快照、累计业务组合键去重、空号保留、冲突核对、来源定位；只读 SQL 和可视化查询，源 Excel 保存后扫描更新。
- 脱敏/分词预览最多 20,000 条，分页与搜索；脱敏结果直接复制到系统剪贴板。
- 导航右键管理、F2 重命名、拖放排序；图标字形居中。
- 清洁版 ZIP 只含运行文件，许可合并到“设置 → 关于与许可”。

- 词频前后缀建议、手动格式与跨表同步；完整大表分块处理、实际有内容行列统计与实时进度。
- 智能脱敏：标题与多行合并表头识别、字段建议、跨表一致映射、本地预览、实际 Excel 重算并重开验证。
- 可选保留基础单元格样式、行高列宽、隐藏行列和合并。另存独立 XLSX，源工作簿保持原样。
- 青瓷、晴蓝、石墨图标配色；透明聚焦与可调四向涟漪，不修改工作簿格式。
- 工作表导航、中英文离线分词、原样复制、批量重命名、文本/数值转换及撤销。

![总设置](docs/v1.0/settings.png)

![报表数据库，合成演示数据](docs/v1.1/database.png)

## 架构与构建

`SheetPilot.Core` 负责纯逻辑，`SheetPilot.Excel` 负责 Excel COM 与主线程调度，`SheetPilot.UI` 保留轻量 WPF 控件；`App/` 提供浏览器界面，`Browser/` 是经 SHA-256 校验的 Electron 44.7.0 便携运行时。

浏览器采用沙盒与上下文隔离，禁止远程导航、网络请求与外部窗口。每个窗口通过只允许当前用户访问的随机命名管道通信，源 Excel COM 只在宿主主线程运行，私有导出 Excel 使用专用 STA 线程。

开发机安装 .NET SDK 后运行 `prepare-browser.ps1`，再运行 `build.ps1`。终端用户不需要 SDK 或安装器。构建使用锁定的 NuGet 依赖。

详见 [架构](ARCHITECTURE.txt)、[脱敏边界](MASKING.txt)、[原样复制](COPYING.txt)、[第三方许可](THIRD-PARTY-NOTICES.txt)。
