# GridPilot

Windows 桌面 Excel 的便携工具集。当前版本 **0.9.0**。

[下载完整便携包](https://github.com/ionize1080/GridPilot/releases/tag/v0.9.0) · [更新说明](RELEASE-v0.9.0.md) · [验证记录](VALIDATION.txt)

新版脱敏与总设置采用独立 Chromium 窗口，免安装、离线运行、按需启动。导航保留轻量原生任务面板。XLL、业务 DLL、词典和浏览器已拆分，**必须完整解压，不能单独复制 XLL**。

![数据脱敏界面，合成演示数据](docs/v0.9/masking.png)

## 使用

1. 便携包要求 64 位 Windows 10/11 和已有 .NET Framework 4.6.2+；无需 WebView2、Chrome、Node.js 或管理员安装。
2. 按 **Excel 位数**加载 `GridPilot-x86.xll` 或 `GridPilot-x64.xll`。32 位 Excel 可以配合独立 64 位 Chromium 界面运行。
3. 升级时退出 Excel，保留 `GridPilot.navigation.json` 和 `GridPilot.tools.json`。新旧版本不要同时加载。

本机验证环境是 Windows 11 + Excel 16 x64。**Win10 21H1 + Excel 2010、32 位 Excel 尚未实机验收**。构建基线与兼容路径保留，但不能把新版 Excel 的测试结果视为旧版验收。

## 功能

- 智能脱敏：标题与多行合并表头识别、字段建议、跨表一致映射、本地预览、实际 Excel 重算并重开验证。
- 可选保留基础单元格样式、行高列宽、隐藏行列和合并。另存独立 XLSX，源工作簿保持原样。
- 青瓷、晴蓝、石墨图标配色；透明聚焦与可调四向涟漪，不修改工作簿格式。
- 工作表导航、中英文离线分词、原样复制、批量重命名、文本/数值转换及撤销。

![总设置](docs/v0.9/settings.png)

## 架构与构建

`SheetPilot.Core` 负责纯逻辑，`SheetPilot.Excel` 负责 Excel COM 与主线程调度，`SheetPilot.UI` 保留轻量 WPF 控件；`App/` 提供浏览器界面，`Browser/` 是经 SHA-256 校验的 Electron 44.7.0 便携运行时。

浏览器采用沙盒与上下文隔离，禁止远程导航、网络请求与外部窗口。每个窗口通过只允许当前用户访问的随机命名管道通信，COM 只在 Excel 主线程运行。

开发机安装 .NET SDK 后运行 `prepare-browser.ps1`，再运行 `build.ps1`。终端用户不需要 SDK 或安装器。构建使用锁定的 NuGet 依赖。

详见 [架构](ARCHITECTURE.txt)、[脱敏边界](MASKING.txt)、[原样复制](COPYING.txt)、[第三方许可](THIRD-PARTY-NOTICES.txt)。
