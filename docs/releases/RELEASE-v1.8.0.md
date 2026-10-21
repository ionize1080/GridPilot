# GridPilot v1.8.0 — 更轻巧的 Excel AI 工作区

固定集成 Chatbox 1.20.3，模型请求与请求头继续沿用原生链路。

- 修复侧栏键盘焦点：改用跟随 Excel 面板的原生窗口承载。
- Excel MCP 改为本地 stdio，移除 HTTP 监听端口；引入官方开源 MCP SDK。
- 工作范围移入原生输入工具栏，待应用公式和内容支持一键应用、整批撤销与冲突检查。
- 移除20万格限制，支持完整工作表边界，非空区域分块扫描和磁盘快照。
- 修复新版 Electron 下 PDF、DOCX、TXT 等附件路径与原生解析兼容。
- 重绘 AI 助手图标；功能开关支持开启/关闭 AI。
- 侧栏与独立窗口即时切换，保留会话和草稿，无需重启 Excel。
- 9 个 Excel MCP 工具，新增批量内容/公式、范围格式、原生图表、工作簿结构检查；移植 MIT Excel MCP 范围/格式校验。

请完整解压便携包，按 Excel 位数加载 x86/x64 XLL。升级程序文件时退出正在加载旧版的 Excel，保留 GridPilot.tools.json、GridPilot.navigation.json 和原 Chatbox 用户目录。

验证环境：Windows 11 + Excel 16 x64。Win10 / Excel 2010 / 32位 Excel 未实机验收。普通写入不包含表格计算列、合并或数组区域；范围格式要求原格式一致。扫描型 PDF 的 OCR 沿用 Chatbox 配置。完整范围处理仍受机器内存、磁盘与处理时间约束。没有引入任意 VBA 执行。

详见仓库 `docs/v1.8/VALIDATION.md` 与 `docs/v1.8/DESIGN.md`。
