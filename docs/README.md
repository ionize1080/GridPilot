# GridPilot 文档

**适用版本：v2.1.0 · 更新：2026-10-22**

先按你要完成的事情选择入口。当前指南描述最新行为；版本目录保存当时的设计和测试事实。

## 开始使用

| 入口 | 内容 |
| --- | --- |
| [快速开始](guides/QUICKSTART.md) | 下载、选择 XLL、加载、第一次使用与升级 |
| [AI 助手](guides/AI.md) | 范围选择、13 个 MCP 入口与按需操作目录、自动批准、批量建议、快照与窗口切换 |
| [设置与偏好](guides/SETTINGS.md) | 功能开关、外观、导航、聚焦与配置保存 |
| [兼容性](guides/COMPATIBILITY.md) | 运行条件、已测平台与目标机验收 |
| [故障排查](guides/TROUBLESHOOTING.md) | 加载、卡顿、建议应用、恢复与窗口问题 |
| [数据与存储](guides/DATA-AND-STORAGE.md) | 哪些数据联网、文件保存在哪里、升级与清理 |

## 功能指南

| 工作方式 | 指南 |
| --- | --- |
| 整理工作簿 | [导航、聚焦、改名、文本与数值](guides/WORKSHEET-TOOLS.md) · [多级排序](guides/SORTING.md) |
| 输出与清理 | [原样复制](guides/COPYING.md) · [简易脱敏](guides/SIMPLE-MASKING.md) · [智能脱敏](guides/MASKING.md) · [中英文分词](guides/SEGMENTATION.md) |
| 报表分析 | [报表数据库](guides/DATABASE.md) · [监管与财务核对](guides/RECONCILIATION.md) |
| AI 分析与操作 | [使用流程与恢复](guides/AI.md) · [MCP 参数速查](development/MCP.md) |

## 开发与验证

- [架构](development/ARCHITECTURE.md)：Excel 主线程、原生窗格、Electron、MCP、数据流和恢复机制。
- [构建与测试](development/BUILD.md)：固定依赖、构建命令、专项测试、真实 API 测试与发布检查。
- [界面规范](development/UI.md)：窗口、图标、菜单、进度和应用反馈。
- [验证总览](validation/README.md)：当前证据、测试层次与尚未完成的目标机验收。
- [版本记录](releases/README.md)：所有正式版本说明、历史设计和验证入口。
- [第三方说明](licenses/THIRD-PARTY-NOTICES.txt) · [Chatbox 来源与补丁](licenses/CHATBOX-NOTICES.txt)：许可原文保留在 `licenses/`。

## 文档约定

当前指南使用 Markdown，方便在 GitHub 阅读图表、表格和截图。旧的 `.txt` 路径保留为新文档的入口，便携包中也可直接找到同名 Markdown。

历史页面顶部标明适用版本；其中的数字、限制、问题与结论只代表记录当时，不累加为当前版本的测试结果。截图使用合成数据；较早版本的功能示意会注明拍摄版本。许可证原文与测试样本不作改写。

[返回项目首页](../README.md)
