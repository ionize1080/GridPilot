# 快速开始

[文档首页](../README.md) · 适用 v1.9.1 · 2026-10-21

## 下载与加载

1. 从 [最新正式版](https://github.com/ionize1080/GridPilot/releases/latest) 下载便携 ZIP 并完整解压。需要校验时，对照同一 Release 的 SHA-256 文件。
2. 选择自己有写入权限的目录。不要直接在压缩包内运行，也不要只移动 XLL。
3. 在 Excel 的「文件 → 帮助／账户 → 关于 Excel」查看位数；菜单名称随 Office 版本略有不同。
4. 在「Excel 选项 → 加载项 → 管理 Excel 加载项 → 转到 → 浏览」中加载相应 XLL。32 位 Excel 选 `GridPilot-x86.xll`，64 位选 `GridPilot-x64.xll`。
5. 在功能区找到 **GridPilot**。未显示的工具可从「设置 → 功能开关」启用。

```text
GridPilot/
├─ GridPilot-x86.xll     ← 32 位 Excel 使用
├─ GridPilot-x64.xll     ← 64 位 Excel 使用
├─ *.dll                ← 业务组件
├─ App/                 ← 本地工具界面
├─ Browser/             ← 随包提供的运行时
├─ Chatbox/             ← AI 助手与 MCP 集成
├─ Resources/           ← 离线词典
└─ docs/                ← 使用说明与许可
```

需要 64 位 Windows 10/11 和 .NET Framework 4.6.2+。Excel 位数决定 XLL，Windows 位数决定便携运行时能否运行。无需另装 Node.js 或 WebView2；公司设备的 Office 信任与执行策略仍然适用。已测平台见 [兼容性](COMPATIBILITY.md)。

## 第一次使用 AI

1. 打开「AI 助手」，在 Chatbox 设置中配置你的服务地址、密钥和支持工具调用的模型。
2. 在 Excel 准备合成数据或可分享的数据；选择「选区」或「工作表」。
3. 试着输入：「请使用 Excel 工具，在 E2 写入 B2:B4 的求和公式。」
4. 查看待应用建议，确认目标后点「一键应用」。聊天里显示公式文字本身不会写入单元格。
5. 查看结果，再用「变更前／变更后」验证恢复。熟悉后再按需开启工作簿自动批准。

<p><img src="../v1.9.1/chatbox-settings.png" width="780" alt="v1.9.1 原生 Chatbox 设置入口"></p>

AI 使用你配置的服务，可能产生 API 费用；本地排序、脱敏等工具不需要模型。完整说明见 [AI 助手](AI.md)。

## 先试三个本地工具

| 工具 | 建议的首次操作 |
| --- | --- |
| 工作表导航 | 搜索表名，单击跳转，F2 改名，Esc 取消 |
| 多级排序 | 打开内置示例工作簿，选择部门 → 客户 → 项目示例，预览后应用 |
| 简易脱敏 | 选择少量文字，检查表头识别，先预览并导出新文件 |

## 升级与回退

关闭所有 Excel 和 GridPilot 窗口后，备份原目录，完整解压新版。保留 `GridPilot.navigation.json`、`GridPilot.tools.json` 及其 `.bak`，不要把旧程序文件覆盖回新版。确保 Excel 只加载一个版本。

AI 聊天配置、Recovery 副本和报表数据库位于用户目录，单独备份方法见 [数据与存储](DATA-AND-STORAGE.md)。新版本不能正常启动时，可关闭 Excel 并恢复旧目录；不要仅替换某个 DLL 或混用版本。

[下一步：AI 助手](AI.md) · [常见问题](TROUBLESHOOTING.md)
