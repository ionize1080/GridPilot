# SheetPilot

面向 Windows 桌面 Excel 的便携式工作表导航加载项。原生 Ribbon 入口搭配 WPF 侧边栏，像翻阅 PDF 书签一样查找并切换 Sheet。

**当前版本：v0.1.0 预发布原型。** 目标为 Windows 10 + Excel 2010 标准版；已在 Excel 16.0 64 位完成集成测试，Excel 2010 和 x86 加载仍需目标机验收。

## 下载与使用

从 [Releases](../../releases) 下载 `SheetPilot-0.1.0-portable.zip` 并解压。

1. 在 Excel 的“关于”页面确认 **Excel 位数**，不要按 Windows 位数选择。
2. 32 位 Excel 加载 `SheetPilot-x86.xll`；64 位 Excel 加载 `SheetPilot-x64.xll`。一次只加载一个。
3. 双击 XLL，或在 Excel 中通过“文件 → 打开”加载。
4. 点击“SheetPilot → 工作表导航”。

无需安装本插件、VSTO、浏览器组件或管理员权限。系统需已有 .NET Framework 4.6.2+；Windows 10 1607 及之后的正常系统满足这一运行基线。`Check-Environment.ps1` 提供只读检查。

Excel-DNA 创建 Ribbon 和任务面板时会做当前用户的临时 COM 注册，所以便携不等同于完全不访问注册表。XLL 未进行发布者数字签名；加载仍受 Office 信任策略约束。

## 功能

- 范围切换：当前工作簿 / 同一 Excel 进程内所有已打开工作簿。
- 枚举工作表、图表页、隐藏页和 VeryHidden 页，保留原始顺序。
- Sheet 名称包含搜索、正则表达式、大小写与隐藏页选项。
- 单击或 Enter 跳转，方向键选择；面板内 Ctrl+F 搜索、Esc 清空、F5 刷新。
- 自动同步页面名称、顺序和活动状态；正则搜索具有取消和超时保护。
- 隐藏页仅显示和标注，不擅自取消隐藏或修改工作簿。

这是 Sheet 页面目录；自定义书签树、单元格书签和跨会话收藏尚未实现。

<img src="docs/navigation-preview.png" width="340" alt="实际 WPF 控件渲染的工作表导航预览，使用示例数据" />

图像来自实际 WPF 控件的离屏渲染，并非 Excel 2010 屏幕截图。

## 架构与扩展

| 项目 | 职责 |
| --- | --- |
| `SheetPilot.Core` | 不依赖 Office 的数据模型、搜索、模块与命令合同 |
| `SheetPilot.UI` | WPF 导航视图和视图模型 |
| `SheetPilot.Excel` | Excel-DNA、Ribbon、COM 适配和面板生命周期 |
| `SheetPilot.Tests` | 基础测试、WPF 渲染和可选 Excel 集成测试 |

新功能实现 `IFeatureModule`，提供命令并在组合根注册，Ribbon 自动生成分组和入口。UI 不持有 Office COM 对象；后台任务只处理脱离 Excel 的快照。

采用 Excel-DNA 1.9.0 + .NET Framework 4.6.2 + WPF/ElementHost。不使用 VBA、WebView、WebView2、Electron 或窗口父子关系修改。

详细决策、边界与后续路线见 [ARCHITECTURE.txt](ARCHITECTURE.txt)。

## 构建与验证

开发机需要 Windows 和 .NET SDK 8/9。终端用户不需要 SDK。

```powershell
powershell -NoProfile -File .\build.ps1
```

脚本使用依赖锁文件，构建双位数 XLL、运行基础测试并输出 `dist/`。构建与基础测试不要求安装 Excel。

已完成 45 项基础及 Excel 16 集成断言、35 项包含五张控件渲染的基础断言；两组有重复检查，不能相加作为独立测试数。实际集成测试覆盖从单 XLL 隔离目录加载、创建任务面板和正常退出测试 Excel 实例。

**Windows 10 + Excel 2010 标准版的任务栏文件切换、退出、停靠和 x86 加载尚未验收。** 未据新版 Excel 的测试结果保证旧版兼容性。

- [详细构建说明](BUILD.txt)
- [测试记录与目标机验收清单](VALIDATION.txt)
- [使用说明](README.txt)
- [第三方组件说明](THIRD-PARTY-NOTICES.txt)

## 发布内容

Release 提供完整便携包、独立 x86/x64 XLL 与 SHA-256 校验文件。源码由 Git 标签固定。升级时先退出 Excel，再替换文件。
