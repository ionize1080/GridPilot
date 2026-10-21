# GridPilot v1.6.1

> 本页记录该次发布，保留原始功能与验证结论。[最新使用文档](../README.md) · [全部版本](README.md)

修复导航右键菜单与工作表原地重命名的三个交互问题。

- 导航条目右键改用 Windows 原生菜单，避免冷启动 Office 原生窗格中 WPF 弹出菜单的焦点和鼠标捕获依赖，无需先打开再关闭高级排序。菜单期间暂停目录刷新和自动焦点恢复。
- 双击或 F2 改名后，点击列表空白、窗格背景、同一条目的非编辑区域或其他条目，均会提交有效名称并退出编辑。Enter 提交，Esc 取消；非法名称保留编辑并提示。
- 编辑框右键不会提前结束修改，支持撤销、剪切、复制、粘贴、删除及全选。可将复制来的文字粘贴为新 sheet 名；菜单关闭后继续编辑，确认后才写入工作表。Ctrl+V 仍可使用。

## 验证

2026-10-17，本机 Windows 11 + Excel 16 x64：

- 完整离线依赖构建通过，.NET Framework 4.6.2，生成 x86/x64 XLL。
- 新增独立冷启动原生窗格回归，在 x64 与 x86 进程分别验证菜单持续显示、原生菜单循环、刷新保护、点击空白/背景/同一行提交、编辑右键保留选区、粘贴、Enter/Esc、非法名称、失焦与卸载处理。测试不创建 WPF Application/Window，也不预先启动 Electron。
- 既有基础 265、颜色选择器 90、功能 52、升级 27、v1 169、v1.1 97、多级排序纯逻辑 5 项检查通过；监管核对脚本通过。
- 使用真实打包 XLL 和新建合成工作簿手动式界面验收：首次右键菜单持续显示；右键菜单进入改名；双击/F2 编辑；点击窗格空白后 Excel 表签更新；编辑框右键粘贴中文名称并保持编辑，Enter 后表签更新；Esc 放弃草稿；高级排序窗口打开/关闭后再次右键正常。单元格内容未被名称输入改写。

本次针对 Win10 + Office 2010 的报告修正宿主菜单路径，但本机没有该组合，**未完成 Win10 + Office 2010 或 32 位 Excel 实机验收**。x86 进程回归不等同于 x86 Excel 验收。

原生菜单使用系统 [TrackPopupMenuEx](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-trackpopupmenuex)，由 Windows 管理菜单循环及关闭，不需要 Electron 预热。

## 下载与升级

- `GridPilot-1.6.1-portable-win-x64.zip`：完整便携包，含文档与许可。
- `GridPilot-1.6.1-clean-win-x64.zip`：仅运行文件，许可可从“设置 → 关于与许可”查看。
- 两个包均含适配 32/64 位 Excel 的 XLL，以及独立 64 位 Electron 运行时。要求 64 位 Windows 10/11 和 .NET Framework 4.6.2+。

先退出 Excel，再完整解压并替换运行文件，保留 `GridPilot.navigation.json`、`GridPilot.tools.json`。按 Excel 位数加载对应 XLL，不能只更新 XLL。发布附件附 SHA-256 校验清单。
