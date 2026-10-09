SheetPilot 0.1.0 — 便携式 Excel 工作表导航
========================================

交付目标：Windows 10 + Excel 2010 标准版，提供 x86 / x64 两种单文件 XLL。
这是第一版可运行原型；目标机验收状态请看 VALIDATION.txt。

快速使用
--------
1. 将整个压缩包解压到普通文件夹，不要在压缩包内直接运行。
2. 在 Excel 2010 的「文件 → 帮助 → 关于 Microsoft Excel」确认 Excel 位数。
   32 位 Excel 选择 SheetPilot-x86.xll；64 位 Excel 选择 SheetPilot-x64.xll。
   按 Excel 位数选择，不能按 Windows 位数选择；一次只加载其中一个。
3. 双击对应 XLL，或在 Excel 中用「文件 → 打开」选择它，按 Excel 提示启用加载项。
4. 点击 Ribbon「SheetPilot → 工作表导航」。首次点击才创建左侧面板。
5. 选择「当前工作簿」或「所有已打开工作簿」，输入工作表名称，单击结果跳转。

这是按次加载，无安装程序、无需管理员权限、不修改工作簿为宏格式。
如果希望每次启动 Excel 自动加载，可自行通过「文件 → 选项 → 加载项 → 管理：Excel 加载项
→ 转到 → 浏览」添加 XLL；这会让 Excel 记住文件路径，并非 Windows 安装。
采用自动加载后，移动文件夹前先取消勾选原路径，移动后重新添加。

功能
----
• 列出工作表、图表页、隐藏页、VeryHidden 页，保留 Excel 原始顺序。
• 普通包含搜索；.* 正则开关；Aa 大小写开关；含隐藏页开关。
• 当前页面有绿色圆点；同名 Sheet 通过副标题中的工作簿名区分。
• 单击结果或按 Enter 跳转，↑ ↓ 选择；面板内 Ctrl+F 聚焦搜索，Esc 清空，F5 刷新。
• 默认约 1.2 秒同步目录；编辑单元格或 Excel 弹出对话框时，等待 Excel 可接受操作。
• 关闭面板可再次从 Ribbon 打开。拖动 Office 面板标题可以调整停靠位置。
• 查询、大小写和范围设置只在本次面板会话中保存。

范围说明
--------
「所有已打开工作簿」指同一个 Excel 进程内的普通工作簿，不跨独立 Excel 进程；
不读取受保护视图中的未启用文档，也不列出 IsAddin=true 的工作簿。
范围搜索作用于 Sheet 名称，工作簿名作为结果说明显示。
隐藏页可搜索，但不会自动取消隐藏；本版不修改任何单元格、公式、格式或工作簿可见性。
Excel 2010 使用一个应用级面板。较新 Excel 按窗口创建面板，可在对应窗口再次打开导航。

正则示例（先开启 .* 正则）
---------------------------
^销售             以“销售”开头
汇总$             以“汇总”结尾
华东|华南         匹配任一词
^Sheet\d+$        Sheet 后跟数字
^(?!.*归档).*$    名称中不含“归档”
表达式最多 512 字符；单条匹配约 40ms 超时，整次搜索约 200ms 预算。
超时或语法错误会显示提示，不会执行表达式中的代码。

便携性及运行条件
----------------
• XLL 内已打包业务、界面和 Excel-DNA 运行组件，运行不需要旁置 DLL 或网络。
• 不使用 VBA、UserForm WebBrowser、WebView2、Electron、VSTO 或 Office.js。
• 不需要在 Windows 中安装本加载项，也不需要运行 regsvr32。
• 需要系统已有 .NET Framework 4.6.2 或更高版本。Win10 1607 及之后正常系统自带满足
  要求的版本；未经升级的 Win10 1507/1511 不在本次免安装运行基线内。
• Check-Environment.ps1 仅检查环境、识别 Excel.exe 位数，不安装组件或更改注册表。
• Excel-DNA 创建 Ribbon/任务面板时会做当前用户的临时 COM 注册并清理，因此
  “便携”不代表完全不访问注册表。企业策略禁止 HKCU 临时 COM 注册时，面板可能无法加载。
• 日志写入 %TEMP%\SheetPilot-<进程号>.log，最多约 1MB，记录操作和错误类型，
  不记录搜索词、单元格内容或工作簿完整路径。XLL 所在文件夹可以只读。
• 本版 XLL 未进行发布者数字签名。若从互联网下载被标记，核对来源和校验值后可在
  文件属性中解除该文件的阻止；不要为此关闭 Excel 的全部安全设置。

移除与升级
----------
按次加载：退出 Excel 后即可删除或替换 XLL。
自动加载：先在 Excel 加载项列表取消勾选，再退出 Excel 后删除或替换。
不要在 Excel 尚未退出时覆盖已加载的文件；不要同时加载不同目录中的两个副本。

开发与扩展
----------
source/ 内含完整源码、模块接口、构建脚本、测试和架构文档。
开发机需要 .NET SDK 8/9 或对应 Visual Studio；终端用户不需要 SDK。
详见 source/ARCHITECTURE.txt 和 source/BUILD.txt。

依据
----
Excel-DNA Ribbon/任务面板接口：https://excel-dna.net/reference-various/
Excel-DNA 1.9 运行基线：https://github.com/Excel-DNA/ExcelDna/releases
.NET 与 Windows 版本对应：https://learn.microsoft.com/en-us/dotnet/framework/install/versions-and-dependencies
