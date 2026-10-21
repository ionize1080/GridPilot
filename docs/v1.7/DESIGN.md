# GridPilot 1.7：Chatbox 1.20.3 与 Excel

> 此页保留该版本的设计／验证记录；后续行为以 [当前指南](../README.md) 与 [当前架构](../development/ARCHITECTURE.md) 为准。

调研日期：2026-10-20。以官方文档确认交互方式，不把公开文档视为各订阅、地区和 Excel 版本均已开放的保证。

## 市场交互与本版选择

| 产品 | 官方说明中的交互 | GridPilot 对应设计 |
| --- | --- | --- |
| Microsoft Copilot in Excel | 对话窗格、公式解释、插入公式列、分析与应用操作；当前文档区分编辑、计划和聊天模式 | 默认右侧对话；解释公式、生成公式、查找、统计快捷提示；公式先展示再应用 |
| Claude for Excel | 工作簿问答、多工作表分析、公式解释、单元格级可点击引用 | 可选 workbook 上下文；搜索结果有单元格定位；回答携带 sheet/address 证据 |
| Chatbox 1.20.3 | 原生多模型会话、设置、MCP 与数据管理 | 使用官方二进制界面，保留全部原生设置入口，Excel 能力接入原生 MCP 工具链 |

来源：[Microsoft 使用指南](https://support.microsoft.com/en-us/excel/copilot/get-started-with-copilot-in-excel)、[公式解释与插入流程](https://support.microsoft.com/en-us/accessibility/copilot/use-copilot-in-excel-with-a-screen-reader)、[Claude for Excel](https://claude.com/docs/office-agents/excel)、[Chatbox v1.20.3](https://github.com/chatboxai/chatbox/releases/tag/v1.20.3)。

这是基于上述交互的设计取舍，不宣称实现了这些产品全部的编辑、建模、联网和数据透视表能力。

## 上下文合同

- 默认选区。整行/列与 UsedRange 相交；不连续区域逐块读取。完全空白的选区只读取其首格，避免扫描整张空表。
- Sheet 读取当前普通工作表的 UsedRange；Workbook 读取当前工作簿的可见 Worksheets。图表页、隐藏及 VeryHidden 工作表不进入上下文。已选范围内的隐藏行列包括在内。
- 输入框上方显示当前工作簿、工作表和选区。每次发送重新抓取快照。历史消息保留简短范围摘要，原始快照仅在当前窗口内存中保留；只在调用模型时展开最新用户消息的快照。重开会话或快照过期后需要重新发送。选择“不附加数据”后，新消息不读取 Excel；原有会话历史仍由 Chatbox 管理。
- 最多扫描200,000格；超限整体失败并提示缩小范围，绝不把抽样当完整扫描。提示中给出120个非空单元格样本、完整数值汇总、实际扫描数和采样标记。按需读取/查找每页200格。
- MCP 工具明确传入消息中的随机 snapshotId。读取、查找、统计均在该快照上完成，不能因当前选区或工作簿变化而改变目标。每窗口最多保留8份快照，过期需重新发送。
- 查找支持包含、精确文本、数值大于/小于和公式文本；数值统计支持按表、按列。统计包含日期序列值，不把数值文本自动转换成数字，不把局部样本推断为整表。
- 原工作簿关闭后快照失效。同名新工作簿不会替代原目标。只在用户发送时将所选数据提供给配置的模型，状态轮询只读取名称和地址。

## 公式与编辑

模型只能提出单格公式建议，原生 MCP 中没有 apply、undo 或任意代码执行的 Excel 工具。建议卡片显示工作簿、工作表、地址、原内容和公式，用户点击“应用公式”后才执行。目标表须属于快照；目标单元格可以在原选区之外，位置始终明确展示。

应用前检查原工作簿是否仍打开、目标内容是否改变、只读/保护/隐藏状态、合并和数组公式。表格计算列拒绝写入，避免 Excel 自动填充其他行。普通 Formula 接口采用隐式交叉；不写 Formula2 动态溢出数组。外部链接、结构化引用、数组常量及执行/网络函数被拒绝。

“撤销AI公式”恢复最近一次成功写入的目标内容，恢复前再次检查当前公式；不覆盖后续手动编辑。撤销仅针对该单格，不自动保存工作簿，不恢复工作簿宏的附带操作。

不支持工具调用的模型仍可解释已提供的公式与样本；用户可将生成公式放入“公式草稿”预览并应用。要完整查找或统计，应在 Chatbox 原生模型设置中选择支持工具调用的模型，并启用 GridPilot Excel MCP。

## 宿主、版本与配置

GridPilot 设置 → AI 助手 → 侧边栏 / 独立窗口；保存后下次打开生效。关闭 Chatbox 窗口再打开，不会丢弃其原生持久化会话。Excel 任务面板的关闭按钮只隐藏侧栏，再点 AI 助手可恢复；Chatbox 自己的关闭按钮结束进程。

侧栏通过 Excel-DNA 任务面板承载独立 Chromium 子窗口。Excel COM 仅在宏主线程访问，进程之间使用限定当前 Windows 用户的随机命名管道；原生 MCP 使用本地回环地址、随机端口、每次启动随机令牌及 Origin 检查。协议参考 [MCP Streamable HTTP](https://modelcontextprotocol.io/specification/2025-03-26/basic/transports)。

同一 Excel 进程可切换工作簿。为保护 Chatbox 原生配置与历史数据库，同一 Windows 用户会话同一时刻只允许一个 Excel 进程启动此助手；第二个进程得到明确提示。

Chatbox 配置/会话使用独立目录 `%APPDATA%\GridPilot\Chatbox-1.20.3`，不读取或改写另装 Chatbox 的配置。模型、密钥、代理、MCP、主题、快捷键、导入导出继续使用原生入口。Chatbox 自身的在线服务或模型可能要求用户账号、订阅或密钥。GridPilot 不提供模型配额。

固定官方安装包 SHA-256：`874171c6c7358ff3915f01412eb036154a5ded7a76c4663fb0f592d2c5c23b13`。
固定官方 app.asar SHA-256：`e43647af8168b731caa9afaed269a4700f8686c257fbc4dd6a254da42729fa01`。

公开 CE 标签存在缺失模块和不一致的依赖锁，不能等同于可复现的完整官方发行构建。因此本版直接提取并校验官方发行包，保留其原生程序。只对固定散列的程序应用五处可审计连接补丁：不启动自动更新/托盘、发送前附加范围摘要、模型调用前展开快照、进入 Excel 侧栏时收起原生会话抽屉、附加 Excel preload。补丁前后散列见发布包 `Chatbox/provenance.json`；Chatbox 更新安装入口保持原生位置但拒绝安装其他版本。准备流程不运行官方安装器、不注册其 URI 协议。

## 构建

运行 `prepare-browser.ps1`、`prepare-chatbox.ps1`、`build.ps1`、`package-release.py`。Chatbox 准备工具依赖固定为 `@electron/asar 3.4.1` 和 `7zip-bin 5.2.0`，使用提交的 npm lock。官方资源与原生依赖按发行包保留。源代码与构建连接层位于 `integrations/chatbox/`，不把 API 密钥、模型测试数据或用户会话打入发布包。

复用 GridPilot 的 Electron 44.7.0 运行时，Chatbox 应用版本固定为 1.20.3。保留官方发行包的原生依赖、许可和 source map；连接层准备脚本与变更完整提交到本仓库。
