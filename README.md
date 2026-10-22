<p align="center"><img src="docs/v1.9.1/ai-icon.png" width="64" alt="GridPilot AI 图标"></p>

# GridPilot

让 Excel 里的整理、分析和核对更顺手。

GridPilot 是 Windows 桌面 Excel 的便携工具集：用 AI 助手分析表格、生成公式和图表，用预览与快照管理变更，也可以离线完成排序、导航、复制、脱敏和报表核对。

**当前正式版：v2.0.0** · 文档更新：2026-10-22

[下载正式版](https://github.com/ionize1080/GridPilot/releases/latest) · [快速开始](docs/guides/QUICKSTART.md) · [AI 助手指南](docs/guides/AI.md) · [全部文档](docs/README.md) · [更新记录](docs/releases/README.md)

## 从日常操作到 AI 分析

| 想做什么 | GridPilot 提供什么 | 使用说明 |
| --- | --- | --- |
| 分析当前表格、写公式 | 12 个精简工具入口，按需发现工作簿、透视表、73 种图表、VBA 与插件操作 | [AI 助手](docs/guides/AI.md) |
| 连续应用 AI 建议 | 独立批准开关、危险 VBA 强制手动批准、多版本完整快照 | [批准与恢复](docs/guides/AI.md#批准与恢复) |
| 按业务逻辑排列记录 | 最多 20 层联动排序，中文数字、自定义序列、自然顺序、预览与恢复原序 | [多级排序](docs/guides/SORTING.md) |
| 快速找到、整理工作表 | 搜索、跨工作簿导航、原地改名、标签排序与聚焦高亮 | [导航与常用工具](docs/guides/WORKSHEET-TOOLS.md) |
| 带着排版复制区域 | 保留内部公式或转为数值，输出新工作簿 | [原样复制](docs/guides/COPYING.md) |
| 准备可分享的数据 | 选区简易脱敏、跨表一致映射、离线中英文分词 | [简易脱敏](docs/guides/SIMPLE-MASKING.md) · [智能脱敏](docs/guides/MASKING.md) · [分词](docs/guides/SEGMENTATION.md) |
| 管理多期报表、解释差异 | 本地 SQLite 报表库、来源追溯、两期监管／财务核对 | [数据库](docs/guides/DATABASE.md) · [监管核对](docs/guides/RECONCILIATION.md) |

<table>
<tr>
<td width="40%" valign="top">
<img src="docs/v2.0/native-approval.jpg" alt="v2.0 独立批准开关" width="370">
<p><b>AI 工作区</b>：侧栏与独立窗口互切；工具结果、待应用建议与恢复入口集中显示。</p>
</td>
<td width="60%" valign="top">
<img src="docs/v1.6/sort.png" alt="多级排序工作区：三层规则和结果预览，截图采于 v1.6" width="620">
<p><b>多级排序</b>：每层独立设置规则和方向，先看结果，再应用。</p>
<img src="docs/v1.9.1/applying.png" alt="应用中显示深绿色按钮和保存快照状态" width="620">
<p><b>清楚的执行反馈</b>：应用时显示状态，成功后清理已处理建议。</p>
</td>
</tr>
</table>

截图均来自合成数据测试；批准开关截图为 v2.0，其余 AI 截图为 v1.9.1，排序截图采于 v1.6，具体行为以当前指南为准。

## 三步开始

1. **完整解压** [Release 便携包](https://github.com/ionize1080/GridPilot/releases/latest)，放入有写入权限的目录。保留同目录的 DLL、`App/`、`Browser/`、`Resources/`、`Chatbox/`；不能只复制 XLL。
2. **按 Excel 位数加载**：32 位 Excel 用 `GridPilot-x86.xll`，64 位 Excel 用 `GridPilot-x64.xll`。在 Excel 的加载项管理中浏览并选择；不要同时加载新旧两份。
3. 打开 **GridPilot** 选项卡。AI 助手中先配置自己的模型服务与密钥，再选择「选区／工作表／工作簿／不附加」并发送问题。其他工具可在「设置 → 功能开关」启用。

运行环境是 **64 位 Windows 10/11 + 桌面 Excel + .NET Framework 4.6.2 或更高版本**。运行时随包附带，不需要另装 Node.js、Chrome 或 WebView2。Office 信任策略仍然适用。

已实测环境为 **Windows 11 / Excel 16 x64**。Win10 / Excel 2010 32 位仍待实机验收；x86 测试进程通过不等于 32 位 Excel 验收通过。详见 [兼容性与验收](docs/guides/COMPATIBILITY.md)。

## AI 变更如何落到工作簿

```mermaid
flowchart LR
    A["选择范围并提问"] --> B["本地准备冻结上下文"]
    B --> C["模型调用 Excel 工具"]
    C --> D["生成修改建议"]
    D --> E{"批准方式"}
    E -->|"手动"| F["点击一键应用"]
    E -->|"用户已开启自动批准"| G["检查目标并保存变更前副本"]
    F --> G
    G --> H["执行修改并保存变更后副本"]
    H --> I["查看结果 / 切换变更前后"]
```

自动批准默认关闭，可选择「本工作簿 · 多轮对话有效」或「仅本次对话」。工作簿授权在多轮对话及恢复前后切换后继续有效，关闭工作簿或助手后结束，不会转移给其他工作簿。

输入栏盾牌按钮表示手动批准；开启后显示红色闪电，联网搜索收进更多设置。危险 VBA 不受自动批准开关影响，始终要求查看代码后逐项批准。

「历史快照」可打开任意保存版本的完整工作簿副本，保留原工作簿与未保存编辑；支持被删除的工作表、图表、透视表及 VBA。快照按内容压缩去重，保存在本地。旧基础工具仍提供最新变更前／后快速切换，详见 [批准与恢复](docs/guides/AI.md#批准与恢复)。

## 数据与隐私

- 排序、复制、脱敏、分词、数据库和监管核对在本地处理。
- AI 对话会将所选范围的上下文样本、工具返回结果及主动添加的附件发送到你配置的模型服务。完整本地快照和 Recovery 副本不会自动作为附件发送。
- 「不附加」停止自动加入新的 Excel 上下文，不会删除已有对话中的数据。密钥与聊天配置由独立的本地 Chatbox 配置管理，不随便携发布包分发。
- 备份、数据库和还原记录可能含原始数据。保存位置、清理方法和共享边界见 [数据与存储](docs/guides/DATA-AND-STORAGE.md)。

## v2.0.0 的变化

增强透视表字段、筛选、排序、计算与分组；支持 73 种 Excel 2010 图表以及坐标轴、绘图区、曲线、点、字体和颜色编辑。新增工作簿／工作表／范围、表格、名称、批注、图片、迷你图、CSV、VBA 与常用插件操作。批量重命名和自定义多级排序复用现有引擎。

新增按内容去重的完整历史快照、独立批准按钮和按需 MCP 帮助。调整旧版 Office 侧栏客户区对齐与裁切。真实 Excel、真实 DeepSeek 调用及原生界面测试的具体范围见 [发布说明](docs/releases/RELEASE-v2.0.0.md) · [验证记录](docs/v2.0/VALIDATION.md)。Excel 2010 黑边仍需该版本实机验收，不能以 Office 365 通过代替。

## 开发与文档

[架构与数据流](docs/development/ARCHITECTURE.md) · [构建与测试](docs/development/BUILD.md) · [界面规范](docs/development/UI.md) · [故障排查](docs/guides/TROUBLESHOOTING.md)

源码内部保留 `SheetPilot.*` 命名以兼容既有入口；产品名为 GridPilot。第三方组件包含 Excel-DNA、Electron、Chatbox、MCP SDK 等，许可与来源见 [第三方说明](docs/licenses/THIRD-PARTY-NOTICES.txt) 和 [Chatbox 集成说明](docs/licenses/CHATBOX-NOTICES.txt)。此项目不是 Microsoft、Chatbox 或 Excel-DNA 的官方产品。
