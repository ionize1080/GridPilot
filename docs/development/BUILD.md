# 构建、测试与发布

[文档首页](../README.md) · 适用 v2.1.0 · 2026-10-22

## 开发环境

使用 Windows 交互桌面、PowerShell、Git、Node.js/npm，以及 [global.json](../../global.json) 固定的 **.NET SDK 8.0.206**（同系列最新补丁可用）。编译目标是 .NET Framework 4.6.2，引用程序集通过锁定 NuGet 包获取。

最终用户不需要开发 SDK 或 Node.js。无 Excel 可编译并运行基础／界面逻辑测试；实际 COM、窗口和真实 API 集成测试需要安装桌面 Excel。不要在存有未保存业务数据的工作会话中做自动化验收。

## 准备与构建

在仓库根目录执行：

```powershell
powershell -NoProfile -File .\prepare-browser.ps1
powershell -NoProfile -File .\prepare-chatbox.ps1
powershell -NoProfile -File .\build.ps1
```

| 步骤 | 作用 |
| --- | --- |
| prepare-browser | 下载并校验 Electron 44.7.0 Windows x64，解压到 .temp/Browser |
| prepare-chatbox | 校验固定 Chatbox 1.20.3 安装包，npm ci 恢复锁定工具，提取并施加集成补丁；不执行安装器 |
| build | NuGet locked-mode 恢复、Release 构建、基础套件、准备完整 dist 和 SHA256SUMS.txt |

`build.ps1` 接受 `-PackageSource` 指向可用 NuGet 源，`-BrowserRuntime` 指向已准备运行时。不要将不匹配版本的运行时直接替换进去。Chatbox 的原始／修改后 hash 记录在准备输出 `Chatbox/provenance.json`。

构建默认执行普通测试、color-picker-host、release-features、navigation-host、upgrade、v1、v11、multi-sort、ai，以及 Electron 下的 reconcile、reconcile-manager、ai-tools。它不会自动运行需要 Excel 的专项或付费 API 测试。以 [build.ps1](../../build.ps1) 和 [测试入口](../../tests/SheetPilot.Tests/Program.cs) 为准。

## 分层验证

先构建，再按变更选择专项；不要把 x86 测试进程和 32 位 Excel 宿主混为一谈。

```powershell
$testExe = '.\tests\SheetPilot.Tests\bin\Release\net462\SheetPilot.Tests.exe'
& $testExe --ai --excel
& $testExe --multi-sort --excel
& $testExe --copy-excel .\dist\GridPilot-x64.xll
& $testExe --simple-mask --excel
& $testExe --v11 --excel
& $testExe --v1 --excel --large
```

| 测试 | 覆盖 |
| --- | --- |
| 默认测试、--render 输出目录 | 规则、配置、WPF 控件与图标 |
| --color-picker-host / --navigation-host | 原生容器中的色盘、焦点、菜单与导航 |
| --excel 完整XLL路径 | 实际 XLL 加载、导航与宿主行为 |
| --ai --excel | 10 项工具后端、建议、批准、冲突、图表／透视表、恢复 |
| --multi-sort --excel、--copy-excel | 原生排序、复制、恢复与源保护 |
| --simple-mask --excel、--v1 --excel --large | 脱敏、表头、输出与大范围 |
| --v11 --excel、reconcile*.cjs | 报表读取、数据库／核对逻辑及管理 |
| browser-ai.cjs | Chatbox 原生 UI + 本地模拟模型 + MCP 工具循环 |
| native-ai-switch.cjs | 实际 packed XLL、原生窗口双向切换和草稿保留 |
| live-ai.cjs | 用户明确启用的真实模型 API → MCP → 实际 Excel |

浏览器测试需要开发机可用的 Playwright Node 模块，可通过 `PLAYWRIGHT_MODULE` 指定绝对模块目录；不得将测试调试端口、临时模型服务或测试用户配置打入生产包。

```powershell
node .\tests\browser-ai.cjs
node .\tests\native-ai-switch.cjs
```

`native-ai-switch.cjs` 在独立临时目录注入测试 CDP，使用当前的 x64 packed XLL；设置 GRIDPILOT_TEST_PACKAGE 为已解压发布目录时，测试直接使用该包的 XLL、DLL 和运行时。这条测试本身不能验收 32 位 Excel。测试结束核对自己的 Excel／Electron 实例已退出，不终止用户已有 Excel。

## 真实 API 测试与密钥

仅在有授权使用相应 API 时运行，会产生服务费用。密钥保存在仓库之外，环境变量只传文件路径，不能把密钥本身写入命令、源码、日志或截图。

```powershell
$env:GRIDPILOT_TEST_KEY_FILE = 'D:\private\deepseek-key.txt'
$env:GRIDPILOT_TEST_MODEL = 'deepseek-flash'
try {
    node .\tests\live-ai.cjs
    if ($LASTEXITCODE -ne 0) { throw 'Live API test failed.' }
    python .\tests\secret-scan.py
    if ($LASTEXITCODE -ne 0) { throw 'Credential scan failed.' }
} finally {
    Remove-Item Env:GRIDPILOT_TEST_KEY_FILE -ErrorAction SilentlyContinue
    Remove-Item Env:GRIDPILOT_TEST_MODEL -ErrorAction SilentlyContinue
}
```

文件路径是示例。当前脚本从外部文件提取密钥，以独立临时 Chatbox 配置测试，最后清理配置；输出不得包含密钥。泄漏扫描检查源码、dist 和发布 ZIP 的真实密钥内容（UTF-8、UTF-16）；它不是对全部可能秘密的通用保证。

测试使用合成工作簿，不上传业务表格。失败时检查临时目录与残留配置，不将包含敏感信息的诊断直接上传。当前已执行结果见 [验证总览](../validation/README.md)。

## 发布检查

1. 确认 main 的目标提交、版本号、当前指南与 Release 说明一致。
2. 完整构建并运行相关专项，明确记录平台、宿主位数、测试进程位数和未完成验收。
3. 从 packed XLL 构建完整便携目录，保留运行时、词典、Chatbox 原生依赖、源码映射、许可和 provenance。
4. 排除用户配置、密钥、真实工作簿、Recovery、数据库、测试端口与临时日志；核对 ZIP CRC 和 SHA-256。
5. 用待发布包做加载／窗口冒烟，校验发布文件与记录的 hash；按授权推送提交、标签和正式 Release。
6. 本地升级前关闭 Excel、备份旧目录和设置；逐文件核对新版并验证设置保留。

文档单独更新可推送 main，不需修改二进制版本或覆盖既有 Release 资产。历史 Release 对应其原始代码和验收记录，当前指南以 main 为准。

## v2.0 扩展验收

`SheetPilot.Tests.exe --ai-extended --excel`、`--ai-chart-matrix`、`--ai-formats` 验证高级操作、完整图表矩阵、四种格式与 VBA 审批。`node tests/live-extended.cjs` 从 `GRIDPILOT_TEST_KEY_FILE` 指定的仓库外文件读取真实密钥，仅输出脱敏结果。Node SQLite 测试使用包内 Electron 的 Node 模式。

## v2.1 专项

`SheetPilot.Tests.exe --ai-v21 --excel` 验证预算淘汰、跨工作簿及隐藏表、批准隔离、新版工具。`tests/browser-v21-settings.cjs` 验证总设置；`tests/native-ai-switch.cjs` 验证真实原生风险确认框及窗口切换。`tests/live-v21.cjs` 使用仓库外 `GRIDPILOT_TEST_KEY_FILE`，真实模型读取两本合成工作簿和 VBA，并执行现代工具。

依赖已缓存时可用 `build.ps1 -PackageSource C:/Users/<user>/.nuget/packages` 离线锁定恢复，不修改系统代理或依赖版本。
