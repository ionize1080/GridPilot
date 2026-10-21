# v1.8.0 validation

> 此页保留该版本的设计／验证记录；后续行为以 [当前指南](../README.md) 与 [当前架构](../development/ARCHITECTURE.md) 为准。

Environment: Windows 11, Excel 16 x64, .NET Framework 4.6.2 target, bundled Electron 44.7.0, pinned Chatbox 1.20.3. Test workbooks and attachment fixtures are synthetic. Existing user workbooks and Chatbox profiles were not used for model tests.

## Automated checks

- Release builds: zero warnings/errors. Full existing regression build passed (baseline, release features, native navigation/color host, upgrades, settings, copy/mask, multi-sort, reconciliation and AI mapping). Final AI changes were rebuilt and their affected tests rerun.
- AI core: 16 assertions. Actual Excel: 44 assertions, including 200,001 populated cells, exact statistics, full worksheet logical size of 17,179,869,184 cells and XFD1048576; sparse search and snapshot expiration; values/formulas, whole-batch undo, manual-edit conflict preflight, rollback after a later formula failure, range format/no-fill restoration, native chart creation and edited-chart undo refusal.
- Real XLL host: five assertions each in sidebar and independent modes. Chromium has no WS_CHILD style; only sidebar mode has Excel ownership. Reopening retains the existing host, and closing disconnects the bridge.
- Production Chatbox renderer: native MCP stdio transport and all nine tool definitions; native model-to-tool round trip; one-click apply/undo; scope selection and opt-out; narrow toolbar layout; original native settings. Synthetic HTTP model fixture and debugging ports exist only in the test harness.
- Native attachment parser: Chinese TXT, text PDF and DOCX; actual picker path compatibility; extracted PDF text included in a native model request.
- Native request checks: provider request path and headers were exercised with a synthetic local model. Wrapper identity remains native Chatbox 1.20.3; no custom AI authentication or header injection. No paid/live cloud model request was sent.
- Settings renderer: AI enable/disable persistence and display mode, with existing feature/mask regression coverage.

## Computer-use checks

Native computer use reproduced the original WS_CHILD defect: typing in the sidebar modified the selected cell. This led to replacement with an Excel-owned top-level window. After replacement, actual typing in the independent Chatbox window and switching that same live window to the sidebar preserved the draft; the synthetic worksheet sentinel remained unchanged. The computer-use provider cannot bind the Excel-owned top-level window as a separate usable app window, so a second physical typing action *after docking* could not be independently exercised through that provider. Native host ownership/style tests and renderer input tests cover the resulting architecture, but this limitation is not counted as a passed physical-sidebar typing test.

## Release checks and limits

Both portable and clean archives receive CRC and SHA-256 checks. Local installation preserves existing settings and retains the previous installation for rollback. Installed XLL smoke checks use a temporary Chatbox profile and synthetic workbook. GitHub asset size/digest checks are performed during upload.

Windows 10, Excel 2010 and 32-bit Excel were not tested on physical machines. Large ranges remain subject to available resources and execution time; scan limits are no longer fixed at 200,000 cells. Scanned-PDF OCR depends on native Chatbox configuration. Protected/hidden/merged/array/spill/table write targets and mixed-style formatting remain deliberately unsupported. Upstream Excel MCP components are selectively adapted; this release does not claim all upstream capabilities.
