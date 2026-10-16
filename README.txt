GridPilot 1.5.0 portable

1. Fully extract the ZIP on 64-bit Windows 10/11. No installer and no WebView2 dependency.
2. Choose GridPilot-x86.xll for 32-bit Excel, or GridPilot-x64.xll for 64-bit Excel.
3. Keep all DLLs, App/, Browser/, Resources/ beside the XLL. Never copy the XLL alone.
4. Close Excel before upgrading. Preserve GridPilot.navigation.json and GridPilot.tools.json.

Runtime: .NET Framework 4.6.2+ already present; Electron 44.7.0 bundled for Windows x64.
Complex windows start Chromium on demand and communicate by a current-user named pipe.
No listening network port or remote web content. No global keyboard/mouse hooks.
Temporary browser profile is removed on normal exit. Abnormal termination may leave a
GridPilot-ui-* folder in TEMP; workbook snapshots are not written there.

Simple masking (Ribbon: 简易脱敏) supports selected ranges/rows/columns, fixed initials,
character swaps, regex and numeric arithmetic. Headers are detected and retained by default;
review detection or specify header rows manually. Preview before copying/exporting.
In-place changes require backup confirmation and support Undo. Recovery JSON contains
original sensitive values: keep it separate. See SIMPLE-MASKING.txt for limits.

Smart masking creates a new XLSX and refuses to overwrite existing files. Review suggestions
and the preview before exporting. Amount/date rules are opt-in. Basic formatting can be
preserved, including merged cells. Conditional formatting rules, shapes, notes, macros,
external links, names and connections are not copied.

Tested here: Windows 11 + Excel 16 x64. Target Windows 10 21H1 + Excel 2010 has NOT been
verified on a real machine. See RELEASE-v1.5.0.md and VALIDATION.txt for exact coverage.
Portable does not bypass Office trust policies or enterprise execution restrictions.
