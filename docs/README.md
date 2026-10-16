# 文档目录

- [多级排序使用说明](guides/SORTING.txt)
- [复制](guides/COPYING.txt)、[简易脱敏](guides/SIMPLE-MASKING.txt)、[数据库](guides/DATABASE.txt)、[监管核对](guides/RECONCILIATION.txt)
- [架构](development/ARCHITECTURE.txt)、[构建](development/BUILD.txt)
- [v1.6 验证记录](v1.6/VALIDATION.md)、[历史验证](validation/VALIDATION.txt)
- [v1.6 更新说明](releases/RELEASE-v1.6.0.md)，历史更新位于 `releases/`
- 第三方许可原文位于 `licenses/`，清洁便携包中可从“设置 → 关于与许可”阅读。

## 图标规范

功能图标统一 32 单位画布、圆角笔触、透明背景和 1.65 单位线宽，以 4 倍分辨率绘制后缩小。支持 96–288 DPI 及青瓷、晴蓝、石墨三套配色。每个功能保持不同的轮廓或中心标识，复制公式与复制值分别使用 fx / 123；文本与数值分别使用 A / 12。自动测试检查 16 个功能图标互不重复。

导航工具使用小尺寸矢量图标。升降序箭头表示下一次操作，悬停提示说明规则和作用范围；高级排序、刷新、设置各用独立图形。工具始终在首行，主体空间留给工作表条目。

## 导航交互参考

参考 PDF-XChange 官方[书签面板说明](https://help.pdf-xchange.com/pdfxe10/bookmarks_ed.html)及[快捷键表](https://help.pdf-xchange.com/pdfxt10/keyboard-shortcuts_ed.html)：单击跳转、F2 改名、Space 跳转、Insert 新建、Alt+上/下移动。双击改名、失焦提交及 Enter/Esc 按本项目需求实现。

未覆盖 Excel 的全局 Ctrl+B（加粗）。PDF 的书签层级、删除页面等操作不映射为删除 Excel 工作表。
