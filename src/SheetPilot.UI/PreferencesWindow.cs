using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public sealed class PreferencesWindow : Window
    {
        private readonly StackPanel sections = new StackPanel();
        private readonly StackPanel categories = new StackPanel();
        private readonly ScrollViewer scroll;
        private readonly Dictionary<string, FrameworkElement> anchors = new Dictionary<string, FrameworkElement>();
        private readonly Dictionary<string, ToggleButton> buttons = new Dictionary<string, ToggleButton>();
        public string ActiveSection { get; private set; } = "general";
        public double ScrollOffset => scroll.VerticalOffset;

        public PreferencesWindow(ToolPreferences preferences, NavigationSettings colors, Func<string> save, Action navigationSettings)
        {
            ToolUi.Theme(this, colors, "设置", 980, 820);
            NameScope.SetNameScope(this, new NameScope());
            MinWidth = Math.Min(780, Width);
            var root = new Grid(); Content = root;
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(188) });
            root.ColumnDefinitions.Add(new ColumnDefinition());

            var sidebar = new DockPanel { Margin = new Thickness(18, 24, 14, 20) };
            root.Children.Add(sidebar);
            var brand = new StackPanel { Margin = new Thickness(12, 0, 0, 24) };
            brand.Children.Add(new TextBlock { Text = "设置", FontSize = 26, FontWeight = FontWeights.SemiBold });
            brand.Children.Add(ToolUi.Note("GridPilot")); DockPanel.SetDock(brand, Dock.Top); sidebar.Children.Add(brand);
            var aside = ToolUi.Note("偏好随加载项文件夹携带"); aside.Margin = new Thickness(12, 16, 0, 0);
            DockPanel.SetDock(aside, Dock.Bottom); sidebar.Children.Add(aside);
            sidebar.Children.Add(ToolUi.Scroll(categories));

            var right = new DockPanel { Margin = new Thickness(16, 24, 24, 18) }; Grid.SetColumn(right, 1); root.Children.Add(right);
            var footer = new DockPanel { Margin = new Thickness(0, 14, 0, 0), LastChildFill = true };
            DockPanel.SetDock(footer, Dock.Bottom); right.Children.Add(footer);
            var actions = new StackPanel { Orientation = Orientation.Horizontal }; DockPanel.SetDock(actions, Dock.Right); footer.Children.Add(actions);
            var message = ToolUi.Note("关闭时自动保存"); message.VerticalAlignment = VerticalAlignment.Center; footer.Children.Add(message);
            ToolUi.Button(actions, "保存", () => { KeyboardCommit(); preferences.Normalize(); message.Text = save(); });
            ToolUi.Button(actions, "完成", Close);
            scroll = ToolUi.Scroll(sections); scroll.Padding = new Thickness(0, 0, 14, 0); RegisterName("SettingsScroll", scroll); right.Children.Add(scroll);

            var general = Section("general", "通用", "统一外观与使用偏好");
            var palettes = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 2, 0, 12) }; general.Children.Add(palettes);
            ToolUi.Button(palettes, "浅绿", () => colors.ApplyPalette("green"));
            ToolUi.Button(palettes, "浅灰", () => colors.ApplyPalette("neutral"));
            ToolUi.Button(palettes, "深色", () => colors.ApplyPalette("dark"));
            Color(general, colors, "BackgroundColor", "页面背景"); Color(general, colors, "SurfaceColor", "卡片与控件");
            Color(general, colors, "TextColor", "主要文字"); Color(general, colors, "MutedTextColor", "辅助文字"); Color(general, colors, "AccentColor", "强调色");
            ToolUi.Check(general, colors, "ShowTooltips", "显示导航悬停提示与点击说明");
            general.Children.Add(ToolUi.Note("配色同时用于设置和导航。Excel 自身的菜单、窗口边框与 Ribbon 由 Office 管理。"));

            var nav = Section("navigation", "工作表导航", "范围、布局与信息密度");
            ToolUi.Choice(nav, colors, "ScopeIndex", "默认范围", "当前工作簿", "所有已打开工作簿（同一 Excel）");
            ToolUi.Choice(nav, colors, "DockSide", "停靠位置", "右侧", "左侧");
            ToolUi.Number(nav, colors, "PaneWidth", "面板宽度", 260, 1600);
            ToolUi.Number(nav, colors, "RowHeight", "结果行高", 28, 80); ToolUi.Number(nav, colors, "RowGap", "结果行间距", 0, 10);
            ToolUi.Number(nav, colors, "FontSize", "结果字号", 11, 18);
            ToolUi.Choice(nav, colors, "WorkbookNameMode", "工作簿名", "自动（当前工作簿时省略）", "始终显示", "始终隐藏");
            ToolUi.Check(nav, colors, "IncludeHidden", "显示隐藏页面"); ToolUi.Check(nav, colors, "ShowSheetNumbers", "显示页面序号");
            ToolUi.Check(nav, colors, "ShowSheetType", "显示普通工作表类型"); ToolUi.Check(nav, colors, "ShowHeader", "显示面板品牌栏");
            ToolUi.Check(nav, colors, "ShowFooter", "显示底部工作簿信息"); ToolUi.Check(nav, colors, "ShowKeyboardHints", "显示底部快捷键说明");
            ToolUi.Check(nav, colors, "AutoRefresh", "自动刷新目录"); ToolUi.Number(nav, colors, "RefreshSeconds", "目录刷新间隔（秒）", 1, 10);
            var navActions = new StackPanel { Margin = new Thickness(0, 12, 0, 0) }; nav.Children.Add(navActions);
            ToolUi.Button(navActions, "在导航面板中调整", () => { Close(); navigationSettings(); });

            var focus = Section("focus", "聚焦高亮", "临时覆盖，不改动单元格格式");
            Color(focus, preferences, "FocusColor", "高亮颜色");
            ToolUi.Number(focus, preferences, "FocusOpacity", "不透明度（%）", 5, 60);
            ToolUi.Number(focus, preferences, "FocusInterval", "滚动补偿间隔（毫秒）", 16, 100);
            ToolUi.Check(focus, preferences, "FocusRow", "高亮当前行"); ToolUi.Check(focus, preferences, "FocusColumn", "高亮当前列");
            ToolUi.Check(focus, preferences, "ClearActiveCell", "保持活动单元格清晰，不覆盖它");
            focus.Children.Add(ToolUi.Note("选区变化立即跟随。进入文件菜单、弹出菜单或离开 Excel 时暂停显示，回到表格后恢复；不随工作簿保存或打印。每次启动默认关闭。"));

            var conversion = Section("conversion", "文本与数值", "转换规则与数据保护");
            ToolUi.Check(conversion, preferences, "PreserveVisualAlignment", "保持视觉位置（将常规对齐固定为左 / 右）");
            ToolUi.Check(conversion, preferences, "VisibleCellsOnly", "只处理可见单元格");
            ToolUi.Check(conversion, preferences, "PreserveLeadingZeros", "保留前导零文本（如 00123）");
            ToolUi.Check(conversion, preferences, "ProtectPrecision", "跳过超过15位有效数字，保护长编号");
            ToolUi.Check(conversion, preferences, "TrimNumericText", "解析前忽略首尾空格");
            ToolUi.Check(conversion, preferences, "AllowThousands", "识别规范的千位分隔符");
            ToolUi.Check(conversion, preferences, "AllowPercent", "识别百分比文本（12% → 0.12）");
            ToolUi.Check(conversion, preferences, "NumberAsTextHint", "转文本后允许显示绿色数字文本角标");
            ToolUi.Choice(conversion, preferences, "DecimalCulture", "数字分隔符", "遵循 Excel 小数点 / 千位分隔符", "固定英文格式（1,234.56）");
            ToolUi.Number(conversion, preferences, "ConversionLimit", "每次最多处理单元格数", 100, 200000);
            conversion.Children.Add(ToolUi.Note("只转换常量；公式、空白、错误、合并单元格和锁定单元格跳过。日期型数值转为底层序列文本。绿色角标由 Excel 后台错误检查控制。"));

            if (preferences.Copy == null) preferences.Copy = new CopyOptions();
            var copy = Section("copy", "原样复制", "两个复制按钮共用；输出到新工作簿，保持原单元格位置");
            var copyOptions = preferences.Copy;
            ToolUi.Choice(copy, copyOptions, "Scope", "复制范围", "当前选择区域（支持不连续选区）", "当前工作表已使用区域");
            ToolUi.Check(copy, copyOptions, "EmptySelectionUsesUsedRange", "选中单个空白格时改为复制已使用区域");
            ToolUi.Check(copy, copyOptions, "TrimWholeRowsColumns", "整行 / 整列选择自动裁到已使用区域");
            ToolUi.Check(copy, copyOptions, "VisibleCellsOnly", "只复制可见单元格（保留位置，隐藏处留空）");
            copy.Children.Add(ToolUi.Heading("排版与外观"));
            ToolUi.Check(copy, copyOptions, "CellFormats", "保留字体、填充、边框与对齐");
            ToolUi.Check(copy, copyOptions, "NumberFormats", "保留数字格式（日期、百分比、小数位等）");
            ToolUi.Check(copy, copyOptions, "RowHeights", "保留行高");
            ToolUi.Check(copy, copyOptions, "ColumnWidths", "保留列宽");
            ToolUi.Check(copy, copyOptions, "HiddenState", "保留行列隐藏状态（仅在复制全部单元格时）");
            ToolUi.Check(copy, copyOptions, "MergedCells", "保留合并单元格");
            ToolUi.Check(copy, copyOptions, "Notes", "保留单元格批注（传统备注）");
            ToolUi.Check(copy, copyOptions, "ViewSettings", "保留缩放比例、网格线和行列标题显示");
            ToolUi.Choice(copy, copyOptions, "ConditionalFormatting", "条件格式", "固定当前颜色、字体与边框，不复制规则", "仅保留基础格式，移除条件格式");
            copy.Children.Add(ToolUi.Note("固定外观不包含数据条或图标集；遇到这类规则会提示调整设置。图片、图表、形状、数据验证、超链接及打印设置不随选区复制。合并区域必须完整选中。"));
            copy.Children.Add(ToolUi.Heading("公式与完成行为"));
            ToolUi.Choice(copy, copyOptions, "FormulaFallback", "内部公式版：遇到范围外或无法确认依赖的公式", "转为当前计算值，并在结果提示中统计", "停止复制，提示首个不能保留的公式");
            copy.Children.Add(ToolUi.Note("保留明确的范围内 A1 引用及支持的普通函数。跨表 / 跨簿、命名区域、结构化引用、动态引用、数组 / 溢出和未知函数转值或停止。仅数值版把全部公式转为当前值；复制前请确认源表已计算完成。"));
            ToolUi.Check(copy, copyOptions, "KeepSourceActive", "完成后仍停留在源工作表");
            copy.Children.Add(ToolUi.Heading("大范围保护"));
            ToolUi.Number(copy, copyOptions, "CellLimit", "每次最多复制单元格数", 100, 5000000);
            ToolUi.Number(copy, copyOptions, "AppearanceLimit", "最多固定条件格式单元格数", 100, 100000);
            copy.Children.Add(ToolUi.Note("复制时状态栏显示阶段，可按 Esc 取消；单次 Excel 内部粘贴结束后才会响应。出错或取消时关闭未完成的结果。源表不写入，结果不自动保存；如需撤销复制，关闭结果工作簿即可。"));

            var rename = Section("rename", "批量重命名", "下次打开时使用的默认规则");
            rename.Children.Add(ToolUi.RenameControls(preferences.Rename, true));
            rename.Children.Add(ToolUi.Note("应用前仍需检查预览。修改不自动保存工作簿；GridPilot 可撤销最近一次成功的修改，目标内容需未被再次编辑。"));

            var about = Section("about", "关于", "GridPilot · 便携式 Excel 工具集");
            about.Children.Add(new TextBlock { Text = "GridPilot " + typeof(PreferencesWindow).Assembly.GetName().Version.ToString(3), FontSize = 18, Margin = new Thickness(0, 4, 0, 10) });
            about.Children.Add(ToolUi.Note("原生 Ribbon 与桌面界面，无内嵌浏览器。配置保存在 XLL 所在文件夹：GridPilot.navigation.json 与 GridPilot.tools.json。移动或升级时保留这两个文件即可携带偏好。"));
            about.Children.Add(ToolUi.Note("支持 .NET Framework 4.6.2 及之后版本。遇到问题时可查看发布包内 VALIDATION.txt 与使用说明。"));
            // Leave enough trailing space for the last section to align at the top when selected.
            var tail = new Border(); sections.Children.Add(tail);
            scroll.SizeChanged += (s, e) => tail.Height = Math.Max(0, e.NewSize.Height - anchors["about"].ActualHeight + 32);
            scroll.ScrollChanged += (s, e) => UpdateCategory();
            Loaded += (s, e) => NavigateToSection("general");
            Closing += (s, e) => KeyboardCommit();
            PropertyChangedEventHandler recolor = (s, e) => { if (e.PropertyName.EndsWith("Color", StringComparison.Ordinal)) ViewTheme.Apply(Resources, colors); };
            colors.PropertyChanged += recolor; Closed += (s, e) => colors.PropertyChanged -= recolor;
        }

        private StackPanel Section(string id, string title, string subtitle)
        {
            var wrapper = new StackPanel { Margin = new Thickness(0, 0, 0, 28) }; anchors.Add(id, wrapper); sections.Children.Add(wrapper);
            wrapper.Children.Add(new TextBlock { Text = title, FontSize = 21, FontWeight = FontWeights.SemiBold });
            var note = ToolUi.Note(subtitle); note.Margin = new Thickness(0, 5, 0, 14); wrapper.Children.Add(note);
            var panel = new StackPanel();
            var card = new Border { CornerRadius = new CornerRadius(10), Padding = new Thickness(20, 16, 20, 18), BorderThickness = new Thickness(1), Child = panel };
            card.SetResourceReference(Border.BackgroundProperty, "SurfaceBrush"); card.SetResourceReference(Border.BorderBrushProperty, "BorderBrush"); wrapper.Children.Add(card);
            var button = new ToggleButton { Content = new TextBlock { Text = title, Width = 122 }, MinHeight = 42, Margin = new Thickness(0, 0, 0, 6), IsChecked = id == "general" };
            button.Click += (s, e) => NavigateToSection(id); categories.Children.Add(button); buttons.Add(id, button); RegisterName("Category_" + id, button);
            return panel;
        }
        private static void Color(Panel panel, object source, string path, string caption)
        { var editor = new ColorEditor { Caption = caption }; ToolUi.Bind(editor, ColorEditor.ValueProperty, source, path); panel.Children.Add(editor); }
        public void NavigateToSection(string id)
        {
            if (!anchors.TryGetValue(id, out var anchor)) return;
            KeyboardCommit(); sections.UpdateLayout();
            scroll.ScrollToVerticalOffset(anchor.TranslatePoint(new Point(), sections).Y); SetCategory(id);
        }
        private void UpdateCategory()
        {
            var id = anchors.Where(pair => pair.Value.TranslatePoint(new Point(), sections).Y <= scroll.VerticalOffset + 24).Select(pair => pair.Key).LastOrDefault();
            if (id != null) SetCategory(id);
        }
        private void SetCategory(string id) { ActiveSection = id; foreach (var pair in buttons) pair.Value.IsChecked = pair.Key == id; }
        private void KeyboardCommit() { FocusManager.SetFocusedElement(this, this); Keyboard.ClearFocus(); }
    }
}
