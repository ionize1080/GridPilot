using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public sealed class SortWindow : Window
    {
        public SortOptions Options { get; }
        public int Column { get; private set; }
        private readonly TextBlock status;
        private readonly ListView preview;
        private readonly Button apply;
        private readonly Func<int,bool,IReadOnlyList<string>> read;
        private readonly Func<int,SortOptions,string> execute;
        public SortWindow(NavigationSettings colors, SortOptions saved, string description, string[] columns,
            int initialColumn, Func<int,bool,IReadOnlyList<string>> read, Func<int,SortOptions,string> execute, bool sheets=false, Func<string> undo=null)
        {
            this.read=read;this.execute=execute;Options=saved.Clone();Column=initialColumn;
            ToolUi.Theme(this,colors,sheets?"工作表排序":"自定义排序",940,740);
            var root=new DockPanel{Margin=new Thickness(22)};Content=root;
            var footer=new StackPanel();DockPanel.SetDock(footer,Dock.Bottom);root.Children.Add(footer);
            status=ToolUi.Note("修改规则后预览，再应用。排序后可使用 GridPilot 撤销。");footer.Children.Add(status);
            var actions=new StackPanel{Orientation=Orientation.Horizontal,HorizontalAlignment=HorizontalAlignment.Right};footer.Children.Add(actions);
            ToolUi.Button(actions,"预览排序",Preview);
            Button undoButton=null;
            apply=ToolUi.Button(actions,"应用排序",()=>{try{status.Text=execute(Column,Options.Clone());apply.IsEnabled=false;if(undoButton!=null)undoButton.IsEnabled=true;}catch(Exception ex){status.Text=ex.Message;}});apply.IsEnabled=false;
            if(undo!=null){undoButton=ToolUi.Button(actions,"恢复原序",()=>{try{status.Text=undo();undoButton.IsEnabled=false;apply.IsEnabled=false;}catch(Exception ex){status.Text=ex.Message;}});undoButton.IsEnabled=false;}
            ToolUi.Button(actions,"关闭",Close);
            var top=new StackPanel();DockPanel.SetDock(top,Dock.Top);root.Children.Add(top);
            top.Children.Add(ToolUi.Heading(sheets?"调整工作表标签顺序":"按完整记录排序"));top.Children.Add(ToolUi.Note(description));
            var grid=new Grid();grid.ColumnDefinitions.Add(new ColumnDefinition{Width=new GridLength(1,GridUnitType.Star)});grid.ColumnDefinitions.Add(new ColumnDefinition{Width=new GridLength(1,GridUnitType.Star)});root.Children.Add(grid);
            var left=new StackPanel{Margin=new Thickness(0,0,18,0)};var scroll=new ScrollViewer{Content=left,VerticalScrollBarVisibility=ScrollBarVisibility.Auto};grid.Children.Add(scroll);
            var right=new DockPanel();Grid.SetColumn(right,1);grid.Children.Add(right);var label=ToolUi.Heading("排序预览 · 前 200 项");DockPanel.SetDock(label,Dock.Top);right.Children.Add(label);
            preview=new ListView();preview.SetResourceReference(Control.ForegroundProperty,"TextBrush");var view=new GridView();view.Columns.Add(new GridViewColumn{Header="原位置",DisplayMemberBinding=new System.Windows.Data.Binding("Position"),Width=70});view.Columns.Add(new GridViewColumn{Header="排序结果",DisplayMemberBinding=new System.Windows.Data.Binding("Value"),Width=270});preview.View=view;right.Children.Add(preview);
            if(!sheets)
            {
                left.Children.Add(ToolUi.Note("排序依据列（移动范围内整条记录）"));var column=new ComboBox{ItemsSource=columns,SelectedIndex=initialColumn,MinHeight=30};column.SelectionChanged+=(s,e)=>{Column=column.SelectedIndex;Invalidate();};left.Children.Add(column);
                var header=ToolUi.Check(left,Options,nameof(Options.HasHeader),"第一行为表头，不参与排序");header.Click+=(s,e)=>Invalidate();
            }
            var presets=new ComboBox{ItemsSource=new[]{"选择预设…","中文数字","大中小微","国际 → 业务 + 中文数字","大中小微 + 中文数字","名称自然升降序"},SelectedIndex=0,MinHeight=32};left.Children.Add(ToolUi.Heading("预设与组合"));left.Children.Add(presets);
            left.Children.Add(ToolUi.Note("每行一条规则，靠前的组先排。默认从开头匹配；前缀未知时可加 .*? 。例如：(国际)[中文数字]"));
            var rules=new TextBox{Text=Options.Rules,AcceptsReturn=true,Height=106,MaxLength=8192,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(8)};left.Children.Add(rules);rules.TextChanged+=(s,e)=>{Options.Rules=rules.Text;Invalidate();};
            left.Children.Add(ToolUi.Note("自定义序列：每行 名称=项目1,项目2,…，规则中用 [名称] 引用。例：级别=特大,大型,中型,小型,微型"));
            var sequences=new TextBox{Text=Options.Sequences,AcceptsReturn=true,Height=70,MaxLength=8192,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(8)};left.Children.Add(sequences);sequences.TextChanged+=(s,e)=>{Options.Sequences=sequences.Text;Invalidate();};
            var natural=ToolUi.Check(left,Options,nameof(Options.Natural),"名称自然排序（忽略规则，数字 2 在 10 前）");natural.Click+=(s,e)=>Invalidate();
            var desc=ToolUi.Check(left,Options,nameof(Options.Descending),"组内降序（规则优先级保持不变）");desc.Click+=(s,e)=>Invalidate();
            var match=ToolUi.Check(left,Options,nameof(Options.MatchCase),"区分英文字母大小写");match.Click+=(s,e)=>Invalidate();
            left.Children.Add(ToolUi.Note("中文数字支持零、十、百、千、万、亿等大数及大写数字。组合示例：[大中小微][中文数字]。未匹配项置后，相同键值保持原序。普通正则字符组如 [A-Z] 仍可使用。"));
            presets.SelectionChanged+=(s,e)=>
            {
                int i=presets.SelectedIndex;if(i<1)return;
                Options.Natural=i==5;natural.IsChecked=Options.Natural;
                if(i<5)rules.Text=new[]{"[中文数字]","[大中小微]","(国际)[中文数字]\n(业务)[中文数字]","[大中小微][中文数字]"}[i-1];Invalidate();
            };
            Loaded+=(s,e)=>Preview();
        }
        private void Invalidate(){apply.IsEnabled=false;status.Text="规则已改变，请重新预览。";}
        private void Preview()
        {
            apply.IsEnabled=false;
            try{var values=read(Column,Options.HasHeader);var order=CustomSort.Order(values,Options);preview.ItemsSource=order.Take(200).Select(i=>new{Position=i+1,Value=values[i]}).ToArray();status.Text="共 "+values.Count+" 项；已预览。应用会调整实际顺序，工作簿不会自动保存。";apply.IsEnabled=values.Count>1;}
            catch(Exception ex){status.Text=ex.Message;preview.ItemsSource=null;}
        }
    }
}
