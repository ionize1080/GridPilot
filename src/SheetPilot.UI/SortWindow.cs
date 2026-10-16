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
            ToolUi.Theme(this,colors,sheets?"工作表排序":"自定义排序",1000,760);
            ShowInTaskbar=true; // Keep the editor reachable when switching between applications.
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
                left.Children.Add(ToolUi.Heading("1 · 选择依据列"));var column=new ComboBox{ItemsSource=columns,SelectedIndex=initialColumn,MinHeight=30};column.SelectionChanged+=(s,e)=>{Column=column.SelectedIndex;Invalidate();};left.Children.Add(column);
                var header=ToolUi.Check(left,Options,nameof(Options.HasHeader),"第一行为表头，不参与排序");header.Click+=(s,e)=>Invalidate();
            }
            left.Children.Add(ToolUi.Heading(sheets?"1 · 选一个示例开始":"2 · 选一个示例开始"));
            var examples=new DockPanel();left.Children.Add(examples);
            var sampleButton=new Button{Content="套用示例",Margin=new Thickness(8,0,0,0)};DockPanel.SetDock(sampleButton,Dock.Right);examples.Children.Add(sampleButton);
            var presets=new ComboBox{Name="SortExamples",ItemsSource=new[]{"客户一处、客户二处…（中文数字）","大 → 中 → 小 → 微","国际部门优先，再排业务部门","大中小微 × 中文数字","自定义顺序：准备 → 实施 → 验收","名称自然排序：表2 → 表10"},SelectedIndex=0,MinHeight=32};examples.Children.Add(presets);
            var exampleNote=ToolUi.Note("");left.Children.Add(exampleNote);
            left.Children.Add(ToolUi.Heading("排序规则 · 决定先后顺序"));
            left.Children.Add(ToolUi.Note("每行一组，第一行的组排最前；[中文数字] 表示按数值排序。普通文字按名称开头匹配，.*? 表示跳过前面的文字。"));
            var rules=new TextBox{Name="SortRules",Text=Options.Rules,AcceptsReturn=true,Height=94,MaxLength=8192,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(8)};System.Windows.Automation.AutomationProperties.SetName(rules,"排序规则");left.Children.Add(rules);rules.TextChanged+=(s,e)=>{Options.Rules=rules.Text;Invalidate();};
            var sequenceBody=new StackPanel();
            var sequenceExpander=new Expander{Header="自定义顺序（可选）· 内置中文数字、大中小微无需填写",Content=sequenceBody,IsExpanded=!string.IsNullOrWhiteSpace(Options.Sequences),Margin=new Thickness(0,12,0,8)};left.Children.Add(sequenceExpander);
            sequenceBody.Children.Add(ToolUi.Note("这里只定义你自己的排列顺序，不填写待排序的客户或部门名称。每行：顺序名称=第一项,第二项,第三项"));
            sequenceBody.Children.Add(ToolUi.Note("例如填 阶段=准备,实施,验收，再在上方规则中填 [阶段]。"));
            var sequences=new TextBox{Name="SortSequences",Text=Options.Sequences,AcceptsReturn=true,Height=64,MaxLength=8192,VerticalScrollBarVisibility=ScrollBarVisibility.Auto,Padding=new Thickness(8)};System.Windows.Automation.AutomationProperties.SetName(sequences,"自定义顺序定义（可选）");sequenceBody.Children.Add(sequences);sequences.TextChanged+=(s,e)=>{Options.Sequences=sequences.Text;Invalidate();};
            var natural=ToolUi.Check(left,Options,nameof(Options.Natural),"名称自然排序（忽略规则，数字 2 在 10 前）");natural.Click+=(s,e)=>Invalidate();
            var desc=ToolUi.Check(left,Options,nameof(Options.Descending),"组内降序（规则优先级保持不变）");desc.Click+=(s,e)=>Invalidate();
            var match=ToolUi.Check(left,Options,nameof(Options.MatchCase),"区分英文字母大小写");match.Click+=(s,e)=>Invalidate();
            left.Children.Add(ToolUi.Note("3 · 点击“预览排序”核对右侧结果，再应用。未匹配项排在最后；相同键值保持原序。排序会移动整条记录。"));
            Action updateExample=()=>exampleNote.Text=new[]{"示例结果：客户一处 → 客户二处 → 客户十一处 → 客户一百零一处。名称中的中文数字都可识别。","示例结果：大型客户 → 中型客户 → 小型客户 → 微型客户。","示例结果：国际二部 → 国际十部 → 业务一部 → 业务三部。第一行规则优先。","示例结果：大一 → 大十 → 中二 → 小一 → 微一。先按规模，再按数字。","示例结果：准备 → 实施 → 验收。本示例会同时填写排序规则和自定义顺序。","示例结果：表1 → 表2 → 表10。不使用下方规则。"}[presets.SelectedIndex];
            presets.SelectionChanged+=(s,e)=>updateExample();updateExample();
            sampleButton.Click+=(s,e)=>
            {
                int i=presets.SelectedIndex;
                Options.Natural=i==5;natural.IsChecked=Options.Natural;
                if(i<5)rules.Text=new[]{".*?[中文数字]","[大中小微]","(国际)[中文数字]\n(业务)[中文数字]","[大中小微][中文数字]","[阶段]"}[i];
                sequences.Text=i==4?"阶段=准备,实施,验收":"";sequenceExpander.IsExpanded=i==4;Options.Descending=false;desc.IsChecked=false;Invalidate();Preview();
            };
            Action syncMode=()=>{rules.IsEnabled=sequenceExpander.IsEnabled=!Options.Natural;};natural.Click+=(s,e)=>syncMode();sampleButton.Click+=(s,e)=>syncMode();syncMode();
            Loaded+=(s,e)=>Preview();
        }
        private void Invalidate(){apply.IsEnabled=false;status.Text="规则已改变，请重新预览。";}
        private void Preview()
        {
            apply.IsEnabled=false;
            try{var values=read(Column,Options.HasHeader);var order=CustomSort.Order(values,Options,out int matched);preview.ItemsSource=order.Take(200).Select(i=>new{Position=i+1,Value=values[i]}).ToArray();status.Text=matched==0?"没有名称匹配规则。可点“套用示例”；数字在名称中间时，使用 .*?[中文数字]。":"共 "+values.Count+" 项，匹配 "+matched+" 项；未匹配 "+(values.Count-matched)+" 项置后。核对后应用，工作簿不会自动保存。";apply.IsEnabled=values.Count>1&&matched>0;}
            catch(Exception ex){status.Text=ex.Message;preview.ItemsSource=null;}
        }
    }
}
