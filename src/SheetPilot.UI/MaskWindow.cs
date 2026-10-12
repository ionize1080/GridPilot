using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.Win32;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public sealed class MaskWindow : Window
    {
        private CancellationTokenSource pending;
        private MaskResult previewResult;
        private string previewRules;
        public MaskWindow(NavigationSettings colors,MaskSnapshot source,IReadOnlyList<MaskRule> rules,Func<MaskResult,string,CancellationToken,string> export)
        {
            ToolUi.Theme(this,colors,"关系保持脱敏",1100,790);
            var root=new DockPanel {Margin=new Thickness(24)};Content=root;
            var heading=new StackPanel();DockPanel.SetDock(heading,Dock.Top);root.Children.Add(heading);
            heading.Children.Add(new TextBlock {Text="关系保持脱敏 · 独立数据副本",FontSize=23});
            heading.Children.Add(ToolUi.Note("逐列选择规则。对应列填写相同“逻辑字段”，例如客户名称／借款人名称都填写“客户”。保留列不主动改写；直接引用身份字段的公式会同步变化。包含隐藏表。"));
            heading.Children.Add(ToolUi.Note("结果使用 Sheet1、Sheet2 等中性表名，仅输出单元格数据与受支持公式，不复制样式、批注、图表、连接、宏等对象。编号生成带 TEST 前缀的模拟文本，不是真实信用代码。"));
            heading.Children.Add(ToolUi.Note("金额按正系数缩放：相关金额计算列也应设置同一规则；日期整体平移：不能保证月份、季度或节假日不变。分类、绿色贷款指标等留为 Keep。规则不等于强匿名，交付前仍需确认保留字段。"));
            var bottom=new StackPanel();DockPanel.SetDock(bottom,Dock.Bottom);root.Children.Add(bottom);
            var status=ToolUi.Note("已读取 "+source.Sheets.Count+" 张表、"+source.Cells.Count+" 个有内容单元格。默认全部保留，请选择脱敏字段。");bottom.Children.Add(status);
            var actions=new StackPanel {Orientation=Orientation.Horizontal};bottom.Children.Add(actions);
            var grid=new DataGrid {ItemsSource=rules,AutoGenerateColumns=false,CanUserAddRows=false,Margin=new Thickness(0,12,0,12)};root.Children.Add(grid);
            Action<string,string,int,bool> column=(header,path,width,readOnly)=>grid.Columns.Add(new DataGridTextColumn {Header=header,Binding=new Binding(path) {UpdateSourceTrigger=UpdateSourceTrigger.PropertyChanged},Width=width,IsReadOnly=readOnly});
            column("工作表","SheetName",125,true);column("列","Column",40,true);column("首个值／表头","Header",140,true);
            grid.Columns.Add(new DataGridComboBoxColumn {Header="规则",ItemsSource=new Dictionary<MaskKind,string>{{MaskKind.Keep,"保留原值"},{MaskKind.ChineseName,"中文假名"},{MaskKind.EnglishName,"英文假名"},{MaskKind.Identifier,"模拟编号"},{MaskKind.Amount,"金额缩放"},{MaskKind.Date,"日期平移"}},DisplayMemberPath="Value",SelectedValuePath="Key",SelectedValueBinding=new Binding("Kind") {UpdateSourceTrigger=UpdateSourceTrigger.PropertyChanged},Width=120});
            column("逻辑字段","Domain",130,false);column("起始行","StartRow",65,false);column("金额系数","Factor",80,false);column("日期天数","DayShift",80,false);
            heading.Children.Add(ToolUi.Note("双击规则单元格选择规则。没有表头时，请把起始行改为第一条数据所在行。映射限本次工作簿，不自动猜测别名或跨月复用。"));
            Button preview=null,save=null;
            Func<bool,Task> run=async write=>
            {
                if(pending!=null)return;
                if(!grid.CommitEdit(DataGridEditingUnit.Cell,true) || !grid.CommitEdit(DataGridEditingUnit.Row,true) || Validation.GetHasError(grid)) {status.Text="请修正字段输入。";return;}
                string fingerprint=Newtonsoft.Json.JsonConvert.SerializeObject(rules);
                string path=null;
                if(write) {var dialog=new SaveFileDialog {Filter="Excel 工作簿|*.xlsx",FileName="GridPilot-masked-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+".xlsx",AddExtension=true,OverwritePrompt=true};if(dialog.ShowDialog(this)!=true)return;path=dialog.FileName;}
                pending=new CancellationTokenSource();var token=pending.Token;grid.IsEnabled=false;preview.IsEnabled=save.IsEnabled=false;
                try
                {
                    status.Text="正在分析字段映射与公式关系，可取消…";
                    var result=previewResult!=null && previewRules==fingerprint?previewResult:await Task.Run(()=>MaskPlanner.Build(source,rules,token));token.ThrowIfCancellationRequested();
                    previewResult=result;previewRules=fingerprint;
                    if(write) {status.Text="正在创建独立副本，并由 Excel 重算、重新打开验证…";string saved=export(result,path,token);status.Text="已保存："+saved+"\n"+result.Summary;}
                    else {var originals=source.Cells.ToDictionary(c=>c.Key);var sample=result.Output.Cells.Where(c=>c.Formula==null && !Equals(c.Value,originals[c.Key].Value)).Take(12).Select(c=>new {位置=c.Location,原值=originals[c.Key].Value,替代值=c.Value}); var window=new Window {Title="GridPilot · 本地脱敏预览",Owner=this,Width=760,Height=420,WindowStartupLocation=WindowStartupLocation.CenterOwner,Content=new DataGrid {ItemsSource=sample.ToArray(),IsReadOnly=true,AutoGenerateColumns=true,CanUserAddRows=false}};window.ShowDialog();status.Text=result.Summary+" 尚未保存。规则不变时，导出沿用当前预览映射。";}
                }
                catch(OperationCanceledException) {status.Text="已取消，未交付结果。";}
                catch(Exception ex) {status.Text="已停止："+ex.Message;}
                finally {pending.Dispose();pending=null;grid.IsEnabled=true;preview.IsEnabled=save.IsEnabled=true;}
            };
            preview=ToolUi.Button(actions,"分析与预览",async()=>await run(false));save=ToolUi.Button(actions,"验证并另存新文件",async()=>await run(true));
            ToolUi.Button(actions,"取消分析",()=>pending?.Cancel());ToolUi.Button(actions,"关闭",Close);
            Closing+=(s,e)=> {if(pending!=null){pending.Cancel();e.Cancel=true;}};
        }
    }
}
