using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using SheetPilot.Core;

namespace SheetPilot.UI
{
    public sealed class SegmentationWindow : Window
    {
        private CancellationTokenSource pending;
        public SegmentationWindow(NavigationSettings colors, IReadOnlyList<TextInput> input, string resources, Action<object[,]> output)
        {
            ToolUi.Theme(this,colors,"中英文分词",930,730);
            var root=new DockPanel { Margin=new Thickness(24) }; Content=root;
            var options=new SegmentationOptions();
            var top=new StackPanel(); DockPanel.SetDock(top,Dock.Top); root.Children.Add(top);
            top.Children.Add(new TextBlock { Text="中英文分词",FontSize=24 });
            top.Children.Add(ToolUi.Note("处理当前选区的文本及公式文本结果；保留词序和重复词。空白、数字、错误及合并区域跳过。输出到新工作簿，源数据不改写。"));
            ToolUi.Choice(top,options,"OutputMode","输出形式","每条原文对应一格","横向分列","纵向明细（含词序及字符位置）");
            ToolUi.Check(top,options,"KeepPunctuation","保留标点与符号"); ToolUi.Check(top,options,"Distinct","每条原文内去重（保留首次出现）");
            ToolUi.Text(top,options,"Separator","单格输出分隔符");
            top.Children.Add(ToolUi.Note("专业词与英文短语：每行一条，优先保持完整（最多5000条）。"));
            var terms=new TextBox { Text=options.Terms, AcceptsReturn=true,Height=90,VerticalScrollBarVisibility=ScrollBarVisibility.Auto }; top.Children.Add(terms);
            var import=new StackPanel {Orientation=Orientation.Horizontal,Margin=new Thickness(0,8,0,8)}; top.Children.Add(import);
            ToolUi.Button(import,"导入 UTF-8 词表",()=> { var dialog=new OpenFileDialog { Filter="UTF-8 文本|*.txt" }; if(dialog.ShowDialog(this)==true) { try { if(new System.IO.FileInfo(dialog.FileName).Length>1024*1024) throw new InvalidOperationException("词表超过1MB。"); terms.Text=System.IO.File.ReadAllText(dialog.FileName,new System.Text.UTF8Encoding(false,true)); } catch(Exception ex) { MessageBox.Show(this,ex.Message,Title); } } });
            var bottom=new StackPanel(); DockPanel.SetDock(bottom,Dock.Bottom); root.Children.Add(bottom);
            var status=ToolUi.Note("已读取 "+input.Count+" 条文本；公式使用当前计算值，请先确保源表已完成计算。"); bottom.Children.Add(status);
            var actions=new StackPanel { Orientation=Orientation.Horizontal }; bottom.Children.Add(actions);
            var grid=new DataGrid { AutoGenerateColumns=false,IsReadOnly=true,CanUserAddRows=false,Margin=new Thickness(0,10,0,10) };
            grid.Columns.Add(new DataGridTextColumn { Header="位置",Binding=new System.Windows.Data.Binding("Source.Address"),Width=90 });
            grid.Columns.Add(new DataGridTextColumn { Header="原文",Binding=new System.Windows.Data.Binding("Source.Text"),MinWidth=180,Width=new DataGridLength(1,DataGridLengthUnitType.Star) });
            grid.Columns.Add(new DataGridTextColumn { Header="分词结果",Binding=new System.Windows.Data.Binding("Preview"),MinWidth=300,Width=new DataGridLength(2,DataGridLengthUnitType.Star) }); root.Children.Add(grid);
            Button preview=null,execute=null;
            Func<bool,Task> run=async export=>
            {
                if(pending!=null) return; pending=new CancellationTokenSource(); var token=pending.Token;
                preview.IsEnabled=execute.IsEnabled=false; top.IsEnabled=false;
                try
                {
                    options.Terms=terms.Text; status.Text="正在分词，可取消…";
                    var rows=await Task.Run(()=> { var segmenter=new TextSegmenter(resources); var result=new List<TextOutput>(); foreach(var item in (export?input:input.Take(30))) { token.ThrowIfCancellationRequested(); result.Add(new TextOutput { Source=item,Words=segmenter.Split(item.Text,options,token) }); } return result; });
                    token.ThrowIfCancellationRequested(); grid.ItemsSource=rows.Take(30).ToArray();
                    if(export) { var values=TextBatch.Layout(rows,options); output(values); status.Text="完成："+rows.Count+" 条文本，结果已写入新工作簿。";Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Background,new Action(Close)); }
                    else status.Text="已预览前 "+rows.Count+" 条；生成结果时处理全部 "+input.Count+" 条。";
                }
                catch(OperationCanceledException) { status.Text="已取消，未写入结果。"; }
                catch(Exception ex) { status.Text="未完成："+ex.Message; }
                finally { pending.Dispose(); pending=null; preview.IsEnabled=execute.IsEnabled=true; top.IsEnabled=true; }
            };
            preview=ToolUi.Button(actions,"预览",async()=>await run(false)); execute=ToolUi.Button(actions,"生成新工作簿",async()=>await run(true));
            ToolUi.Button(actions,"取消计算",()=>pending?.Cancel()); ToolUi.Button(actions,"关闭",Close);
            Closing+=(s,e)=> { if(pending!=null) { pending.Cancel(); e.Cancel=true; status.Text="正在取消，结束后可关闭。"; } };
        }
    }
}
