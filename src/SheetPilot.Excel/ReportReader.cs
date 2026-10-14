using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace SheetPilot.Excel
{
    // Read cached XLSX values without opening Excel, running macros, refreshing links or saving the source.
    public static class ReportReader
    {
        private static readonly XNamespace Ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static XDocument Xml(ZipArchive zip,string name)
        {
            var entry=zip.GetEntry(name)??throw new InvalidOperationException("报表结构不完整："+name);
            using(var stream=entry.Open()) using(var reader=XmlReader.Create(stream,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=160000000}))return XDocument.Load(reader);
        }
        public static object Read(string path,string sheetName,int maxRows,CancellationToken token)
        {
            if(!new[]{".xlsx",".xlsm"}.Contains(Path.GetExtension(path).ToLowerInvariant()))throw new InvalidOperationException("第一版支持 .xlsx / .xlsm；请先将旧版 .xls 另存为 .xlsx。");
            using(var stream=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite|FileShare.Delete)) using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))
            {
                if(zip.Entries.Sum(e=>e.Length)>512L*1024*1024)throw new InvalidOperationException("报表解压后超过 512 MB，请拆分文件。");
                var book=Xml(zip,"xl/workbook.xml");var sheets=book.Descendants(Ns+"sheet").ToArray();
                var chosen=string.IsNullOrEmpty(sheetName)?sheets.FirstOrDefault():sheets.FirstOrDefault(s=>(string)s.Attribute("name")==sheetName);
                if(chosen==null)throw new InvalidOperationException("未找到指定工作表，请检查文件结构。");
                string rid=(string)chosen.Attribute(XName.Get("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
                var rel=Xml(zip,"xl/_rels/workbook.xml.rels").Root.Elements().FirstOrDefault(r=>(string)r.Attribute("Id")==rid);
                string target=(string)rel?.Attribute("Target");
                if(string.IsNullOrEmpty(target)||(string)rel.Attribute("TargetMode")=="External")throw new InvalidOperationException("不支持外部工作表。");
                string entryName=target.StartsWith("/")?target.TrimStart('/'):("xl/"+target);
                if(entryName.Contains(".."))throw new InvalidOperationException("工作表路径无效。");
                var strings=zip.GetEntry("xl/sharedStrings.xml")==null?new string[0]:Xml(zip,"xl/sharedStrings.xml").Root.Elements(Ns+"si").Select(si=>string.Concat(si.Descendants(Ns+"t").Select(t=>t.Value))).ToArray();
                var formats=new Dictionary<int,string>();var styles=new List<int>();
                if(zip.GetEntry("xl/styles.xml")!=null){var style=Xml(zip,"xl/styles.xml");foreach(var f in style.Descendants(Ns+"numFmt"))formats[(int)f.Attribute("numFmtId")]=(string)f.Attribute("formatCode");foreach(var f in style.Root.Element(Ns+"cellXfs")?.Elements()??Enumerable.Empty<XElement>())styles.Add((int?)f.Attribute("numFmtId")??0);}
                var rows=new List<object>();int cells=0;
                using(var data=zip.GetEntry(entryName)?.Open()??throw new InvalidOperationException("缺少工作表数据。"))
                using(var reader=XmlReader.Create(data,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=320000000}))
                {
                    while(reader.Read())
                    {
                        token.ThrowIfCancellationRequested();if((rows.Count&1023)==0&&GC.GetTotalMemory(false)>(Environment.Is64BitProcess?1024L*1024*1024:300L*1024*1024))throw new InvalidOperationException("读取接近进程内存限制，请拆分报表或使用 64 位 Excel。");if(reader.NodeType!=XmlNodeType.Element||reader.LocalName!="row")continue;
                        XElement row;using(var subtree=reader.ReadSubtree())row=XElement.Load(subtree);
                        int number=(int?)row.Attribute("r")??rows.Count+1;
                        var values=new List<object>();var numeric=new List<int>();
                        foreach(var c in row.Elements(Ns+"c"))
                        {
                            if(++cells>2000000)throw new InvalidOperationException("单个报表超过 200 万个单元格，请拆分后重试。");
                            string address=(string)c.Attribute("r")??"";int column=0;foreach(char ch in address){if(ch<'A'||ch>'Z')break;column=column*26+ch-'A'+1;}
                            if(column<1||column>256)throw new InvalidOperationException("报表最多支持 256 列。");
                            while(values.Count<column)values.Add(null);
                            string type=(string)c.Attribute("t"),raw=(string)c.Element(Ns+"v");object value=raw;if(raw!=null&&(type==null||type=="n"))numeric.Add(column-1);
                            if(type=="s")value=raw==null?null:strings[int.Parse(raw,CultureInfo.InvariantCulture)];
                            else if(type=="inlineStr")value=string.Concat(c.Descendants(Ns+"t").Select(t=>t.Value));
                            else if(type=="e")throw new InvalidOperationException("第 "+number+" 行含 Excel 错误值，请在源表修正。");
                            else if(type=="b")value=raw=="1"?"TRUE":"FALSE";
                            else if(type!="str" && raw!=null)
                            {
                                int index=(int?)c.Attribute("s")??0;string format=null;if(index<styles.Count)formats.TryGetValue(styles[index],out format);
                                // Keep every numeric lexeme as text; configured amount/date columns are converted later.
                                if(format!=null && format.Length<=32 && format.All(ch=>ch=='0') && raw.All(char.IsDigit))value=raw.PadLeft(format.Length,'0');
                            }
                            if(c.Element(Ns+"f")!=null&&raw==null)throw new InvalidOperationException("第 "+number+" 行公式没有缓存值，请用 Excel 计算并保存后扫描。");
                            values[column-1]=value;
                        }
                        if(values.Any(v=>v!=null&&!Equals(v,"")))rows.Add(new {row=number,values,numeric});
                        if(maxRows>0&&rows.Count>=maxRows)break;
                    }
                }
                return new {sheet=(string)chosen.Attribute("name"),sheets=sheets.Select(s=>(string)s.Attribute("name")).ToArray(),date1904=new[]{"1","true"}.Contains((string)book.Root.Element(Ns+"workbookPr")?.Attribute("date1904")),rows};
            }
        }
    }
}
