using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using SheetPilot.Core;

namespace SheetPilot.Excel
{
    public sealed partial class ExcelGateway
    {
        private static readonly XNamespace SpreadsheetNs="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static XDocument ReadPart(ZipArchive archive,string part)
        {
            using(var stream=archive.GetEntry(part).Open())
            using(var reader=XmlReader.Create(stream,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))return XDocument.Load(reader);
        }
        private static void WritePart(ZipArchive archive,string part,XDocument document)
        {archive.GetEntry(part)?.Delete();using(var stream=archive.CreateEntry(part,CompressionLevel.Optimal).Open())document.Save(stream);}
        private static void PreserveMaskAppearance(object sourceObject,object writerObject,string target,string directory,MaskResult result,CancellationToken cancel,MaskComLifetime lifetime)
        {
            dynamic source=sourceObject,writer=writerObject,copy=null;
            string extension=Path.GetExtension((string)source.Name);
            if(string.IsNullOrEmpty(extension))extension=".xlsx";
            string original=Path.Combine(directory,"appearance-source"+extension),xlsx=Path.Combine(directory,"appearance.xlsx");
            try
            {
                source.SaveCopyAs(original);
                // Open only our snapshot, with macro execution/events/links disabled on the owned writer.
                try {copy=lifetime.Hold(lifetime.Hold(writer.Workbooks).Open(original,0,true));}
                catch(System.Runtime.InteropServices.COMException ex){throw new InvalidOperationException("无法读取排版快照，请检查工作簿是否加密或受保护。",ex);}
                copy.SaveAs(xlsx,51);copy.Close(false);copy=null;
                using(var input=ZipFile.OpenRead(xlsx))using(var output=ZipFile.Open(target,ZipArchiveMode.Update))
                {
                    var ns=SpreadsheetNs;
                    var styles=ReadPart(input,"xl/styles.xml");
                    // Custom display formats may contain identities as literal text.
                    var matcher=new IdentityMatcher(result.IdentityMappings);
                    foreach(var attr in styles.Descendants().Attributes())if(matcher.Contains(attr.Value,cancel))attr.Value=matcher.Replace(attr.Value,cancel);
                    // Only presentation parts enter the clean workbook, never source cell values/relationships.
                    WritePart(output,"xl/styles.xml",styles);
                    if(input.GetEntry("xl/theme/theme1.xml")!=null && output.GetEntry("xl/theme/theme1.xml")!=null)WritePart(output,"xl/theme/theme1.xml",ReadPart(input,"xl/theme/theme1.xml"));
                    var relationships=ReadPart(input,"xl/_rels/workbook.xml.rels").Root.Elements().ToDictionary(e=>(string)e.Attribute("Id"),e=>(string)e.Attribute("Target"));
                    XNamespace rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    var parts=ReadPart(input,"xl/workbook.xml").Root.Element(ns+"sheets").Elements().Select(e=>relationships[(string)e.Attribute(rel+"id")]).Where(p=>p.Contains("worksheets/")).ToArray();
                    for(int i=0;i<parts.Length;i++)
                    {
                        cancel.ThrowIfCancellationRequested();string part=parts[i].StartsWith("/")?parts[i].TrimStart('/'):"xl/"+parts[i];
                        var from=ReadPart(input,part);string toPart="xl/worksheets/sheet"+(i+1)+".xml";var to=ReadPart(output,toPart);
                        // The new workbook may have generated quote-prefix/date styles whose indices
                        // do not exist in the source style table. Reset all references before copying.
                        foreach(var attr in to.Descendants(ns+"c").Attributes("s").ToArray())attr.Remove();
                        foreach(var attr in to.Descendants(ns+"row").Attributes("s").ToArray())attr.Remove();
                        var targetData=to.Root.Element(ns+"sheetData");var rows=targetData.Elements(ns+"row").ToDictionary(e=>(string)e.Attribute("r"));
                        foreach(var row in from.Root.Element(ns+"sheetData").Elements(ns+"row"))
                        {
                            string id=(string)row.Attribute("r");
                            if(!rows.TryGetValue(id,out var dest)){dest=new XElement(ns+"row",new XAttribute("r",id));rows.Add(id,dest);}
                            foreach(string attr in new[]{"ht","customHeight","hidden","s","customFormat","outlineLevel","collapsed"})if(row.Attribute(attr)!=null)dest.SetAttributeValue(attr,(string)row.Attribute(attr));
                            var cells=dest.Elements(ns+"c").ToDictionary(e=>(string)e.Attribute("r"));
                            foreach(var cell in row.Elements(ns+"c").Where(e=>e.Attribute("s")!=null))
                            {
                                string address=(string)cell.Attribute("r");
                                if(!cells.TryGetValue(address,out var dc)){dc=new XElement(ns+"c",new XAttribute("r",address));cells.Add(address,dc);}
                                dc.SetAttributeValue("s",(string)cell.Attribute("s"));
                            }
                            dest.ReplaceNodes(cells.Values.OrderBy(c=>ColumnIndex((string)c.Attribute("r"))));
                        }
                        targetData.ReplaceNodes(rows.Values.OrderBy(r=>(int)r.Attribute("r")));
                        foreach(string element in new[]{"sheetFormatPr","cols","mergeCells"})
                        {
                            to.Root.Element(ns+element)?.Remove();var content=from.Root.Element(ns+element);if(content==null)continue;
                            if(element=="mergeCells")targetData.AddAfterSelf(new XElement(content));else targetData.AddBeforeSelf(new XElement(content));
                        }
                        WritePart(output,toPart,to);
                    }
                }
            }
            finally
            {
                try {if(copy!=null)copy.Close(false);}catch{}
                foreach(string path in new[]{original,xlsx})try {if(File.Exists(path))File.Delete(path);}catch(IOException){}
            }
        }
        private static int ColumnIndex(string address){int column=0;foreach(char c in address){if(c<'A'||c>'Z')break;column=column*26+c-'A'+1;}return column;}
    }
}
