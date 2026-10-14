using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class ExcelGateway
    {
        private static XmlReader SheetReader(ZipArchiveEntry entry)=>XmlReader.Create(entry.Open(),new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,CloseInput=true,IgnoreWhitespace=true});
        private static XElement SheetShell(ZipArchiveEntry entry,bool appearanceOnly)
        {
            using(var reader=SheetReader(entry))
            {
                reader.MoveToContent();var root=new XElement(XName.Get(reader.LocalName,reader.NamespaceURI));
                if(reader.MoveToFirstAttribute())do
                {root.Add(new XAttribute(reader.Prefix=="xmlns"?XNamespace.Xmlns+reader.LocalName:reader.Name=="xmlns"?XName.Get("xmlns"):XName.Get(reader.LocalName,reader.NamespaceURI),reader.Value));}while(reader.MoveToNextAttribute());
                reader.MoveToElement();reader.Read();
                while(!reader.EOF && !(reader.NodeType==XmlNodeType.EndElement && reader.Depth==0))
                {
                    if(reader.NodeType!=XmlNodeType.Element){reader.Read();continue;}
                    string name=reader.LocalName;
                    if(name=="sheetData"){root.Add(new XElement(SpreadsheetNs+"sheetData"));reader.Skip();}
                    else if(!appearanceOnly || name=="sheetFormatPr" || name=="cols" || name=="mergeCells")root.Add(XNode.ReadFrom(reader));
                    else reader.Skip();
                }
                return root;
            }
        }
        private static IEnumerable<XElement> SheetRows(ZipArchiveEntry entry,CancellationToken cancel)
        {
            using(var reader=SheetReader(entry))
            {
                while(!reader.EOF)
                {
                    cancel.ThrowIfCancellationRequested();
                    if(reader.NodeType==XmlNodeType.Element && reader.Depth==2 && reader.LocalName=="row" && reader.NamespaceURI==SpreadsheetNs.NamespaceName)
                        yield return (XElement)XNode.ReadFrom(reader);
                    else reader.Read();
                }
            }
        }
        private static void StreamMaskAppearance(ZipArchiveEntry source,ZipArchive output,string targetPart,string directory,CancellationToken cancel)
        {
            var ns=SpreadsheetNs;var target=output.GetEntry(targetPart);var shell=SheetShell(target,false);var appearance=SheetShell(source,true);var data=shell.Element(ns+"sheetData");
            foreach(string name in new[]{"sheetFormatPr","cols","mergeCells"})
            {
                shell.Element(ns+name)?.Remove();var element=appearance.Element(ns+name);if(element==null)continue;
                if(name=="mergeCells")data.AddAfterSelf(new XElement(element));else data.AddBeforeSelf(new XElement(element));
            }
            string temporary=Path.Combine(directory,"appearance-sheet.xml");
            try
            {
                using(var writer=XmlWriter.Create(temporary,new XmlWriterSettings{Encoding=new System.Text.UTF8Encoding(false),Indent=false}))
                {
                    writer.WriteStartDocument();writer.WriteStartElement("worksheet",ns.NamespaceName);
                    foreach(var attr in shell.Attributes())
                    {
                        if(attr.IsNamespaceDeclaration)
                        {if(attr.Name.LocalName=="xmlns")writer.WriteAttributeString("xmlns",attr.Value);else writer.WriteAttributeString("xmlns",attr.Name.LocalName,null,attr.Value);}
                        else writer.WriteAttributeString(shell.GetPrefixOfNamespace(attr.Name.Namespace)??"",attr.Name.LocalName,attr.Name.NamespaceName,attr.Value);
                    }
                    foreach(var element in shell.Elements())
                    {
                        if(element.Name!=ns+"sheetData"){element.WriteTo(writer);continue;}
                        writer.WriteStartElement("sheetData",ns.NamespaceName);
                        using(var from=SheetRows(source,cancel).GetEnumerator())using(var to=SheetRows(target,cancel).GetEnumerator())
                        {
                            bool hasFrom=from.MoveNext(),hasTo=to.MoveNext();
                            while(hasFrom || hasTo)
                            {
                                cancel.ThrowIfCancellationRequested();int sourceRow=hasFrom?(int)from.Current.Attribute("r"):int.MaxValue,targetRow=hasTo?(int)to.Current.Attribute("r"):int.MaxValue;
                                int row=Math.Min(sourceRow,targetRow);var dest=targetRow==row?to.Current:new XElement(ns+"row",new XAttribute("r",row));
                                dest.Attribute("s")?.Remove();foreach(var cell in dest.Elements(ns+"c"))cell.Attribute("s")?.Remove();
                                if(sourceRow==row)
                                {
                                    var original=from.Current;
                                    foreach(string attr in new[]{"ht","customHeight","hidden","s","customFormat","outlineLevel","collapsed"})if(original.Attribute(attr)!=null)dest.SetAttributeValue(attr,(string)original.Attribute(attr));
                                    var cells=dest.Elements(ns+"c").ToDictionary(c=>(string)c.Attribute("r"));
                                    foreach(var cell in original.Elements(ns+"c").Where(c=>c.Attribute("s")!=null))
                                    {
                                        string address=(string)cell.Attribute("r");
                                        if(!cells.TryGetValue(address,out var dc)){dc=new XElement(ns+"c",new XAttribute("r",address));cells.Add(address,dc);}
                                        dc.SetAttributeValue("s",(string)cell.Attribute("s"));
                                    }
                                    dest.ReplaceNodes(cells.Values.OrderBy(c=>ColumnIndex((string)c.Attribute("r"))));hasFrom=from.MoveNext();
                                }
                                dest.WriteTo(writer);if(targetRow==row)hasTo=to.MoveNext();
                            }
                        }
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();writer.WriteEndDocument();
                }
                target.Delete();using(var stream=output.CreateEntry(targetPart,CompressionLevel.Optimal).Open())using(var input=File.OpenRead(temporary))input.CopyTo(stream);
            }
            finally{if(File.Exists(temporary))File.Delete(temporary);}
        }
    }
}
