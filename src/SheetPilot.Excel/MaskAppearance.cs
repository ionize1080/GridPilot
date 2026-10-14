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
        private static void PreserveMaskAppearance(string original,object writerObject,string target,string directory,MaskResult result,CancellationToken cancel,MaskComLifetime lifetime)
        {
            dynamic writer=writerObject,copy=null;
            string xlsx=Path.Combine(directory,"appearance.xlsx");
            try
            {
                // Open only our snapshot, with macro execution/events/links disabled on the owned writer.
                try {copy=lifetime.Hold(lifetime.Hold(writer.Workbooks).Open(original,0,true));}
                catch(System.Runtime.InteropServices.COMException ex){throw new InvalidOperationException("无法读取排版快照，请检查工作簿是否加密或受保护。",ex);}
                copy.SaveAs(xlsx,51);copy.Close(false);lifetime.Release((object)copy);copy=null;
                using(var input=ZipFile.OpenRead(xlsx))using(var output=ZipFile.Open(target,ZipArchiveMode.Update))
                {
                    var ns=SpreadsheetNs;
                    var styles=ReadPart(input,"xl/styles.xml");
                    // Custom display formats may contain identities as literal text.
                    var matcher=new IdentityMatcher(result.IdentityMappings);
                    foreach(var fmt in styles.Descendants(ns+"numFmt"))
                    {
                        var attr=fmt.Attribute("formatCode");if(attr==null || !matcher.Contains(attr.Value,cancel))continue;
                        // Preserve grammar/indices/colors. Sensitive custom formats fall back to General;
                        // arbitrary replacements could inject Excel number-format syntax.
                        attr.Value="General";
                    }
                    foreach(var attr in styles.Descendants(ns+"cellStyle").Attributes("name"))
                        if(matcher.Contains(attr.Value,cancel))attr.Value="Style"+(string)attr.Parent.Attribute("xfId");
                    // Only presentation parts enter the clean workbook, never source cell values/relationships.
                    WritePart(output,"xl/styles.xml",styles);
                    if(input.GetEntry("xl/theme/theme1.xml")!=null && output.GetEntry("xl/theme/theme1.xml")!=null)WritePart(output,"xl/theme/theme1.xml",ReadPart(input,"xl/theme/theme1.xml"));
                    var relationships=ReadPart(input,"xl/_rels/workbook.xml.rels").Root.Elements().ToDictionary(e=>(string)e.Attribute("Id"),e=>(string)e.Attribute("Target"));
                    XNamespace rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    var parts=ReadPart(input,"xl/workbook.xml").Root.Element(ns+"sheets").Elements().Select(e=>relationships[(string)e.Attribute(rel+"id")]).Where(p=>p.Contains("worksheets/")).ToArray();
                    for(int i=0;i<parts.Length;i++)
                    {
                        cancel.ThrowIfCancellationRequested();string part=parts[i].StartsWith("/")?parts[i].TrimStart('/'):"xl/"+parts[i];
                        StreamMaskAppearance(input.GetEntry(part),output,"xl/worksheets/sheet"+(i+1)+".xml",directory,cancel);
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
