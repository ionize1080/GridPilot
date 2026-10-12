using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using JiebaNet.Segmenter;

namespace SheetPilot.Core
{
    public sealed class WordToken
    {
        public string Text { get; set; }
        public int Offset { get; set; }
    }
    public sealed class SegmentationOptions
    {
        public int OutputMode { get; set; }
        public bool KeepPunctuation { get; set; }
        public bool Distinct { get; set; }
        public string Separator { get; set; } = " / ";
        public string Terms { get; set; } = "绿色贷款\n统一社会信用代码\n贷款余额\nIFRS 9\nMAS 610";
    }
    public sealed class TextSegmenter
    {
        private readonly JiebaSegmenter engine;
        private static readonly object Gate = new object();
        private static JiebaSegmenter shared;
        private static readonly Regex Protected = new Regex(@"\G(?:https?://[^\s<>\u3000-\u303f\uff00-\uffef]+|[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}|\d{4}[-/]\d{1,2}[-/]\d{1,2}|[0-9０-９]+[A-Za-zＡ-Ｚａ-ｚ][A-Za-zＡ-Ｚａ-ｚ0-9０-９]*(?:[-'’._][A-Za-z0-9]+)*|[+-]?\d+(?:,\d{3})*(?:\.\d+)?[%％]?|[A-Za-zＡ-Ｚａ-ｚ0-9０-９]+(?:[-'’._][A-Za-z0-9]+)*)", RegexOptions.CultureInvariant, TimeSpan.FromMilliseconds(200));
        public TextSegmenter(string resourceDirectory)
        {
            lock(Gate)
            {
                if(shared == null)
                {
                    if(!File.Exists(Path.Combine(resourceDirectory,"dict.txt"))) throw new InvalidOperationException("分词词典缺失，请完整解压发布包中的 Resources 文件夹。");
                    ConfigManager.ConfigFileBaseDir=Path.GetFullPath(resourceDirectory);
                    shared=new JiebaSegmenter();
                }
                engine=shared;
            }
        }
        public IReadOnlyList<WordToken> Split(string text, SegmentationOptions options, CancellationToken cancel)
        {
            if(text == null) return new WordToken[0];
            var terms=(options.Terms??"").Split(new[]{'\r','\n'},StringSplitOptions.RemoveEmptyEntries)
                .Select(t=>t.Trim()).Where(t=>t.Length>0).Distinct(StringComparer.Ordinal).OrderByDescending(t=>t.Length).ToArray();
            if(terms.Length>5000 || terms.Any(t=>t.Length>200)) throw new InvalidOperationException("专业词最多 5000 条，每条不超过 200 字。");
            var result=new List<WordToken>(); var seen=new HashSet<string>(StringComparer.Ordinal);
            Action<string,int> add=(value,offset)=> { if(string.IsNullOrWhiteSpace(value)) return; if(options.Distinct && !seen.Add(value)) return; result.Add(new WordToken { Text=value, Offset=offset }); };
            for(int i=0;i<text.Length;)
            {
                cancel.ThrowIfCancellationRequested();
                string term=terms.FirstOrDefault(t=>i+t.Length<=text.Length && string.CompareOrdinal(text,i,t,0,t.Length)==0);
                if(term!=null) { add(term,i); i+=term.Length; continue; }
                if(char.IsWhiteSpace(text,i)) { i++; continue; }
                var match=Protected.Match(text,i);
                if(match.Success) { add(match.Value,i); i+=match.Length; continue; }
                if(IsHan(text[i]))
                {
                    int start=i++;
                    while(i<text.Length && IsHan(text[i]) && !terms.Any(t=>i+t.Length<=text.Length && string.CompareOrdinal(text,i,t,0,t.Length)==0)) i++;
                    string chunk=text.Substring(start,i-start); string[] words;
                    lock(Gate) words=engine.Cut(chunk,false,true).ToArray();
                    int offset=start;
                    foreach(var word in words) { cancel.ThrowIfCancellationRequested(); add(word,offset); offset+=word.Length; }
                }
                else
                {
                    int length=char.IsHighSurrogate(text[i]) && i+1<text.Length && char.IsLowSurrogate(text[i+1])?2:1;
                    if(options.KeepPunctuation || char.IsLetter(text,i)) add(text.Substring(i,length),i);
                    i+=length;
                }
            }
            return result;
        }
        private static bool IsHan(char c) => c>='\u3400' && c<='\u9fff' || c>='\uf900' && c<='\ufaff';
    }
}
