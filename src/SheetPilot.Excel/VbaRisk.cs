using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SheetPilot.Excel
{
    // A review gate, not a VBA sandbox. Keep ordinary Excel/file operations out of it.
    public static class VbaRisk
    {
        public static string[] Inspect(string code)
        {
            if (code == null || code.Length > 1000000) throw new ArgumentException("VBA 代码为空或超过 1 MB。");
            var lines = new List<string>();
            foreach (string line in code.Replace("\r", "").Split('\n'))
            {
                bool quoted = false; int end = line.Length;
                for (int i = 0; i < line.Length; i++)
                {
                    if (line[i] == '"') { if (quoted && i + 1 < line.Length && line[i + 1] == '"') { i++; continue; } quoted = !quoted; }
                    if (!quoted && line[i] == '\'') { end = i; break; }
                }
                var clean = line.Substring(0, end);
                if (!Regex.IsMatch(clean, @"^\s*Rem(?:\s|$)", RegexOptions.IgnoreCase)) lines.Add(clean);
            }
            string source = Regex.Replace(string.Join("\n", lines), @"\s+_\s*\n", " ");
            var risks = new List<string>();
            string statements=Regex.Replace(source,@"""(?:[^""]|"""")*""","\"\"");
            if(Regex.IsMatch(statements,@"\b(?:Shell|ShellExecute\w*|WinExec|CreateProcess\w*)\b",RegexOptions.IgnoreCase)||Regex.IsMatch(source,@"\bAlias\s+""(?:ShellExecute\w*|WinExec|CreateProcess\w*)""",RegexOptions.IgnoreCase))risks.Add("shell/process");
            foreach (var rule in new[] {
                Tuple.Create(@"(?:WScript\s*\.\s*Shell|Shell\s*\.\s*Application)", "shell automation"),
                Tuple.Create(@"(?:XMLHTTP|ServerXMLHTTP|WinHttp|InternetOpen|InternetConnect|URLDownloadToFile|URLMon|WinInet|WS2_32|Winsock)", "network API"),
                Tuple.Create(@"\b(?:FollowHyperlink|SendMail|SendForReview)\b", "network send/navigation"),
                Tuple.Create(@"\b(?:CallByName|ExecuteExcel4Macro)\b", "dynamic execution"),
                Tuple.Create(@"\b(?:CreateObject|GetObject)\s*\(\s*(?!""[^""\r\n]+""\s*\))", "dynamic automation target")
            }) if (Regex.IsMatch(source, rule.Item1, RegexOptions.IgnoreCase)) risks.Add(rule.Item2);
            // Concatenated literals commonly obscure shell/network ProgIDs.
            string folded = Regex.Replace(source, "\"\\s*&\\s*\"", "");
            if (folded != source && Regex.IsMatch(folded, @"WScript\.Shell|Shell\.Application|XMLHTTP|WinHttp", RegexOptions.IgnoreCase)) risks.Add("constructed shell/network target");
            return risks.Distinct().ToArray();
        }
    }
}
