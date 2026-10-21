// C# adaptation of excel-mcp-server refs.py and formatting.py (MIT).
// Copyright (c) 2025-2026 Haris Musa. See docs/licenses/EXCEL-MCP-LICENSE.txt.
using System;
using System.Globalization;
using System.Text.RegularExpressions;
namespace SheetPilot.Excel
{
    public sealed class AiRange
    {
        public int Row, Column, LastRow, LastColumn;
        public long Size => (long)(LastRow-Row+1)*(LastColumn-Column+1);
        public static AiRange Parse(string text)
        {
            var match=Regex.Match(text?.Trim()??"", @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}))?$");
            if(!match.Success)throw new InvalidOperationException("请使用完整 A1 地址，例如 B2 或 A1:C10。");
            int c=ColumnNumber(match.Groups[1].Value),r=int.Parse(match.Groups[2].Value);
            int c2=match.Groups[3].Success?ColumnNumber(match.Groups[3].Value):c,r2=match.Groups[4].Success?int.Parse(match.Groups[4].Value):r;
            if(r>1048576||r2>1048576||c>16384||c2>16384)throw new InvalidOperationException("地址超出 Excel 工作表边界 XFD1048576。");
            return new AiRange{Row=Math.Min(r,r2),LastRow=Math.Max(r,r2),Column=Math.Min(c,c2),LastColumn=Math.Max(c,c2)};
        }
        private static int ColumnNumber(string text) { int n=0;foreach(char c in text.ToUpperInvariant())n=n*26+c-'A'+1;return n; }
        public static int Color(string text)
        {
            var match=Regex.Match(text?.Trim()??"",@"^#?([0-9A-Fa-f]{6})$");
            if(!match.Success)throw new InvalidOperationException("颜色需为 #RRGGBB。");
            int rgb=int.Parse(match.Groups[1].Value,NumberStyles.HexNumber);
            return ((rgb>>16)&255)|(rgb&0x00ff00)|((rgb&255)<<16);
        }
    }
}
