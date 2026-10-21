using System;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    // One bridge owns its snapshots. All methods run on Excel's macro thread.
    public sealed partial class AiOperations : IDisposable
    {
        private readonly Func<object> application;
        private readonly Dictionary<string, Snapshot> snapshots = new Dictionary<string, Snapshot>();
        private readonly Queue<string> snapshotOrder = new Queue<string>();
        private readonly Dictionary<string, Proposal> proposals = new Dictionary<string, Proposal>();
        private Proposal[] undo;
        private sealed class Cell
        {
            public string sheet, address, formula;
            public object value;
            public int row, column;
            public bool error;
        }
        private sealed class Snapshot
        {
            public string id, scope, bookName;
            public object book;
            private readonly FileStream spool = new FileStream(Path.Combine(Path.GetTempPath(),"GridPilot-ai-"+Guid.NewGuid().ToString("N")),FileMode.CreateNew,FileAccess.ReadWrite,FileShare.None,65536,FileOptions.DeleteOnClose);
            public long count;
            public IEnumerable<Cell> cells { get { spool.Flush();spool.Position=0;using(var reader=new StreamReader(spool,Encoding.UTF8,false,65536,true)) { string line;while((line=reader.ReadLine())!=null)yield return JsonConvert.DeserializeObject<Cell>(line); } } }
            public void Add(Cell cell) { var bytes=Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(cell)+"\n");spool.Write(bytes,0,bytes.Length);count++; }
            public void Dispose() { spool.Dispose();Release(book); }
            public readonly List<object> ranges = new List<object>();
            public readonly HashSet<string> sheets = new HashSet<string>();
            public long scanned;
        }
        private sealed class Proposal
        {
            public string id, snapshot, sheet, address, formula;
            public string kind, chartName;
            public JObject style, beforeStyle, afterStyle;
            public Edit[] edits;
            public object sourceFormula;
            public JObject options, chartState;
        }
        public AiOperations(Func<object> application) { this.application = application; }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        private static object At(object value, int row, int column)
        {
            var array = value as Array;
            return array == null ? value : array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1));
        }
        private static string Address(int row, int column)
        {
            string letters = "";
            for (int n = column; n > 0; n = (n - 1) / 26) letters = (char)('A' + (n - 1) % 26) + letters;
            return letters + row;
        }
        public object Context()
        {
            dynamic app = application(); object book = null, sheet = null, selection = null;
            try
            {
                book = app.ActiveWorkbook; sheet = app.ActiveSheet; selection = app.Selection;
                if (book == null || sheet == null) return new { available = false };
                return new { available = true, book = (string)((dynamic)book).Name, sheet = (string)((dynamic)sheet).Name,
                    selection = (string)((dynamic)selection).Address[false, false], excelVersion = (string)app.Version };
            }
            catch { return new { available = false }; }
            finally { Release(selection); Release(sheet); Release(book); }
        }
        public object Capture(string scope)
        {
            if (scope != "selection" && scope != "sheet" && scope != "workbook")
                throw new InvalidOperationException("请选择选区、当前工作表或当前工作簿。");
            dynamic app = application(); object book = app.ActiveWorkbook;
            if (book == null) throw new InvalidOperationException("请先打开工作簿。");
            var snapshot = new Snapshot { id = Guid.NewGuid().ToString("N"), scope = scope, book = book, bookName = (string)((dynamic)book).Name };
            object active = null, sheets = null;
            try
            {
                if (scope == "workbook")
                {
                    sheets = ((dynamic)book).Worksheets;
                    for (int i = 1; i <= (int)((dynamic)sheets).Count; i++)
                    {
                        object sheet = ((dynamic)sheets).Item[i];
                        try { if ((int)((dynamic)sheet).Visible == -1) CaptureSheet(snapshot, sheet, null); }
                        finally { Release(sheet); }
                    }
                }
                else
                {
                    active = app.ActiveSheet;
                    if (Microsoft.VisualBasic.Information.TypeName(active).TrimStart('_') != "Worksheet")
                        throw new InvalidOperationException("请切换到普通工作表。");
                    object selection = scope == "selection" ? app.Selection : null;
                    try { CaptureSheet(snapshot, active, selection); } finally { Release(selection); }
                }
                // Expire before retaining new COM references. Old chat messages remain readable,
                // but tools must ask for a fresh snapshot rather than resolve by workbook name.
                if (snapshots.Count >= 8)
                {
                    string oldest = snapshotOrder.Dequeue();
                    snapshots[oldest].Dispose(); snapshots.Remove(oldest);
                    foreach (var key in proposals.Where(p => p.Value.snapshot == oldest).Select(p => p.Key).ToArray()) proposals.Remove(key);
                    if (undo?.Any(p=>p.snapshot==oldest)==true) undo = null;
                }
                snapshots.Add(snapshot.id, snapshot);
                snapshotOrder.Enqueue(snapshot.id);
                return new { snapshotId = snapshot.id, book = snapshot.bookName, scope, capturedAt = DateTime.UtcNow.ToString("o"),
                    excelVersion = (string)app.Version, ranges = snapshot.ranges, scannedCells = snapshot.scanned,
                    nonEmptyCells = snapshot.count, hiddenCellsIncluded = true, hiddenSheetsIncluded = false,
                    sample = snapshot.cells.Take(120).ToArray(), sampleTruncated = snapshot.count > 120,
                    statistics = Statistics(snapshot.cells), note = "数据是发送时的快照；查找和统计覆盖全部已扫描范围，sample 只是前120个非空单元格。空白单元格不列在cells中。" };
            }
            catch { snapshot.Dispose(); throw; }
            finally { Release(active); Release(sheets); }
        }
        private void CaptureSheet(Snapshot snapshot, dynamic sheet, object selection)
        {
            object used = null, areas = null, intersection = null;
            try
            {
                used = sheet.UsedRange;
                if (selection != null)
                {
                    dynamic app = application();
                    intersection = app.Intersect(selection, used);
                    if (intersection == null)
                    {
                        // Preserve a blank destination as context without scanning whole rows.
                        object first = ((dynamic)selection).Cells[1, 1];
                        try { CaptureBlock(snapshot, sheet, first); } finally { Release(first); }
                        return;
                    }
                }
                areas = ((dynamic)(intersection ?? used)).Areas;
                for (int i = 1; i <= (int)((dynamic)areas).Count; i++)
                {
                    object block = ((dynamic)areas).Item[i];
                    try { CaptureBlock(snapshot, sheet, block); } finally { Release(block); }
                }
            }
            finally { Release(areas); Release(intersection); Release(used); }
        }
        private void CaptureBlock(Snapshot snapshot, dynamic sheet, dynamic range)
        {
            object rows = null, columns = null;
            try
            {
                rows = range.Rows; columns = range.Columns;
                int nr = ((dynamic)rows).Count, nc = ((dynamic)columns).Count;
                long count = (long)nr * nc;
                snapshot.scanned += count;
                string name = sheet.Name; int startRow = range.Row, startColumn = range.Column;
                snapshot.sheets.Add(name);
                snapshot.ranges.Add(new { sheet = name, address = (string)range.Address[false, false], rows = nr, columns = nc });
                // Sparse areas avoid iterating 17 billion blanks when the last cell is used.
                foreach(int cellType in new[]{2,-4123})
                {
                    object populated=null,clipped=null,parts=null;
                    try
                    {
                        try { populated=range.SpecialCells(cellType); }
                        catch(COMException)
                        {
                            // SpecialCells also fails for some highly fragmented ranges.
                            // Enumerate matching cells rather than silently losing that data.
                            if(cellType==-4123 && Equals(range.HasFormula,false))continue;
                            if(cellType==2 && Equals(range.HasFormula,true))continue;
                            CaptureByFind(snapshot,sheet,range,cellType);continue;
                        }
                        dynamic app=application();clipped=app.Intersect(range,populated);if(clipped==null)continue;
                        parts=((dynamic)clipped).Areas;
                        for(int i=1;i<=(int)((dynamic)parts).Count;i++)
                        {
                            object part=((dynamic)parts).Item[i],pr=null,pc=null;
                            try
                            {
                                dynamic block=part;pr=block.Rows;pc=block.Columns;
                                int height=((dynamic)pr).Count,width=((dynamic)pc).Count,top=block.Row,left=block.Column;
                                int chunkRows=Math.Max(1,16384/width);
                                for(int offset=0;offset<height;offset+=chunkRows)
                                {
                                    int take=Math.Min(chunkRows,height-offset);object chunk=null;
                                    try
                                    {
                                        chunk=sheet.Range[Address(top+offset,left)+":"+Address(top+offset+take-1,left+width-1)];
                                        object values=((dynamic)chunk).Value2,formulas=((dynamic)chunk).Formula;
                                        for(int r=0;r<take;r++)for(int c=0;c<width;c++)
                                        {
                                            object value=At(values,r,c),formula=At(formulas,r,c);
                                            if(value==null&&(formula==null||Equals(formula,"")))continue;
                                            bool isError=value is ErrorWrapper||value is int code&&(code&unchecked((int)0xFFFF0000))==unchecked((int)0x800A0000);
                                            snapshot.Add(new Cell{sheet=name,row=top+offset+r,column=left+c,error=isError,address=Address(top+offset+r,left+c),value=isError?ErrorText(value):value,formula=cellType==-4123?Convert.ToString(formula):null});
                                        }
                                    }finally{Release(chunk);}
                                }
                            }finally{Release(pc);Release(pr);Release(part);}
                        }
                    }finally{Release(parts);Release(clipped);Release(populated);}
                }
            }
            finally { Release(rows); Release(columns); }
        }
        private void CaptureByFind(Snapshot snapshot,dynamic sheet,dynamic range,int cellType)
        {
            object cell=null;
            try
            {
                cell=range.Find(What:"*",LookIn:-4123,LookAt:2,SearchOrder:1,SearchDirection:1,MatchCase:false,SearchFormat:false);
                string first=null;
                while(cell!=null)
                {
                    dynamic found=cell;string address=found.Address[false,false];
                    if(first==address)break;if(first==null)first=address;
                    bool hasFormula=found.HasFormula;
                    if(hasFormula==(cellType==-4123))
                    {
                        object value=found.Value2;bool error=value is ErrorWrapper||value is int code&&(code&unchecked((int)0xFFFF0000))==unchecked((int)0x800A0000);
                        snapshot.Add(new Cell{sheet=(string)sheet.Name,address=address,row=(int)found.Row,column=(int)found.Column,value=error?ErrorText(value):value,error=error,formula=hasFormula?(string)found.Formula:null});
                    }
                    object next=range.FindNext(cell);Release(cell);cell=next;
                }
            }finally{Release(cell);}
        }
        private static string ErrorText(object value)
        {
            int code = value is ErrorWrapper wrapped ? wrapped.ErrorCode : (int)value;
            switch (code & 0xFFFF)
            {
                case 2000: return "#NULL!"; case 2007: return "#DIV/0!"; case 2015: return "#VALUE!";
                case 2023: return "#REF!"; case 2029: return "#NAME?"; case 2036: return "#NUM!"; case 2042: return "#N/A";
                default: return "#ExcelError(" + (code & 0xFFFF) + ")";
            }
        }
        private Snapshot Get(string id)
        {
            if (id == null || !snapshots.TryGetValue(id, out var snapshot)) throw new InvalidOperationException("上下文已过期，请重新发送问题以获取快照。");
            object books = null;
            try
            {
                dynamic app = application(); books = app.Workbooks; bool found = false;
                for (int i = 1; i <= (int)((dynamic)books).Count; i++)
                {
                    object candidate = ((dynamic)books).Item[i];
                    try { if (ReferenceEquals(candidate, snapshot.book) || candidate.Equals(snapshot.book)) found = true; } finally { Release(candidate); }
                }
                if (!found) throw new InvalidOperationException("原工作簿已关闭；同名工作簿不会替代原目标，请重新获取上下文。");
                return snapshot;
            }
            finally { Release(books); }
        }
        private static object Statistics(IEnumerable<Cell> input)
        {
            long count=0,numeric=0,errors=0,formulas=0;double sum=0,min=double.PositiveInfinity,max=double.NegativeInfinity;
            foreach(var cell in input)
            {
                count++;if(cell.error)errors++;if(cell.formula!=null)formulas++;
                if(!cell.error && (cell.value is double||cell.value is long||cell.value is int||cell.value is float||cell.value is decimal))
                { double value=Convert.ToDouble(cell.value);numeric++;sum+=value;min=Math.Min(min,value);max=Math.Max(max,value); }
            }
            return new {nonEmpty=count,numericCount=numeric,errorCount=errors,formulaCount=formulas,sum,average=numeric==0?(double?)null:sum/numeric,min=numeric==0?(double?)null:min,max=numeric==0?(double?)null:max};
        }
        public object Read(JObject request)
        {
            var snapshot = Get((string)request["snapshotId"]);
            string sheet = (string)request["sheet"], query = (string)request["query"] ?? "";
            int column = (int?)request["column"] ?? 0;
            var cells = snapshot.cells.Where(c => (string.IsNullOrEmpty(sheet) || c.sheet == sheet) && (column == 0 || c.column == column));
            string operation = (string)request["operation"];
            if (operation == "statistics") return new { snapshotId = snapshot.id, book = snapshot.bookName, sheet, column, statistics = Statistics(cells) };
            if (operation == "find")
            {
                string comparison = (string)request["comparison"] ?? "contains"; double number = (double?)request["number"] ?? 0;
                if (!new[] { "contains", "equals", "greater", "less", "formula" }.Contains(comparison)) throw new InvalidOperationException("不支持的查找条件。");
                cells = cells.Where(c => comparison == "formula" ? (c.formula ?? "").IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 && c.formula != null :
                    comparison == "greater" || comparison == "less" ? (c.value is double||c.value is long) && (comparison == "greater" ? Convert.ToDouble(c.value) > number : Convert.ToDouble(c.value) < number) :
                    comparison == "equals" ? string.Equals(Convert.ToString(c.value), query, StringComparison.OrdinalIgnoreCase) :
                    Convert.ToString(c.value).IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            else if (operation != "read") throw new InvalidOperationException("不支持的读取操作。");
            long offset=Math.Max(0,(long?)request["offset"]??0),total=0;var page=new List<Cell>();
            foreach(var cell in cells){if(total>=offset&&page.Count<200)page.Add(cell);total++;}
            return new {snapshotId=snapshot.id,total,offset,cells=page,hasMore=offset+page.Count<total};
        }
        public static void ValidateFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith("=") || formula.Length > 8192 ||
                formula.IndexOfAny(new[] { '\r', '\n', '\0', '|' }) >= 0 ||
                Regex.IsMatch(formula,@"\[[^\]]*\.(?:xlsx?|xlsm|xlsb|csv)\]") ||
                Regex.IsMatch(formula, @"(?i)(?:WEBSERVICE|HYPERLINK|RTD|CALL|REGISTER(?:\.ID)?|EXEC|EVALUATE)\s*\("))
                throw new InvalidOperationException("公式无效，或包含外部链接、执行/网络函数。");
        }
        private static void ValidateAddress(string address)
        {
            if(AiRange.Parse(address).Size!=1)throw new InvalidOperationException("请指定单个目标单元格。");
        }
        public object Navigate(string id, string sheetName, string address)
        {
            var snapshot = Get(id); ValidateAddress(address);
            if (!snapshot.sheets.Contains(sheetName)) throw new InvalidOperationException("引用不在已授权的工作表中。");
            object sheets = null, sheet = null, range = null;
            try
            {
                sheets = ((dynamic)snapshot.book).Worksheets; sheet = ((dynamic)sheets).Item[sheetName];
                if ((int)((dynamic)sheet).Visible != -1) throw new InvalidOperationException("工作表已隐藏。");
                range = ((dynamic)sheet).Range[address]; dynamic app = application(); app.Goto(range, true);
                return new { located = true, sheet = sheetName, address };
            }
            finally { Release(range); Release(sheet); Release(sheets); }
        }
        public void Dispose() { foreach (var snapshot in snapshots.Values) snapshot.Dispose(); snapshots.Clear(); snapshotOrder.Clear(); proposals.Clear(); undo = null; }
    }
}
