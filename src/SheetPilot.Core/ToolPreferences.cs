using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace SheetPilot.Core
{
    [DataContract]
    public sealed class RenameOptions
    {
        public RenameOptions Clone() => (RenameOptions)MemberwiseClone();
        [DataMember] public bool AllWorkbooks { get; set; }
        [DataMember] public bool IncludeHidden { get; set; } = true;
        [DataMember] public bool IncludeCharts { get; set; } = true;
        [DataMember] public string Find { get; set; } = "";
        [DataMember] public string Replacement { get; set; } = "";
        [DataMember] public bool Regex { get; set; }
        [DataMember] public bool MatchCase { get; set; }
        [DataMember] public bool FirstOnly { get; set; }
        [DataMember] public bool MatchesOnly { get; set; } = true;
        [DataMember] public string Prefix { get; set; } = "";
        [DataMember] public string Suffix { get; set; } = "";
        [DataMember] public bool Trim { get; set; }
        [DataMember] public int LetterCase { get; set; }
        [DataMember] public bool Numbering { get; set; }
        [DataMember] public int Start { get; set; } = 1;
        [DataMember] public int Step { get; set; } = 1;
        [DataMember] public int Digits { get; set; } = 2;
        [DataMember] public bool RestartPerBook { get; set; } = true;
        [DataMember] public bool NumberAtEnd { get; set; }
        [DataMember] public string Separator { get; set; } = " ";
        [DataMember] public bool Truncate { get; set; }
        [OnDeserializing] private void Initialize(StreamingContext c)
        { IncludeHidden = IncludeCharts = MatchesOnly = RestartPerBook = true; Start = Step = 1; Digits = 2; Find = Replacement = Prefix = Suffix = ""; Separator = " "; }
    }
    [DataContract]
    public sealed class ToolPreferences
    {
        [DataMember] public int SchemaVersion { get; set; } = 1;
        [DataMember] public RenameOptions Rename { get; set; } = new RenameOptions();
        [DataMember] public CopyOptions Copy { get; set; } = new CopyOptions();
        [DataMember] public string FocusColor { get; set; } = "#42AD91";
        [DataMember] public int FocusOpacity { get; set; } = 20;
        [DataMember] public bool FocusRow { get; set; } = true;
        [DataMember] public bool FocusColumn { get; set; } = true;
        [DataMember] public bool ClearActiveCell { get; set; } = true;
        [DataMember] public int FocusInterval { get; set; } = 33;
        [DataMember] public bool FocusAnimation { get; set; }
        [DataMember] public bool PreserveVisualAlignment { get; set; } = true;
        [DataMember] public bool VisibleCellsOnly { get; set; }
        [DataMember] public bool PreserveLeadingZeros { get; set; } = true;
        [DataMember] public bool ProtectPrecision { get; set; } = true;
        [DataMember] public bool TrimNumericText { get; set; } = true;
        [DataMember] public bool AllowPercent { get; set; }
        [DataMember] public bool AllowThousands { get; set; } = true;
        [DataMember] public bool NumberAsTextHint { get; set; } = true;
        [DataMember] public int ConversionLimit { get; set; } = 50000;
        [DataMember] public int DecimalCulture { get; set; }
        [OnDeserializing] private void Initialize(StreamingContext c)
        { SchemaVersion = 1; Rename = new RenameOptions(); FocusColor = "#42AD91"; FocusOpacity = 20; FocusInterval = 33; PreserveVisualAlignment = FocusRow = FocusColumn = ClearActiveCell = PreserveLeadingZeros = ProtectPrecision = TrimNumericText = AllowThousands = NumberAsTextHint = true; ConversionLimit = 50000; }
        public void Normalize()
        {
            if (Copy == null) Copy = new CopyOptions();
            Copy.Normalize();
            if (Rename == null) Rename = new RenameOptions();
            FocusColor = PreferenceInput.TryColor(FocusColor, out var color) ? color : "#42AD91";
            FocusOpacity = Math.Max(5, Math.Min(60, FocusOpacity)); FocusInterval = Math.Max(16, Math.Min(100, FocusInterval));
            ConversionLimit = Math.Max(100, Math.Min(200000, ConversionLimit)); DecimalCulture = Math.Max(0, Math.Min(1, DecimalCulture));
            Rename.Start = Math.Max(0, Math.Min(1000000, Rename.Start)); Rename.Step = Math.Max(1, Math.Min(1000, Rename.Step)); Rename.Digits = Math.Max(1, Math.Min(10, Rename.Digits));
        }
    }
    public sealed class ToolPreferencesStore
    {
        private readonly string path;
        private bool incompatible;
        public string Status { get; private set; } = "选项保存到加载项所在文件夹。";
        public ToolPreferencesStore(string path) { this.path = path; }
        public ToolPreferences Load()
        {
            try
            {
                if (!File.Exists(path)) return new ToolPreferences();
                using (var stream = File.OpenRead(path))
                {
                    if (stream.Length > 65536) throw new InvalidDataException();
                    var value = (ToolPreferences)new DataContractJsonSerializer(typeof(ToolPreferences)).ReadObject(stream);
                    if (value == null) throw new InvalidDataException();
                    incompatible = value.SchemaVersion != 1;
                    if (incompatible) throw new InvalidDataException();
                    value.Normalize(); return value;
                }
            }
            catch { Status = incompatible ? "配置版本不兼容，不覆盖原文件。" : "配置无法读取，暂用默认选项。"; return new ToolPreferences(); }
        }
        public bool Save(ToolPreferences value)
        {
            if (incompatible) return false;
            string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                value.Normalize();
                using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write)) { new DataContractJsonSerializer(typeof(ToolPreferences)).WriteObject(stream, value); stream.Flush(true); }
                if (File.Exists(path)) File.Replace(temp, path, path + ".bak"); else File.Move(temp, path);
                Status = "已保存，下次启动继续使用。"; return true;
            }
            catch { Status = "未能保存，请确认加载项目录可写；选项仅本次有效。"; return false; }
            finally { try { if (File.Exists(temp)) File.Delete(temp); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
        }
    }
}
