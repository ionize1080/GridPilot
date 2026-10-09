using System;
using System.IO;
using System.Runtime.Serialization.Json;

namespace SheetPilot.Core
{
    public sealed class SettingsStore
    {
        public string FilePath { get; }
        public string Status { get; private set; } = "设置更改后自动保存到 XLL 同目录。";
        public bool HasError { get; private set; }
        private bool newerSchema;
        public SettingsStore(string filePath) { FilePath = filePath; }
        public NavigationSettings Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return new NavigationSettings();
                using (var file = File.OpenRead(FilePath))
                {
                    if (file.Length > 65536) throw new InvalidDataException();
                    var settings = (NavigationSettings)new DataContractJsonSerializer(typeof(NavigationSettings)).ReadObject(file);
                    if (settings == null) throw new InvalidDataException();
                    if (settings.SchemaVersion != 1)
                    {
                        newerSchema = true;
                        throw new InvalidDataException();
                    }
                    Status = "已读取便携设置 · 更改会自动保存";
                    return settings;
                }
            }
            catch
            {
                HasError = true;
                Status = newerSchema ? "配置版本不兼容：本次使用默认设置，不覆盖原文件。" : "配置无法读取：本次使用默认设置，更改后重试保存。";
                return new NavigationSettings();
            }
        }
        public bool Save(NavigationSettings settings)
        {
            if (newerSchema) return false;
            string temporary = FilePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var file = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    new DataContractJsonSerializer(typeof(NavigationSettings)).WriteObject(file, settings);
                    file.Flush(true);
                }
                // Same-directory atomic replacement preserves the last complete file on failure.
                if (File.Exists(FilePath)) File.Replace(temporary, FilePath, FilePath + ".bak");
                else File.Move(temporary, FilePath);
                HasError = false;
                Status = "已保存 · 下次启动自动恢复";
                return true;
            }
            catch
            {
                HasError = true;
                Status = "未能保存：请确认 XLL 目录可写。当前设置仅本次有效。";
                return false;
            }
            finally { try { if (File.Exists(temporary)) File.Delete(temporary); } catch { } }
        }
    }
}
