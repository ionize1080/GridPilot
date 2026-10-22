using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    // Immutable manifests + content-addressed compressed chunks. ZIP parts are split
    // before deduplication so recompressing an xlsx does not invalidate every chunk.
    public sealed class WorkbookSnapshotStore
    {
        public string Root { get; }
        public WorkbookSnapshotStore(string root) { Root = Path.GetFullPath(root); Directory.CreateDirectory(Path.Combine(Root, "blobs")); Directory.CreateDirectory(Path.Combine(Root, "versions")); }
        private static string Hash(byte[] bytes) { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }
        private static void Atomic(string path, byte[] data)
        {
            if (File.Exists(path)) return;
            string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try { using (var file = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { file.Write(data, 0, data.Length); file.Flush(true); } try { File.Move(temp, path); } catch (IOException) { if (!File.Exists(path)) throw; } }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }
        private JArray Put(Stream input)
        {
            var chunks = new JArray(); byte[] buffer = new byte[256 * 1024];
            while (true)
            {
                int length = 0, read; while (length < buffer.Length && (read = input.Read(buffer, length, buffer.Length - length)) > 0) length += read;
                if (length == 0) break;
                byte[] bytes = buffer.Take(length).ToArray(); string hash = Hash(bytes); chunks.Add(hash);
                string path = Path.Combine(Root, "blobs", hash + ".gz");
                if (!File.Exists(path)) using (var memory = new MemoryStream()) { using (var gzip = new GZipStream(memory, CompressionLevel.Optimal, true)) gzip.Write(bytes, 0, bytes.Length); Atomic(path, memory.ToArray()); }
            }
            return chunks;
        }
        public string Save(string source, string book, string label)
        {
            string id = Guid.NewGuid().ToString("N"), extension = Path.GetExtension(source).ToLowerInvariant();
            var parts = new JArray(); bool zip = extension != ".xls";
            if (zip) using (var archive = ZipFile.OpenRead(source)) foreach (var entry in archive.Entries) using (var stream = entry.Open()) parts.Add(new JObject { ["name"] = entry.FullName, ["chunks"] = Put(stream) });
            else using (var stream = File.OpenRead(source)) parts.Add(new JObject { ["name"] = "workbook", ["chunks"] = Put(stream) });
            var manifest = new JObject { ["id"] = id, ["book"] = book, ["label"] = label, ["created"] = DateTime.UtcNow.ToString("o"), ["extension"] = extension, ["zip"] = zip, ["parts"] = parts };
            Atomic(Path.Combine(Root, "versions", id + ".json"), System.Text.Encoding.UTF8.GetBytes(manifest.ToString())); return id;
        }
        public JObject Manifest(string id)
        {
            if (!System.Text.RegularExpressions.Regex.IsMatch(id ?? "", "^[a-f0-9]{32}$")) throw new ArgumentException("无效快照。");
            return JObject.Parse(File.ReadAllText(Path.Combine(Root, "versions", id + ".json")));
        }
        public string Signature(string id)
        {
            // Saving updates core timestamps even without an edit. All workbook,
            // sheet, style, VBA, pivot and drawing parts remain part of this guard.
            var parts = new JArray();
            foreach (var part in Manifest(id)["parts"].OrderBy(p => (string)p["name"]))
            {
                if ((string)part["name"] != "docProps/core.xml" && (string)part["name"] != "xl/workbook.xml" && (string)part["name"] != "xl/workbook.bin" && (string)part["name"] != "workbook" && (string)part["name"] != "xl/vbaProject.bin") { parts.Add(part.DeepClone()); continue; }
                using (var memory = new MemoryStream())
                {
                    Copy((JArray)part["chunks"], memory);
                    byte[] normalized;
                    if ((string)part["name"] == "workbook" || (string)part["name"] == "xl/vbaProject.bin") normalized = CompoundSignature(memory.ToArray());
                    else if ((string)part["name"] == "xl/workbook.bin")
                    {
                        var latin=System.Text.Encoding.GetEncoding(28591);string binary=latin.GetString(memory.ToArray());
                        binary=System.Text.RegularExpressions.Regex.Replace(binary,@"8\x00_\x00\{\x00(?:[A-Fa-f0-9-]\x00){36}\}\x00",m=>new string('\0',m.Length));
                        normalized=latin.GetBytes(binary);
                    }
                    else { string xml = System.Text.Encoding.UTF8.GetString(memory.ToArray()); xml = System.Text.RegularExpressions.Regex.Replace(xml, @"<\w+:revisionPtr\b[^>]*/>", ""); if((string)part["name"]=="docProps/core.xml")xml=System.Text.RegularExpressions.Regex.Replace(xml,@"<dcterms:modified\b[^>]*>[^<]*</dcterms:modified>",""); normalized=System.Text.Encoding.UTF8.GetBytes(xml); }
                    parts.Add(new JObject { ["name"] = part["name"], ["hash"] = Hash(normalized) });
                }
            }
            return Hash(System.Text.Encoding.UTF8.GetBytes(parts.ToString(Newtonsoft.Json.Formatting.None)));
        }
        // Canonical stream contents, excluding allocation layout and save timestamps.
        // PROJECT protection bytes are decoded (MS-OVBA 2.4.3.3), so randomized
        // obfuscation does not look like an edit; changed protection still does.
        private static byte[] CompoundSignature(byte[] bytes)
        {
            if(bytes.Length<512||BitConverter.ToUInt64(bytes,0)!=0xE11AB1A1E011CFD0UL)return bytes;
            int shift=BitConverter.ToUInt16(bytes,30);if(shift!=9&&shift!=12)throw new IOException("无效 CFB 扇区。");
            int size=1<<shift,count=bytes.Length/size-1;
            Func<uint,int> offset=sector=>{if(sector>=count)throw new IOException("无效 CFB 扇区。");return checked(((int)sector+1)*size);};
            var fats=new System.Collections.Generic.List<uint>();
            for(int i=0;i<109;i++){uint sec=BitConverter.ToUInt32(bytes,76+i*4);if(sec<0xFFFFFFFA)fats.Add(sec);}
            uint dif=BitConverter.ToUInt32(bytes,68),difCount=BitConverter.ToUInt32(bytes,72);if(difCount>count)throw new IOException("无效 CFB FAT。");
            for(uint i=0;i<difCount;i++){int at=offset(dif);for(int k=0;k<size/4-1;k++){uint sec=BitConverter.ToUInt32(bytes,at+k*4);if(sec<0xFFFFFFFA)fats.Add(sec);}dif=BitConverter.ToUInt32(bytes,at+size-4);}
            Func<uint,uint> next=sec=>{uint page=sec/(uint)(size/4);if(page>=fats.Count)throw new IOException("无效 CFB FAT。");return BitConverter.ToUInt32(bytes,offset(fats[(int)page])+(int)(sec%(size/4))*4);};
            Func<uint,long,byte[]> read=(start,length)=>{
                using(var data=new MemoryStream()){
                    var seen=new System.Collections.Generic.HashSet<uint>();uint sec=start;
                    while(sec!=0xFFFFFFFE&&data.Length<length){if(!seen.Add(sec))throw new IOException("CFB 扇区循环。");int at=offset(sec),n=(int)Math.Min(size,length-data.Length);data.Write(bytes,at,n);sec=next(sec);}
                    if(length!=long.MaxValue&&data.Length!=length)throw new IOException("CFB 流不完整。");return data.ToArray();
                }
            };
            byte[] dirs=read(BitConverter.ToUInt32(bytes,48),long.MaxValue);
            if(dirs.Length<128)throw new IOException("CFB 根目录缺失。");
            byte[] mini=read(BitConverter.ToUInt32(dirs,116),BitConverter.ToUInt32(dirs,120));
            uint miniCount=BitConverter.ToUInt32(bytes,64);
            byte[] miniFat=miniCount==0?new byte[0]:read(BitConverter.ToUInt32(bytes,60),(long)miniCount*size);
            var entries=new JArray();var visited=new System.Collections.Generic.HashSet<uint>();Action<uint,string> visit=null;
            visit=(id,parent)=>{
                if(id==0xFFFFFFFF)return;if(id>=dirs.Length/128||!visited.Add(id))throw new IOException("无效 CFB 目录。");
                int at=(int)id*128,len=BitConverter.ToUInt16(dirs,at+64);if(len<2||len>64)throw new IOException("无效 CFB 名称。");
                string name=System.Text.Encoding.Unicode.GetString(dirs,at,len-2),path=parent+"/"+name;byte type=dirs[at+66];
                visit(BitConverter.ToUInt32(dirs,at+68),parent);visit(BitConverter.ToUInt32(dirs,at+72),parent);
                var entry=new JObject{["path"]=path,["type"]=type,["class"]=Convert.ToBase64String(dirs,at+80,20)};
                if(type==2){
                    uint start=BitConverter.ToUInt32(dirs,at+116);long length=BitConverter.ToUInt32(dirs,at+120);byte[] data;
                    if(length>=BitConverter.ToUInt32(bytes,56))data=read(start,length);
                    else using(var m=new MemoryStream()){
                        var seen=new System.Collections.Generic.HashSet<uint>();uint sec=start;
                        while(m.Length<length){if(sec>=mini.Length/64||sec>=miniFat.Length/4||!seen.Add(sec))throw new IOException("无效 CFB 小流。");int n=(int)Math.Min(64,length-m.Length);m.Write(mini,(int)sec*64,n);sec=BitConverter.ToUInt32(miniFat,(int)sec*4);}data=m.ToArray();
                    }
                    if(name=="PROJECT"){
                        var latin=System.Text.Encoding.GetEncoding(28591);string project=latin.GetString(data);
                        project=System.Text.RegularExpressions.Regex.Replace(project,@"(?m)^ID=""\{[A-Fa-f0-9-]{36}\}""","ID=stable");
                        project=System.Text.RegularExpressions.Regex.Replace(project,@"(?m)^(CMG|DPB|GC)=""([A-Fa-f0-9]+)""",m=>m.Groups[1].Value+"="+DecodeProject(m.Groups[2].Value));
                        data=latin.GetBytes(project);
                    }
                    entry["hash"]=Hash(data);
                }
                entries.Add(entry);if(type==1||type==5)visit(BitConverter.ToUInt32(dirs,at+76),path);
            };
            visit(BitConverter.ToUInt32(dirs,76),"");
            return System.Text.Encoding.UTF8.GetBytes(new JArray(entries.OrderBy(e=>(string)e["path"])).ToString(Newtonsoft.Json.Formatting.None));
        }
        private static string DecodeProject(string hex)
        {
            if(hex.Length%2!=0)throw new IOException("无效 VBA 工程元数据。");
            var b=Enumerable.Range(0,hex.Length/2).Select(i=>Convert.ToByte(hex.Substring(i*2,2),16)).ToArray();
            if(b.Length<8||(b[0]^b[1])!=2)throw new IOException("无效 VBA 工程元数据。");
            int plain=b[0]^b[2],last=b[2],prior=b[1],skip=(b[0]&6)/2;var data=new byte[b.Length-3];
            for(int i=3;i<b.Length;i++){int value=b[i]^((prior+plain)&255);prior=last;last=b[i];plain=value;data[i-3]=(byte)value;}
            if(skip+4>data.Length)throw new IOException("无效 VBA 工程元数据。");uint length=BitConverter.ToUInt32(data,skip);
            if(length!=data.Length-skip-4)throw new IOException("无效 VBA 工程元数据。");
            return Convert.ToBase64String(data,skip+4,(int)length);
        }
        private void Copy(JArray chunks, Stream output)
        {
            foreach (string hash in chunks.Values<string>())
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(hash, "^[a-f0-9]{64}$")) throw new IOException("快照清单损坏。");
                using (var input = File.OpenRead(Path.Combine(Root, "blobs", hash + ".gz"))) using (var gzip = new GZipStream(input, CompressionMode.Decompress)) using (var memory = new MemoryStream())
                { gzip.CopyTo(memory); var bytes = memory.ToArray(); if (Hash(bytes) != hash) throw new IOException("快照校验失败，未打开恢复文件。"); output.Write(bytes, 0, bytes.Length); }
            }
        }
        public string Materialize(string id, string directory)
        {
            var manifest = Manifest(id); Directory.CreateDirectory(directory);
            string path = Path.Combine(directory, "Restore-" + id + "-" + Guid.NewGuid().ToString("N").Substring(0, 6) + (string)manifest["extension"]);
            try
            {
                if ((bool)manifest["zip"]) using (var archive = ZipFile.Open(path, ZipArchiveMode.Create)) foreach (JObject part in manifest["parts"]) using (var output = archive.CreateEntry((string)part["name"], CompressionLevel.Optimal).Open()) Copy((JArray)part["chunks"], output);
                else using (var output = new FileStream(path, FileMode.CreateNew)) Copy((JArray)manifest["parts"][0]["chunks"], output);
                return path;
            }
            catch { if (File.Exists(path)) File.Delete(path); throw; }
        }
        public JArray List() => new JArray(Directory.GetFiles(Path.Combine(Root, "versions"), "*.json").Select(p => JObject.Parse(File.ReadAllText(p))).OrderByDescending(m => (string)m["created"]).Select(m => new JObject { ["id"] = m["id"], ["book"] = m["book"], ["label"] = m["label"], ["created"] = m["created"] }));
    }
}
