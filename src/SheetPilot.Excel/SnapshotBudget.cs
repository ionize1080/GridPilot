using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace SheetPilot.Excel
{
    public sealed partial class WorkbookSnapshotStore
    {
        public long BudgetBytes { get; set; } = 500L * 1024 * 1024;
        public long UsedBytes => Directory.GetFiles(Path.Combine(Root,"blobs"),"*.gz").Concat(Directory.GetFiles(Path.Combine(Root,"versions"),"*.json")).Sum(p=>new FileInfo(p).Length);
        // Serialize all store transactions across Excel processes sharing this history.
        private IDisposable Transaction()
        {
            var mutex=new System.Threading.Mutex(false,@"Local\GridPilotHistory-"+Hash(System.Text.Encoding.UTF8.GetBytes(Root.ToUpperInvariant())));
            try { mutex.WaitOne(); } catch(System.Threading.AbandonedMutexException) { }
            return new Lease(mutex);
        }
        private sealed class Lease : IDisposable
        {
            private readonly System.Threading.Mutex mutex;
            public Lease(System.Threading.Mutex value){mutex=value;}
            public void Dispose(){mutex.ReleaseMutex();mutex.Dispose();}
        }
        public object Prune()
        {
            using(Transaction())
            {
                var manifests=Directory.GetFiles(Path.Combine(Root,"versions"),"*.json").Select(p=>new {path=p,json=JObject.Parse(File.ReadAllText(p))}).OrderBy(x=>x.json["created"].Value<DateTime>()).ToList();
                var references=new Dictionary<string,int>();
                Func<JObject,IEnumerable<string>> hashes=m=>m["parts"].SelectMany(p=>p["chunks"].Values<string>()).Distinct();
                foreach(var m in manifests)foreach(string h in hashes(m.json)){if(!references.ContainsKey(h))references[h]=0;references[h]++;}
                long used=0;var sizes=new Dictionary<string,long>();
                foreach(string file in Directory.GetFiles(Path.Combine(Root,"blobs"),"*.gz"))
                {
                    string h=Path.GetFileNameWithoutExtension(file);
                    if(!references.ContainsKey(h)){File.Delete(file);continue;}
                    sizes[h]=new FileInfo(file).Length;used+=sizes[h];
                }
                used+=manifests.Sum(m=>new FileInfo(m.path).Length);int removed=0;
                foreach(var m in manifests)
                {
                    if(used<=BudgetBytes)break;
                    long length=new FileInfo(m.path).Length;File.Delete(m.path);used-=length;removed++;
                    foreach(string h in hashes(m.json))if(--references[h]==0&&sizes.ContainsKey(h))
                    {File.Delete(Path.Combine(Root,"blobs",h+".gz"));used-=sizes[h];}
                }
                return new {usedBytes=used,budgetBytes=BudgetBytes,removed};
            }
        }
    }
}
