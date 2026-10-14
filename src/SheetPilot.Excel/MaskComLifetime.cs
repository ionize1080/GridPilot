using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
namespace SheetPilot.Excel
{
    // Only for the private export Excel instance. Never release the user's Excel RCWs.
    internal sealed class MaskComLifetime : IDisposable
    {
        private readonly List<object> objects=new List<object>();
        private readonly HashSet<object> seen=new HashSet<object>();
        public dynamic Hold(object value)
        {if(value!=null && Marshal.IsComObject(value) && seen.Add(value))objects.Add(value);return value;}
        public IDisposable Scope()=>new ScopeLease(this,objects.Count);
        public void Release(object value)
        {if(value!=null && seen.Remove(value)){objects.Remove(value);try{Marshal.FinalReleaseComObject(value);}catch(InvalidComObjectException){}}}
        private sealed class ScopeLease : IDisposable
        {
            private readonly MaskComLifetime owner;private readonly int start;
            internal ScopeLease(MaskComLifetime owner,int start){this.owner=owner;this.start=start;}
            public void Dispose(){for(int i=owner.objects.Count-1;i>=start;i--)owner.Release(owner.objects[i]);}
        }
        public void Dispose()
        {for(int i=objects.Count-1;i>=0;i--)try{Marshal.FinalReleaseComObject(objects[i]);}catch(InvalidComObjectException){}objects.Clear();seen.Clear();}
    }
}
