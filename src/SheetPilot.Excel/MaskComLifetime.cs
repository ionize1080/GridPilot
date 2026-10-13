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
        public void Dispose()
        {for(int i=objects.Count-1;i>=0;i--)try{Marshal.FinalReleaseComObject(objects[i]);}catch(InvalidComObjectException){}objects.Clear();seen.Clear();}
    }
}
