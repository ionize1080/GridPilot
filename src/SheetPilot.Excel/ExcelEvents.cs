using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace SheetPilot.Excel
{
    // Excel AppEvents connection point. No Office PIA or version-specific assembly needed.
    public sealed class ExcelEvents : IDisposable
    {
        private static readonly Guid InterfaceId = new Guid("00024413-0000-0000-C000-000000000046");
        private object application;
        private readonly List<Tuple<int, Delegate>> handlers = new List<Tuple<int, Delegate>>();
        public ExcelEvents(object application) { this.application = application; }
        public void Selection(Action action) => Add(1558, new Action<object, object>((sheet, target) => action()));
        public void Activation(Action action)
        {
            Add(1561, new Action<object>(sheet => action()));
            Add(1556, new Action<object, object>((book, window) => action()));
            Add(1554, new Action<object, object>((book, window) => action()));
        }
        private void Add(int id, Delegate handler)
        { ComEventsHelper.Combine(application, InterfaceId, id, handler); handlers.Add(Tuple.Create(id, handler)); }
        public void Dispose()
        {
            if (application == null) return;
            foreach (var entry in handlers) { try { ComEventsHelper.Remove(application, InterfaceId, entry.Item1, entry.Item2); } catch (COMException) { } }
            handlers.Clear(); application = null;
        }
    }
}
