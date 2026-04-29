using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using BoekSolutions.SheetSetEditor.Helpers;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    /// <summary>
    /// Houdt COM-objecten bij en zorgt voor veilige vrijgave na gebruik.
    /// Gebruik: using (var scope = new ComScope()) { var sheet = scope.Track(...); }
    /// </summary>
    public sealed class ComScope : IDisposable
    {
        private readonly List<object> _tracked = new List<object>();

        /// <summary>
        /// Registreert een COM-object voor latere vrijgave.
        /// </summary>
        public T Track<T>(T obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                _tracked.Add(obj);
                Log.Debug($"[ComScope] Tracked COM object: {obj.GetType().Name}");
            }
            return obj;
        }

        /// <summary>
        /// Laat alle getrackte COM-objecten netjes vrij.
        /// </summary>
        public void Dispose()
        {
            foreach (var obj in _tracked)
            {
                try
                {
                    string typeName = obj?.GetType().Name ?? "unknown";
                    Marshal.ReleaseComObject(obj);
                    Log.Debug($"[ComScope] Released COM object: {typeName}");
                }
                catch (Exception ex)
                {
                    Log.Error($"[ComScope] Fout bij vrijgeven van COM-object: {ex.Message}");
                }
            }

            _tracked.Clear();
        }
    }
}
