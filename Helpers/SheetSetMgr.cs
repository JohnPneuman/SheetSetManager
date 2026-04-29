using System;
using ACSMCOMPONENTS24Lib;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class SheetSetMgr
    {
        /// <summary>
        /// Opent een sheet set database op het opgegeven pad.
        /// </summary>
        public static IAcSmDatabase OpenDatabase(string path, bool readOnly)
        {
            return TryHelper.Run(() =>
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    Log.Error("[SheetSetMgr] Ongeldig pad voor OpenDatabase.");
                    return null;
                }

                Log.Info($"[SheetSetMgr] OpenDatabase: {path} (readOnly={readOnly})");

                var mgr = new AcSmSheetSetMgr();
                return mgr.OpenDatabase(path, readOnly);
            }, "SheetSetMgr.OpenDatabase (with bool)");
        }

        /// <summary>
        /// Maakt een nieuwe instantie van de AcSmSheetSetMgr.
        /// </summary>
        public static AcSmSheetSetMgr CreateManager()
        {
            return TryHelper.Run(() =>
            {
                Log.Debug("[SheetSetMgr] Nieuwe SheetSetMgr instantie aangemaakt.");
                return new AcSmSheetSetMgr();
            }, "SheetSetMgr.CreateManager");
        }
    }
}
