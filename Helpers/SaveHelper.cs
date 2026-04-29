using System;
using System.Windows.Controls;
using ACSMCOMPONENTS24Lib;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class SaveHelper
    {
        /// <summary>
        /// Zorgt voor volledige COM-cleanup vóór het unlocken en saven van de database.
        /// Dit mag NOOIT zelf db.Save(null) aanroepen — dat gebeurt via DatabaseHelper.
        /// </summary>
        public static void SaveDatabase(IAcSmDatabase db)
        {
            TryHelper.Run(() =>
            {
                if (db == null)
                {
                    Log.Error("[SaveHelper] Save afgebroken: database is null.");
                    return;
                }

                Log.Info("[SaveHelper] Start save-proces (GC cleanup)...");

                Log.Debug("[SaveHelper] Eerste GC.Collect ronde");
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Log.Debug("[SaveHelper] Tweede GC.Collect voor achterblijvers");
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Log.Info("[SaveHelper] Klaar voor unlock & save via DatabaseHelper.");
                // db.Save(null) doen we dus NIET hier
            }, "SaveHelper.SaveDatabase");
        }


        //Optioneel: als je ClearTreeView nog ergens gebruikt, anders mag deze eruit
        //public static void ClearTreeView(TreeView tree)
        //{
        //    TryHelper.Run(() =>
        //    {
        //        if (tree == null) return;

        //        foreach (TreeViewItem item in tree.Items)
        //        {
        //            item.Tag = null;
        //            item.Items.Clear();
        //        }
        //        tree.Items.Clear();
        //    }, "SaveHelper.ClearTreeView");
        //}
    }
}
