// Aangepaste versie van SheetHelper.cs met WPF-conforme selectie (TreeViewItem.IsSelected)
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows;
using ACSMCOMPONENTS24Lib;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class SheetHelper
    {
        // Deze methode vervangt oude checkboxlogica en haalt geselecteerde SheetIds uit TreeViewItems
        public static List<IAcSmObjectId> GetSelectedSheetIds(TreeView treeView)
        {
            var selectedSheetIds = new List<IAcSmObjectId>();

            foreach (var item in GetTreeViewItems(treeView))
            {
                if (item is TreeViewItem tvi && tvi.IsSelected)
                {
                    if (tvi.Tag is IAcSmObjectId id)
                    {
                        selectedSheetIds.Add(id);
                    }
                }
            }

            return selectedSheetIds;
        }

        // Recursief alle TreeViewItems ophalen
        private static IEnumerable<TreeViewItem> GetTreeViewItems(ItemsControl parent)
        {
            for (int i = 0; i < parent.Items.Count; i++)
            {
                var item = parent.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                if (item != null)
                {
                    yield return item;
                    foreach (var child in GetTreeViewItems(item))
                        yield return child;
                }
            }
        }

        public static void ApplyChanges(IAcSmObjectId sheetId, string newNumber, string newTitle)
        {
            if (sheetId == null || !sheetId.IsValid()) return;

            TryHelper.Run(() =>
            {
                using (var scope = new ComScope())
                {
                    var sheet = ComHelper.GetObject<IAcSmSheet>(sheetId);
                    if (sheet == null) return;

                    var name = sheet.GetName();
                    if (newNumber != null && sheet.GetNumber() != newNumber)
                    {
                        sheet.SetNumber(newNumber);
                        Log.Info($"[Apply] Nummer '{name}' → {newNumber}");
                    }

                    if (newTitle != null && sheet.GetTitle() != newTitle)
                    {
                        sheet.SetTitle(newTitle);
                        Log.Info($"[Apply] Titel '{name}' → {newTitle}");
                    }
                }
            }, "SheetHelper.ApplyChanges");
        }

        public static IAcSmObjectId GetObjectId(IAcSmPersist persist)
        {
            return TryHelper.Run(() =>
            {
                if (persist == null)
                {
                    Log.Error("[SheetHelper] GetObjectId: persist is null.");
                    return null;
                }

                var db = persist.GetDatabase();
                if (db == null)
                {
                    Log.Error("[SheetHelper] GetObjectId: db is null.");
                    return null;
                }

                using (var scope = new ComScope())
                {
                    var enumerator = scope.Track(db.GetEnumerator());
                    IAcSmPersist item;
                    while ((item = scope.Track(enumerator.Next())) != null)
                    {
                        if (ReferenceEquals(item, persist))
                        {
                            var id = item.GetObjectId();
                            if (id == null)
                                Log.Warn("[SheetHelper] GetObjectId: objectId is null ondanks match.");
                            return id;
                        }
                    }

                    Log.Warn("[SheetHelper] GetObjectId: object niet gevonden in enumerator.");
                    return null;
                }
            }, "SheetHelper.GetObjectId");
        }
    }
}
