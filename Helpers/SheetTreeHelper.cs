using ACSMCOMPONENTS24Lib;
using BoekSolutions.SheetSetEditor.Helpers; // bevat ComScope, TryHelper, Log
using System.Collections.Generic;
using System.Windows.Controls;

public static class SheetTreeHelper
{
    public static List<IAcSmObjectId> GetSelectedSheetIds(TreeView treeView)
    {
        var result = new List<IAcSmObjectId>();

        var nodes = FlattenTree(treeView);
        foreach (var node in nodes)
        {
            if (node.IsSelected && node.Tag is IAcSmObjectId id)
            {
                TryHelper.Run(() =>
                {
                    using (var scope = new ComScope())
                    {
                        var persist = scope.Track(id.GetPersistObject());
                        if (persist is IAcSmSheet)
                            result.Add(id);
                    }
                }, "SheetTreeHelper.GetSelectedSheetIds");
            }
        }

        return result;
    }

    private static List<TreeViewItem> FlattenTree(ItemsControl root)
    {
        var result = new List<TreeViewItem>();

        foreach (var obj in root.Items)
        {
            var container = root.ItemContainerGenerator.ContainerFromItem(obj) as TreeViewItem;
            if (container != null)
            {
                result.Add(container);
                result.AddRange(FlattenTree(container));
            }
        }

        return result;
    }


    private static void CollectCheckedSheetsRecursive(TreeViewItem node, List<IAcSmObjectId> list)
    {
        TryHelper.Run(() =>
        {
            var checkbox = VisualTreeHelperExtensions.FindVisualChild<CheckBox>(node);
            if (checkbox?.IsChecked == true)
            {
                if (node.Tag is IAcSmObjectId id)
                {
                    using (var scope = new ComScope())
                    {
                        var persist = id.GetPersistObject();
                        if (persist is IAcSmSheet)
                        {
                            list.Add(id);
                        }
                        else
                        {
                            Log.Warn($"[TreeHelper] Object met ID is geen sheet maar {persist?.GetType().Name ?? "null"}");
                        }
                    }
                }
                else
                {
                    Log.Warn("[TreeHelper] TreeViewItem.Tag is geen IAcSmObjectId");
                }
            }

            foreach (object child in node.Items)
            {
                if (child is TreeViewItem childItem)
                    CollectCheckedSheetsRecursive(childItem, list);
            }
        });
    }
}
