// TreeViewSelectionHelper.cs
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using System.Windows.Input;
using System.Windows;
using ACSMCOMPONENTS24Lib;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class TreeViewSelectionHelper
    {
        private static TreeViewItem _lastSelectedItem;
        private static readonly HashSet<IAcSmObjectId> _selection = new HashSet<IAcSmObjectId>();

        public static void HandleMouseClick(TreeView tree, TreeViewItem clickedItem, bool shiftPressed, bool ctrlPressed)
        {
            if (tree == null || clickedItem == null) return;
            tree.Focus();

            if (!(clickedItem.Tag is IAcSmObjectId clickedId)) return;

            List<TreeViewItem> allItems = FlattenTree(tree);
            if (shiftPressed && _lastSelectedItem != null)
            {
                int startIndex = allItems.IndexOf(_lastSelectedItem);
                int endIndex = allItems.IndexOf(clickedItem);
                if (startIndex < 0 || endIndex < 0) return;
                if (startIndex > endIndex)
                {
                    int temp = startIndex;
                    startIndex = endIndex;
                    endIndex = temp;
                }

                ClearSelection();
                for (int i = startIndex; i <= endIndex; i++)
                {
                    TreeViewItem item = allItems[i];
                    if (item.Tag is IAcSmObjectId id)
                    {
                        _selection.Add(id);
                        item.BringIntoView();
                    }
                }
            }
            else if (ctrlPressed)
            {
                if (_selection.Contains(clickedId))
                    _selection.Remove(clickedId);
                else
                    _selection.Add(clickedId);
            }
            else
            {
                ClearSelection();
                _selection.Add(clickedId);
            }

            _lastSelectedItem = clickedItem;

            tree.Dispatcher.InvokeAsync(delegate
            {
                UpdateVisuals(tree);
            }, DispatcherPriority.Background);
        }

        public static List<IAcSmObjectId> GetSelectedSheetIds()
        {
            return new List<IAcSmObjectId>(_selection);
        }

        public static void ClearSelection()
        {
            _selection.Clear();
        }

        public static void UpdateVisuals(ItemsControl tree)
        {
            foreach (TreeViewItem item in FlattenTree(tree))
            {
                if (item.Tag is IAcSmObjectId id)
                {
                    bool selected = _selection.Contains(id);
                    item.Background = selected ? Brushes.LightBlue : Brushes.Transparent;
                    item.FontWeight = selected ? FontWeights.Bold : FontWeights.Normal;
                }
            }
        }

        public static List<TreeViewItem> FlattenTree(ItemsControl parent)
        {
            List<TreeViewItem> result = new List<TreeViewItem>();
            foreach (object item in parent.Items)
            {
                if (parent.ItemContainerGenerator.ContainerFromItem(item) is TreeViewItem tvi)
                {
                    result.Add(tvi);
                    result.AddRange(FlattenTree(tvi));
                }
            }
            return result;
        }

        public static TreeViewItem GetLastSelectedTreeItem()
        {
            return _lastSelectedItem;
        }
    }
}