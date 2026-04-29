// SheetTreeControl.xaml.cs
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ACSMCOMPONENTS24Lib;
using BoekSolutions.SheetSetEditor.Helpers;

namespace BoekSolutions.SheetSetEditor.UI.Controls
{
    public partial class SheetTreeControl : UserControl
    {
        public TreeView SheetTree => TreeViewSheets; // Zorg dat je TreeView in XAML de naam "TreeViewSheets" heeft

        public event EventHandler<IAcSmObjectId> SheetSelected;

        public SheetTreeControl()
        {
            InitializeComponent();
        }

        public void LoadSheetSet(IAcSmObjectId sheetSetId)
        {
            TreeBuilder.BuildFromSheetSet(TreeViewSheets, sheetSetId);
        }

        public void ClearTree()
        {
            TreeViewSheets.Items.Clear();
        }

        public List<IAcSmObjectId> GetSelectedSheetIds()
        {
            return SheetTreeHelper.GetSelectedSheetIds(TreeViewSheets);
        }

        public ItemsControl TreeView => TreeViewSheets;

        public void UpdateTreeNodeHeader(IAcSmObjectId targetId, string newNumber, string newTitle = null)
        {
            UpdateTreeNodeHeaderRecursive(TreeViewSheets, targetId, newNumber, newTitle);
        }

        private void UpdateTreeNodeHeaderRecursive(ItemsControl parent, IAcSmObjectId targetId, string newNumber, string newTitle)
        {
            if (parent == null || targetId == null)
                return;

            foreach (var item in parent.Items)
            {
                if (item is TreeViewItem node)
                {
                    if (node.Tag is IAcSmObjectId currentId && currentId.IsEqual(targetId))
                    {
                        var displayText = SheetDisplayHelper.FormatSheetLabel(newNumber ?? "", newTitle ?? "");

                        if (node.Header is StackPanel panel)
                        {
                            foreach (var child in panel.Children)
                            {
                                if (child is TextBlock tb)
                                {
                                    tb.Text = displayText;
                                    Log.Debug($"[SSM] Boomnode geüpdatet: {displayText}");
                                    return;
                                }
                            }
                        }
                        else if (node.Header is string)
                        {
                            node.Header = displayText;
                            Log.Debug($"[SSM] Boomnode (string) geüpdatet: {displayText}");
                            return;
                        }
                    }
                    else if (node.Items.Count > 0)
                    {
                        UpdateTreeNodeHeaderRecursive(node, targetId, newNumber, newTitle);
                    }
                }
            }
        }

        private void TreeViewSheets_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (TreeViewSheets.SelectedItem is TreeViewItem item && item.Tag is IAcSmObjectId id)
            {
                TryHelper.Run(() =>
                {
                    using (var scope = new ComScope())
                    {
                        var persist = scope.Track(id.GetPersistObject());

                        if (persist is IAcSmSheet)
                            SheetSelected?.Invoke(this, id);
                        else
                            SheetSelected?.Invoke(this, null);
                    }
                }, "TreeViewSheets_SelectedItemChanged");
            }
        }
    }
}
