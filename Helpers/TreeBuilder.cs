// TreeBuilder.cs
using System.Windows.Controls;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using ACSMCOMPONENTS24Lib;
using System.Collections.Generic;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class TreeBuilder
    {
        public static void BuildFromSheetSet(TreeView target, IAcSmObjectId sheetSetId)
        {
            target.Items.Clear();

            TryHelper.Run(() =>
            {
                using (var scope = new ComScope())
                {
                    var sheetSet = scope.Track(ComHelper.GetObject<IAcSmSheetSet>(sheetSetId));
                    if (sheetSet == null) return;

                    var rootNode = new TreeViewItem
                    {
                        Header = new TextBlock { Text = sheetSet.GetName(), FontWeight = FontWeights.Bold },
                        Tag = sheetSetId,
                        IsExpanded = true
                    };
                    target.Items.Add(rootNode);

                    var enumComp = scope.Track(sheetSet.GetSheetEnumerator());
                    AddChildren(enumComp, rootNode, scope, target);
                }
            });
        }

        private static void AddChildren(IAcSmEnumComponent enumComp, ItemsControl parent, ComScope scope, TreeView rootTree)
        {
            TryHelper.Run(() =>
            {
                IAcSmComponent comp;
                while ((comp = scope.Track(enumComp.Next())) != null)
                {
                    if (comp is IAcSmSubset subset)
                    {
                        var id = SheetHelper.GetObjectId(subset);

                        var label = CreateTextBlock(subset.GetName(), rootTree);

                        var node = new TreeViewItem
                        {
                            Header = label,
                            Tag = id,
                            IsExpanded = true
                        };
                        parent.Items.Add(node);

                        var childEnum = scope.Track(subset.GetSheetEnumerator());
                        AddChildren(childEnum, node, scope, rootTree);
                    }
                    else if (comp is IAcSmSheet sheet)
                    {
                        var id = SheetHelper.GetObjectId(sheet);

                        var label = CreateTextBlock(SheetDisplayHelper.FormatSheetLabel(sheet.GetNumber(), sheet.GetTitle()), rootTree);

                        var node = new TreeViewItem
                        {
                            Header = label,
                            Tag = id
                        };
                        parent.Items.Add(node);
                    }
                }
            });
        }

        private static TextBlock CreateTextBlock(string text, TreeView tree)
        {
            var label = new TextBlock
            {
                Text = text,
                VerticalAlignment = VerticalAlignment.Center
            };
            label.MouseLeftButtonDown += (s, e) => HandleMouse(label, tree, e);
            return label;
        }

        private static void HandleMouse(TextBlock label, TreeView tree, MouseButtonEventArgs e)
        {
            bool shift = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
            bool ctrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;

            var item = FindParent<TreeViewItem>(label);
            if (item != null)
            {
                TreeViewSelectionHelper.HandleMouseClick(tree, item, shift, ctrl);
            }
        }

        private static T FindParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);
            while (parent != null && !(parent is T))
                parent = VisualTreeHelper.GetParent(parent);
            return parent as T;
        }
    }
}
