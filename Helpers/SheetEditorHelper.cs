using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls.Primitives;
using ACSMCOMPONENTS24Lib;
using BoekSolutions.SheetSetEditor.Models;
using BoekSolutions.SheetSetEditor.UI.Controls;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class SheetEditorHelper
    {
        public static event Action<IAcSmObjectId> SheetSelected;

        public static void AttachTreeSelection(TreeView treeView)
        {
            treeView.PreviewMouseLeftButtonDown += (s, e) =>
            {
                var point = e.GetPosition(treeView);
                var element = treeView.InputHitTest(point) as DependencyObject;

                bool isExpander = false;
                var current = element;
                while (current != null)
                {
                    if (current is ToggleButton)
                    {
                        isExpander = true;
                        break;
                    }
                    current = VisualTreeHelper.GetParent(current);
                }
                if (isExpander) return;

                while (element != null && !(element is TreeViewItem))
                    element = VisualTreeHelper.GetParent(element);

                if (element is TreeViewItem item)
                {
                    TreeViewSelectionHelper.HandleMouseClick(
                        treeView,
                        item,
                        Keyboard.Modifiers.HasFlag(ModifierKeys.Shift),
                        Keyboard.Modifiers.HasFlag(ModifierKeys.Control)
                    );

                    var selectedIds = TreeViewSelectionHelper.GetSelectedSheetIds();
                    if (selectedIds.Count == 1)
                    {
                        SheetSelected?.Invoke(selectedIds[0]); // Enkele selectie
                    }
                    else
                    {
                        SheetSelected?.Invoke(null); // Meervoudige selectie → trigger lege of samengevoegde view
                    }

                    e.Handled = true;
                }

            };

            treeView.PreviewKeyDown += (s, e) =>
            {
                if (e.Key == Key.Up || e.Key == Key.Down)
                {
                    var current = TreeViewSelectionHelper.GetLastSelectedTreeItem();
                    if (current == null) return;

                    var parent = ItemsControl.ItemsControlFromItemContainer(current);
                    if (parent == null) return;

                    int index = parent.ItemContainerGenerator.IndexFromContainer(current);
                    int newIndex = e.Key == Key.Up ? index - 1 : index + 1;

                    if (newIndex >= 0 && newIndex < parent.Items.Count)
                    {
                        var nextItem = parent.ItemContainerGenerator.ContainerFromIndex(newIndex) as TreeViewItem;
                        if (nextItem != null)
                        {
                            TreeViewSelectionHelper.HandleMouseClick(
                                treeView,
                                nextItem,
                                Keyboard.Modifiers.HasFlag(ModifierKeys.Shift),
                                Keyboard.Modifiers.HasFlag(ModifierKeys.Control)
                            );

                            var selectedIds = TreeViewSelectionHelper.GetSelectedSheetIds();
                            if (selectedIds.Count == 1)
                                SheetSelected?.Invoke(selectedIds[0]);
                            else
                                SheetSelected?.Invoke(null);

                            nextItem.BringIntoView();
                            e.Handled = true;
                        }

                    }
                }
            };
        }

        public static void ApplyUpdates(IEnumerable<SheetUpdateModel> updates, TreeView treeView)
        {
            if (updates == null) return;

            string Safe(string val)
            {
                return string.IsNullOrWhiteSpace(val) || val == "…" ? null : val;
            }

            foreach (var update in updates)
            {
                TryHelper.Run(() =>
                {
                    if (update?.SheetId == null || !update.SheetId.IsValid())
                        return;

                    using (var scope = new ComScope())
                    {
                        var sheet = scope.Track(ComHelper.GetObject<IAcSmSheet>(update.SheetId));
                        if (sheet == null) return;

                        // ===== NUMMER =====
                        string curNum = sheet.GetNumber() ?? "";
                        string newNum = Safe(update.NewNumber);
                        if (!string.IsNullOrEmpty(newNum) && newNum != curNum)
                            sheet.SetNumber(newNum);

                        // ===== TITEL =====
                        string curTitle = sheet.GetTitle() ?? "";
                        string newTitle = Safe(update.Title);
                        if (!string.IsNullOrEmpty(newTitle) && newTitle != curTitle)
                            sheet.SetTitle(newTitle);

                        // ===== OMSCHRIJVING =====
                        string curDesc = sheet.GetDesc() ?? "";
                        string newDesc = Safe(update.Description);
                        if (!string.IsNullOrEmpty(newDesc) && newDesc != curDesc)
                            sheet.SetDesc(newDesc);

                        // ===== REVISION, CATEGORY, ISSUE, ETC. =====
                        var sheet2 = sheet as IAcSmSheet2;
                        if (sheet2 != null)
                        {
                            // Revision Number
                            string curRevision = sheet2.GetRevisionNumber() ?? "";
                            string newRevision = Safe(update.RevisionNumber);
                            if (!string.IsNullOrEmpty(newRevision) && newRevision != curRevision)
                                sheet2.SetRevisionNumber(newRevision);

                            // Revision Date
                            string curRevisionDate = sheet2.GetRevisionDate() ?? "";
                            string newRevisionDate = Safe(update.RevisionDate);
                            if (!string.IsNullOrEmpty(newRevisionDate) && newRevisionDate != curRevisionDate)
                                sheet2.SetRevisionDate(newRevisionDate);

                            // Issue Purpose
                            string curIssuePurpose = sheet2.GetIssuePurpose() ?? "";
                            string newIssuePurpose = Safe(update.IssuePurpose);
                            if (!string.IsNullOrEmpty(newIssuePurpose) && newIssuePurpose != curIssuePurpose)
                                sheet2.SetIssuePurpose(newIssuePurpose);

                            // Category
                            string curCategory = sheet2.GetCategory() ?? "";
                            string newCategory = Safe(update.Category);
                            if (!string.IsNullOrEmpty(newCategory) && newCategory != curCategory)
                                sheet2.SetCategory(newCategory);
                        }

                        // ===== DO NOT PLOT =====
                        if (update.DoNotPlot.HasValue)
                        {
                            bool curDoNotPlot = sheet.GetDoNotPlot();
                            if (update.DoNotPlot.Value != curDoNotPlot)
                                sheet.SetDoNotPlot(update.DoNotPlot.Value);
                        }

                        // ===== CUSTOM PROPERTIES =====
                        if (update.CustomProperties != null)
                        {
                            foreach (var kvp in update.CustomProperties)
                            {
                                if (string.Equals(kvp.Key, "Description", StringComparison.OrdinalIgnoreCase))
                                    continue;

                                var safeValue = Safe(kvp.Value);
                                if (string.IsNullOrEmpty(safeValue)) continue;

                                string current = PropertyBagHelper.GetCustomProperty(update.SheetId, kvp.Key) ?? "";
                                if (safeValue != current)
                                {
                                    PropertyBagHelper.SetCustomPropertyIfValid(update.SheetId, kvp.Key, safeValue);
                                }
                            }
                        }
                    }

                    // ===== TREEVIEW REFRESH =====
                    var item = FindTreeViewItemById(treeView, update.SheetId);
                    if (item?.Header is TextBlock tb)
                    {
                        var displayNumber = Safe(update.NewNumber)
                            ?? (ComHelper.GetObject<IAcSmSheet>(update.SheetId)?.GetNumber() ?? "");
                        var displayTitle = Safe(update.Title)
                            ?? (ComHelper.GetObject<IAcSmSheet>(update.SheetId)?.GetTitle() ?? "");

                        tb.Text = SheetDisplayHelper.FormatSheetLabel(displayNumber, displayTitle);
                    }

                }, "ApplyUpdate");
            }
        }

        public static void AutonumberSelected(TreeView treeView, SheetTreeControl sheetTreeControl, AutoNumberOptions options = null)
        {
            var ids = TreeViewSelectionHelper.GetSelectedSheetIds();

            int start = options != null ? options.StartNumber : 1;
            int increment = options != null ? options.Increment : 1;
            string prefix = options != null ? options.Prefix : "";
            string suffix = options != null ? options.Suffix : "";

            int counter = start;
            var updates = new List<SheetUpdateModel>();

            foreach (var id in ids)
            {
                var sheet = ComHelper.GetObject<IAcSmSheet>(id);
                if (sheet == null) continue;

                string pageNumber = prefix + counter + suffix;
                string title = sheet.GetTitle() ?? "";

                updates.Add(new SheetUpdateModel
                {
                    SheetId = id,
                    NewNumber = pageNumber,
                    Title = title
                });

                counter += increment;
            }

            ApplyUpdates(updates, treeView);
        }

        private static TreeViewItem FindTreeViewItemById(ItemsControl parent, IAcSmObjectId id)
        {
            foreach (var item in parent.Items)
            {
                var container = parent.ItemContainerGenerator.ContainerFromItem(item) as TreeViewItem;
                if (container != null)
                {
                    if (container.Tag is IAcSmObjectId foundId && foundId.IsEqual(id))
                        return container;

                    var child = FindTreeViewItemById(container, id);
                    if (child != null)
                        return child;
                }
            }
            return null;
        }

        public static string GetSharedValue(List<string> values)
        {
            return values.Distinct().Count() == 1 ? values.First() : "…";
        }

        public static Dictionary<string, string> GetSharedCustomProperties(List<IAcSmObjectId> ids)
        {
            var result = new Dictionary<string, string>();
            var all = ids.Select(PropertyBagHelper.GetCustomProperties).ToList();
            if (all.Count == 0) return result;

            var keys = all.SelectMany(p => p.Select(x => x.PropertyName)).Distinct();

            foreach (var key in keys)
            {
                var values = all.SelectMany(p => p).Where(p => p.PropertyName == key).Select(p => p.PropertyValue).ToList();
                result[key] = GetSharedValue(values);
            }

            return result;
        }

        public static bool GetSharedBoolValue(List<bool> values)
        {
            if (values == null || values.Count == 0)
                return false; // of null als je NewDoNotPlot nullable maakt

            bool first = values[0];
            foreach (bool b in values)
                if (b != first)
                    return false; // of null voor "mixed"
            return first;
        }


    }
}
