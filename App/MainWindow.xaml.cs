using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using SheetSet.Core.Import;
using SheetSet.Core.Interop;
using SheetSet.Core.Models;
using SheetSet.Core.Parsing;
using SheetSet.Core.Writing;
using SheetSetEditor.Models;
using SheetSetEditor.Services;
using SheetSetEditor.ViewModels;

namespace SheetSetEditor;

public partial class MainWindow : Window
{
    private readonly MainViewModel _vm = new();
    private readonly SheetSetParser _parser = new();

    // Drag & drop state
    private Point _dragStart;
    private bool _isDragging;
    private object? _dragPayload; // SheetNode or SubsetNode
    private DocumentContext? _dragSourceCtx;
    private TreeViewItem? _dropHighlightedItem;

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _vm;
    }

    // Called from App.xaml.cs when launched with a file argument
    public void OpenFromCommandLine(string path)
    {
        if (File.Exists(path)) LoadFile(path);
    }

    // ═══════════════════════════════════════════════════════════
    // File dropdown (dynamic context menu)
    // ═══════════════════════════════════════════════════════════

    private void FileDropdownButton_Click(object sender, RoutedEventArgs e)
    {
        var menu = new ContextMenu { PlacementTarget = FileDropdownButton };

        // ── Open documents ──
        foreach (var ctx in _vm.OpenDocuments)
        {
            bool isActive = ctx == _vm.ActiveDocument;
            var item = new MenuItem
            {
                Header = ctx.DisplayName + (ctx.IsDirty ? " *" : ""),
                FontWeight = isActive ? FontWeights.Bold : FontWeights.Normal,
                IsChecked = isActive
            };
            var capturedCtx = ctx;
            item.Click += (_, _) => SwitchToDocument(capturedCtx);
            menu.Items.Add(item);
        }

        if (_vm.OpenDocuments.Count > 0) menu.Items.Add(new Separator());

        // ── Recent files submenu ──
        var openPaths = _vm.OpenDocuments.Select(c => c.FilePath).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var recent = RecentFilesService.Load().Where(p => !openPaths.Contains(p)).Take(20).ToList();
        var recentMenu = new MenuItem { Header = "Recente bestanden", IsEnabled = recent.Count > 0 };
        foreach (var path in recent)
        {
            var sub = new MenuItem { Header = Path.GetFileNameWithoutExtension(path), ToolTip = path };
            var capturedPath = path;
            sub.Click += (_, _) => LoadFile(capturedPath);
            recentMenu.Items.Add(sub);
        }
        menu.Items.Add(recentMenu);
        menu.Items.Add(new Separator());

        // ── Actions ──
        AddMenuItem(menu, "Openen…",              (_, _) => OpenFile());
        AddMenuItem(menu, "CSV/TSV importeren…",  (_, _) => ImportCsv(),  _vm.IsDocumentLoaded);
        menu.Items.Add(new Separator());
        AddMenuItem(menu, "Sluiten",              (_, _) => CloseActiveDocument(), _vm.IsDocumentLoaded);

        menu.IsOpen = true;
    }

    private static void AddMenuItem(ContextMenu menu, string header,
        RoutedEventHandler click, bool enabled = true)
    {
        var item = new MenuItem { Header = header, IsEnabled = enabled };
        item.Click += click;
        menu.Items.Add(item);
    }

    // ═══════════════════════════════════════════════════════════
    // File operations
    // ═══════════════════════════════════════════════════════════

    private void OpenFile()
    {
        var dlg = new OpenFileDialog
        {
            Title = "Open Sheet Set",
            Filter = "Sheet Set bestanden (*.dst)|*.dst|XML bestanden (*.xml)|*.xml|Alle bestanden (*.*)|*.*",
            Multiselect = true
        };
        if (dlg.ShowDialog() != true) return;
        foreach (var f in dlg.FileNames) LoadFile(f);
    }

    public void LoadFile(string path)
    {
        // Don't open twice
        if (_vm.OpenDocuments.Any(c =>
            string.Equals(c.FilePath, path, StringComparison.OrdinalIgnoreCase)))
        {
            SwitchToDocument(_vm.OpenDocuments.First(c =>
                string.Equals(c.FilePath, path, StringComparison.OrdinalIgnoreCase)));
            return;
        }

        // Close all existing documents first (single-document mode)
        var dirty = _vm.OpenDocuments.Where(c => c.IsDirty).ToList();
        if (dirty.Count > 0)
        {
            var names = string.Join("\n", dirty.Select(c => c.DisplayName));
            var result = MessageBox.Show(
                $"De volgende bestanden hebben niet-opgeslagen wijzigingen:\n{names}\n\nToch sluiten?",
                "Niet-opgeslagen wijzigingen", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes) return;
        }
        foreach (var existing in _vm.OpenDocuments.ToList())
            _vm.CloseDocument(existing);

        string? tempXml = null;
        try
        {
            string xmlPath, dstPath;
            if (path.EndsWith(".dst", StringComparison.OrdinalIgnoreCase))
            {
                tempXml = DstCodec.DecodeDstToTempXml(path);
                xmlPath = tempXml;
                dstPath = path;
            }
            else
            {
                xmlPath = path;
                dstPath = Path.ChangeExtension(path, ".dst");
            }

            var doc = _parser.ParseForEditing(dstPath, xmlPath);
            var ctx = _vm.AddDocument(doc);
            RecentFilesService.Add(path);
            RebuildTree();
            ShowNoSelection(ctx);
            FileDropdownButton.Content = _vm.FileTitle;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fout bij laden:\n{ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            if (tempXml != null && File.Exists(tempXml)) File.Delete(tempXml);
        }
    }

    private void CloseActiveDocument()
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        if (ctx.IsDirty)
        {
            var r = MessageBox.Show($"'{ctx.DisplayName}' heeft niet-opgeslagen wijzigingen.\nToch sluiten?",
                "Sluiten", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (r != MessageBoxResult.Yes) return;
        }

        _vm.CloseDocument(ctx);
        RebuildTree();
        if (_vm.ActiveDocument != null)
            ShowSelectionForDoc(_vm.ActiveDocument);
        else
            ShowNoSelection(null);

        FileDropdownButton.Content = _vm.FileTitle;
    }

    private void SwitchToDocument(DocumentContext ctx)
    {
        _vm.SetActive(ctx);
        RebuildTree();
        ShowSelectionForDoc(ctx);
        FileDropdownButton.Content = _vm.FileTitle;
    }

    private void ImportCsv()
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;
        var dlg = new OpenFileDialog
        {
            Title = "CSV of TSV importeren",
            Filter = "CSV/TSV bestanden (*.csv;*.tsv;*.txt)|*.csv;*.tsv;*.txt|Alle bestanden (*.*)|*.*"
        };
        if (dlg.ShowDialog() != true) return;
        try
        {
            var allSheets = ctx.Document.GetAllSheets().ToList();
            var updates = CsvTsvImporter.Import(dlg.FileName, allSheets);
            if (updates.Count == 0)
            {
                MessageBox.Show("Geen overeenkomende sheets gevonden.", "Geen updates",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            SheetSetWriter.Apply(updates);
            ctx.IsDirty = true;
            _vm.MarkDirty();
            RebuildTree();
            RefreshPropertyPanel(ctx);
            _vm.StatusText = $"{updates.Count} sheets bijgewerkt via CSV/TSV";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fout bij importeren:\n{ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ═══════════════════════════════════════════════════════════
    // Apply / Save
    // ═══════════════════════════════════════════════════════════

    private void ApplyButton_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        // Apply SheetSet title rename
        if (ctx.SelectedRoot == ctx.Document)
        {
            var newName = SheetSetNameBox?.Text.Trim() ?? string.Empty;
            if (newName.Length > 0 && newName != ctx.Document.Info.Name)
            {
                SheetSetWriter.RenameSheetSet(ctx.Document, newName);
                ctx.IsDirty = true;
                _vm.MarkDirty();
                RebuildTree();
                _vm.StatusText = $"Sheet set hernoemd naar '{newName}'";
                FileDropdownButton.Content = _vm.FileTitle;
            }
            return;
        }

        // Apply subset rename
        if (ctx.SelectedSubset != null)
        {
            var newName = SubsetNameBox?.Text.Trim() ?? string.Empty;
            if (newName.Length > 0 && newName != ctx.SelectedSubset.Name)
            {
                SheetSetWriter.RenameSubset(ctx.SelectedSubset, newName);
                ctx.IsDirty = true;
                _vm.MarkDirty();
                RebuildTree();
                _vm.StatusText = $"Subset hernoemd naar '{newName}'";
            }
            return;
        }

        // Apply sheet changes
        if (ctx.SelectedSheets.Count == 0) return;
        var updates = BuildSheetUpdates(ctx.SelectedSheets.ToList());
        if (updates.Count == 0) return;

        SheetSetWriter.Apply(updates);
        ctx.IsDirty = true;
        _vm.MarkDirty();
        RebuildTree();
        RefreshPropertyPanel(ctx);
        _vm.StatusText = $"{updates.Count} sheet(s) bijgewerkt";
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;
        try
        {
            SheetSetWriter.Save(ctx.Document);
            ctx.IsDirty = false;
            _vm.IsDirty = false;
            FileDropdownButton.Content = _vm.FileTitle;
            _vm.StatusText = $"Opgeslagen: {Path.GetFileName(ctx.FilePath)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fout bij opslaan:\n{ex.Message}", "Fout", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ═══════════════════════════════════════════════════════════
    // AutoNumber
    // ═══════════════════════════════════════════════════════════

    private void AutoNumberButton_Click(object sender, RoutedEventArgs e)
        => RunAutoNumber(1, string.Empty, string.Empty);

    private void OptionsButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new AutoNumberDialog { Owner = this };
        if (dlg.ShowDialog() != true) return;
        RunAutoNumber(dlg.StartNumber, dlg.Prefix, dlg.Suffix, dlg.Increment);
    }

    private void RunAutoNumber(int startNumber, string prefix, string suffix, int increment = 1)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null || ctx.SelectedSheets.Count == 0)
        {
            MessageBox.Show("Selecteer eerst één of meer sheets.", "AutoNumber",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (!int.TryParse(StartNumberBox.Text.Trim(), out int start))
            start = startNumber;

        var ordered = GetAllSheetNodesInOrder(ctx.Document).Where(ctx.SelectedSheets.Contains).ToList();
        int counter = start;
        var updates = ordered.Select(s => new SheetUpdateModel
        {
            Sheet = s.Info,
            NewNumber = $"{prefix}{counter++}{suffix}"
        }).ToList();

        SheetSetWriter.Apply(updates);
        ctx.IsDirty = true;
        _vm.MarkDirty();
        RebuildTree();
        RefreshPropertyPanel(ctx);
        _vm.StatusText = $"AutoNumber: {updates.Count} sheet(s) genummerd vanaf {start}";
    }

    // ═══════════════════════════════════════════════════════════
    // Tree building
    // ═══════════════════════════════════════════════════════════

    private void RebuildTree()
    {
        SheetTree.Items.Clear();
        foreach (var ctx in _vm.OpenDocuments)
            SheetTree.Items.Add(BuildDocumentRoot(ctx));
    }

    private TreeViewItem BuildDocumentRoot(DocumentContext ctx)
    {
        bool active = ctx == _vm.ActiveDocument;
        var root = new TreeViewItem
        {
            Tag = ctx,
            IsExpanded = true,
            Header = ctx.DisplayName + (ctx.IsDirty ? " *" : ""),
            FontWeight = FontWeights.Bold,
            Foreground = active
                ? new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3))
                : new SolidColorBrush(Color.FromRgb(0x88, 0x88, 0x88)),
            Background = active
                ? new SolidColorBrush(Color.FromArgb(25, 0x21, 0x96, 0xF3))
                : Brushes.Transparent
        };
        foreach (var subset in ctx.Document.RootSubsets)
            root.Items.Add(BuildSubsetItem(subset));
        foreach (var sheet in ctx.Document.RootSheets)
            root.Items.Add(BuildSheetItem(sheet));
        return root;
    }

    private static TreeViewItem BuildSubsetItem(SubsetNode subset)
    {
        var item = new TreeViewItem
        {
            Tag = subset, IsExpanded = true,
            Header = subset.Name, FontWeight = FontWeights.SemiBold
        };
        foreach (var sub in subset.Subsets)  item.Items.Add(BuildSubsetItem(sub));
        foreach (var sheet in subset.Sheets) item.Items.Add(BuildSheetItem(sheet));
        return item;
    }

    private static TreeViewItem BuildSheetItem(SheetNode sheet)
    {
        var label = (string.IsNullOrWhiteSpace(sheet.Number) && string.IsNullOrWhiteSpace(sheet.Title))
            ? "(zonder naam)"
            : $"{sheet.Number} – {sheet.Title}".Trim(' ', '–', ' ');
        return new TreeViewItem { Tag = sheet, Header = label };
    }

    private void RestoreHighlights()
    {
        var ctx = _vm.ActiveDocument;
        foreach (var item in FlattenTree(SheetTree))
        {
            bool sel = item.Tag is SheetNode s && ctx?.SelectedSheets.Contains(s) == true
                    || item.Tag is SubsetNode sub && sub == ctx?.SelectedSubset
                    || item.Tag is DocumentContext dc && dc == ctx && ctx?.SelectedRoot == ctx?.Document;

            item.Background = sel
                ? new SolidColorBrush(Color.FromArgb(0xFF, 0xBB, 0xDE, 0xF8))
                : item.Tag is DocumentContext activeCtx && activeCtx == _vm.ActiveDocument
                    ? new SolidColorBrush(Color.FromArgb(25, 0x21, 0x96, 0xF3))
                    : Brushes.Transparent;
        }
    }

    // ═══════════════════════════════════════════════════════════
    // Tree interaction: click + keyboard
    // ═══════════════════════════════════════════════════════════

    private void SheetTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (IsOnToggleButton(e.OriginalSource as DependencyObject)) return;
        var item = GetItemFromPoint(e.GetPosition(SheetTree));
        if (item == null) return;

        _dragStart = e.GetPosition(null);
        _isDragging = false;

        bool shift = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift);
        bool ctrl  = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
        HandleClick(item, shift, ctrl);
        SheetTree.Focus();
        e.Handled = true;
    }

    private void SheetTree_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key is not (Key.Up or Key.Down)) return;
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        var nav = FlattenTree(SheetTree)
            .Where(i => i.Tag is SheetNode or SubsetNode or DocumentContext)
            .ToList();
        var current = ctx.LastClickedItem;
        if (current == null && nav.Count > 0) current = nav[0];
        if (current == null) return;

        int idx = nav.IndexOf(current);
        int next = e.Key == Key.Up ? idx - 1 : idx + 1;
        if (next < 0 || next >= nav.Count) return;

        bool shift = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift);
        bool ctrl  = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
        HandleClick(nav[next], shift, ctrl);
        nav[next].BringIntoView();
        e.Handled = true;
    }

    private void HandleClick(TreeViewItem item, bool shift, bool ctrl)
    {
        // Clicking a document root
        if (item.Tag is DocumentContext docCtx)
        {
            SwitchToDocument(docCtx);
            var ctx2 = _vm.ActiveDocument!;
            ctx2.SelectedSheets.Clear();
            ctx2.SelectedSubset = null;
            ctx2.SelectedRoot = docCtx.Document;
            ctx2.LastClickedItem = item;
            RestoreHighlights();
            ShowSheetSetProperties(ctx2);
            return;
        }

        // Clicking a subset
        if (item.Tag is SubsetNode subset)
        {
            var ctx = EnsureActiveForItem(item);
            ctx.SelectedSheets.Clear();
            ctx.SelectedSubset = subset;
            ctx.SelectedRoot = null;
            ctx.LastClickedItem = item;
            RestoreHighlights();
            ShowSubsetProperties(subset);
            return;
        }

        // Clicking a sheet
        if (item.Tag is SheetNode clickedSheet)
        {
            var ctx = EnsureActiveForItem(item);
            ctx.SelectedRoot = null;
            ctx.SelectedSubset = null;

            var allSheetItems = FlattenTree(SheetTree).Where(i => i.Tag is SheetNode).ToList();

            if (shift && ctx.LastClickedItem != null && ctx.LastClickedItem.Tag is SheetNode)
            {
                int start = allSheetItems.IndexOf(ctx.LastClickedItem);
                int end   = allSheetItems.IndexOf(item);
                if (start < 0) { start = end; }
                if (start > end) (start, end) = (end, start);
                ctx.SelectedSheets.Clear();
                for (int i = start; i <= end; i++)
                    if (allSheetItems[i].Tag is SheetNode s) ctx.SelectedSheets.Add(s);
            }
            else if (ctrl)
            {
                if (!ctx.SelectedSheets.Remove(clickedSheet)) ctx.SelectedSheets.Add(clickedSheet);
            }
            else
            {
                ctx.SelectedSheets.Clear();
                ctx.SelectedSheets.Add(clickedSheet);
            }

            ctx.LastClickedItem = item;
            RestoreHighlights();
            ShowSheetProperties(ctx.SelectedSheets.ToList());
        }
    }

    private DocumentContext EnsureActiveForItem(TreeViewItem item)
    {
        // Walk up to find which document this item belongs to
        foreach (var ctx in _vm.OpenDocuments)
        {
            var rootItem = FlattenTree(SheetTree).FirstOrDefault(i => i.Tag is DocumentContext dc && dc == ctx);
            if (rootItem != null && IsDescendantOf(item, rootItem))
            {
                if (ctx != _vm.ActiveDocument)
                {
                    _vm.SetActive(ctx);
                    RebuildTree();
                    FileDropdownButton.Content = _vm.FileTitle;
                }
                return ctx;
            }
        }
        return _vm.ActiveDocument ?? _vm.OpenDocuments.First();
    }

    private static bool IsDescendantOf(TreeViewItem item, TreeViewItem potentialAncestor)
    {
        DependencyObject? current = item;
        while (current != null)
        {
            if (current == potentialAncestor) return true;
            current = VisualTreeHelper.GetParent(current);
        }
        return false;
    }

    // ═══════════════════════════════════════════════════════════
    // Drag & drop
    // ═══════════════════════════════════════════════════════════

    private void SheetTree_PreviewMouseMove(object sender, MouseEventArgs e)
    {
        if (e.LeftButton != MouseButtonState.Pressed || _isDragging) return;
        var pos = e.GetPosition(null);
        if (Math.Abs(pos.X - _dragStart.X) < SystemParameters.MinimumHorizontalDragDistance &&
            Math.Abs(pos.Y - _dragStart.Y) < SystemParameters.MinimumVerticalDragDistance) return;

        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        var item = GetItemFromPoint(e.GetPosition(SheetTree));
        if (item == null) return;
        if (item.Tag is not (SheetNode or SubsetNode)) return;

        _isDragging = true;
        _dragPayload = item.Tag;
        _dragSourceCtx = ctx;
        DragDrop.DoDragDrop(item, item.Tag, DragDropEffects.Move);
        _isDragging = false;
        _dragPayload = null;
        ClearDropHighlight();
    }

    private void SheetTree_DragOver(object sender, DragEventArgs e)
    {
        var target = GetItemFromPoint(e.GetPosition(SheetTree));
        if (target == null || _dragPayload == null) { e.Effects = DragDropEffects.None; return; }

        // Can't drop on itself
        if (target.Tag == _dragPayload) { e.Effects = DragDropEffects.None; return; }

        // SubsetNode can only be dropped on another subset or document root
        if (_dragPayload is SubsetNode && target.Tag is SheetNode) { e.Effects = DragDropEffects.None; return; }

        e.Effects = DragDropEffects.Move;
        HighlightDropTarget(target);
        e.Handled = true;
    }

    private void SheetTree_Drop(object sender, DragEventArgs e)
    {
        ClearDropHighlight();
        var ctx = _dragSourceCtx;
        if (ctx == null || _dragPayload == null) return;

        var target = GetItemFromPoint(e.GetPosition(SheetTree));
        if (target == null || target.Tag == _dragPayload) return;

        var doc = ctx.Document;

        if (_dragPayload is SheetNode draggedSheet)
        {
            if (target.Tag is SheetNode targetSheet)
            {
                // Move sheet before targetSheet
                var srcList = FindSheetList(doc, draggedSheet);
                var tgtList = FindSheetList(doc, targetSheet);
                if (srcList == null || tgtList == null) return;

                SheetSetWriter.MoveElementBefore(draggedSheet.Info.Element!, targetSheet.Info.Element!);
                srcList.Remove(draggedSheet);
                tgtList.Insert(tgtList.IndexOf(targetSheet), draggedSheet);
            }
            else if (target.Tag is SubsetNode targetSubset)
            {
                // Move sheet into subset (append)
                var srcList = FindSheetList(doc, draggedSheet);
                if (srcList == null) return;

                SheetSetWriter.MoveElementInto(draggedSheet.Info.Element!, targetSubset.Info.Element!);
                srcList.Remove(draggedSheet);
                targetSubset.Sheets.Add(draggedSheet);
            }
            else if (target.Tag is DocumentContext)
            {
                // Move to root (append)
                var srcList = FindSheetList(doc, draggedSheet);
                if (srcList == null) return;

                SheetSetWriter.MoveElementInto(draggedSheet.Info.Element!, doc.Info.Element!);
                srcList.Remove(draggedSheet);
                doc.RootSheets.Add(draggedSheet);
            }
        }
        else if (_dragPayload is SubsetNode draggedSubset)
        {
            if (target.Tag is SubsetNode targetSubset)
            {
                // Move subset before targetSubset
                var srcList = FindSubsetList(doc, draggedSubset);
                var tgtList = FindSubsetList(doc, targetSubset);
                if (srcList == null || tgtList == null) return;

                SheetSetWriter.MoveElementBefore(draggedSubset.Info.Element!, targetSubset.Info.Element!);
                srcList.Remove(draggedSubset);
                tgtList.Insert(tgtList.IndexOf(targetSubset), draggedSubset);
            }
            else if (target.Tag is DocumentContext)
            {
                // Move subset to root
                var srcList = FindSubsetList(doc, draggedSubset);
                if (srcList == null) return;

                SheetSetWriter.MoveElementInto(draggedSubset.Info.Element!, doc.Info.Element!);
                srcList.Remove(draggedSubset);
                doc.RootSubsets.Add(draggedSubset);
            }
        }

        ctx.IsDirty = true;
        _vm.MarkDirty();
        RebuildTree();
    }

    private void HighlightDropTarget(TreeViewItem item)
    {
        ClearDropHighlight();
        _dropHighlightedItem = item;
        item.Background = new SolidColorBrush(Color.FromArgb(80, 0x21, 0x96, 0xF3));
    }

    private void ClearDropHighlight()
    {
        if (_dropHighlightedItem != null)
        {
            _dropHighlightedItem.Background = Brushes.Transparent;
            _dropHighlightedItem = null;
        }
        RestoreHighlights();
    }

    // ═══════════════════════════════════════════════════════════
    // Context menu
    // ═══════════════════════════════════════════════════════════

    private void TreeContextMenu_Opened(object sender, RoutedEventArgs e)
    {
        _contextTargetItem = GetItemFromPoint(Mouse.GetPosition(SheetTree));
        bool isSubset    = _contextTargetItem?.Tag is SubsetNode;
        bool isSheet     = _contextTargetItem?.Tag is SheetNode;
        bool isAddTarget = isSubset || isSheet || _contextTargetItem?.Tag is DocumentContext;

        CtxNewSubset.IsEnabled  = isAddTarget;
        CtxRenameItem.IsEnabled = isSubset || isSheet;
        CtxDeleteItem.IsEnabled = isSubset;
    }

    private TreeViewItem? _contextTargetItem;

    private void CtxNewSubset_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        string? name = PromptInput("Nieuwe subset", "Naam van de nieuwe subset:");
        if (string.IsNullOrWhiteSpace(name)) return;

        SubsetNode? parentSubset = _contextTargetItem?.Tag switch
        {
            SubsetNode s    => s,
            SheetNode sheet => FindParentSubset(ctx.Document, sheet),
            _ => null
        };

        SheetSetWriter.AddSubset(ctx.Document, parentSubset, name);
        ctx.IsDirty = true;
        _vm.MarkDirty();
        RebuildTree();
        _vm.StatusText = $"Subset '{name}' toegevoegd";
    }

    private void CtxRenameItem_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        if (_contextTargetItem?.Tag is SubsetNode subset)
        {
            string? name = PromptInput("Hernaam subset", "Nieuwe naam:", subset.Name);
            if (string.IsNullOrWhiteSpace(name) || name == subset.Name) return;
            SheetSetWriter.RenameSubset(subset, name);
            ctx.IsDirty = true;
            _vm.MarkDirty();
            RebuildTree();
            _vm.StatusText = $"Subset hernoemd naar '{name}'";
        }
        else if (_contextTargetItem?.Tag is SheetNode sheet)
        {
            string? title = PromptInput("Hernaam sheet", "Nieuwe titel:", sheet.Info.Title ?? string.Empty);
            if (title == null || title == sheet.Info.Title) return;
            SheetSetWriter.RenameSheet(sheet, title);
            ctx.IsDirty = true;
            _vm.MarkDirty();
            RebuildTree();
            _vm.StatusText = $"Sheet titel gewijzigd naar '{title}'";
        }
    }

    private void CtxDeleteItem_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null || _contextTargetItem?.Tag is not SubsetNode subset) return;

        var r = MessageBox.Show($"Subset '{subset.Name}' verwijderen inclusief alle inhoud?",
            "Verwijderen", MessageBoxButton.YesNo, MessageBoxImage.Warning);
        if (r != MessageBoxResult.Yes) return;

        subset.Info.Element?.Remove();
        RemoveSubsetFromModel(ctx.Document, subset);
        ctx.IsDirty = true;
        _vm.MarkDirty();
        RebuildTree();
        _vm.StatusText = $"Subset '{subset.Name}' verwijderd";
    }

    // ═══════════════════════════════════════════════════════════
    // Property panels
    // ═══════════════════════════════════════════════════════════

    private TextBox? SheetSetNameBox;  // stored ref for Apply

    private void ShowNoSelection(DocumentContext? ctx)
    {
        NoSelectionLabel.Visibility      = Visibility.Visible;
        SubsetPanel.Visibility           = Visibility.Collapsed;
        HoofdveldenPanel.Visibility      = Visibility.Collapsed;
        OverigeVeldenExpander.Visibility = Visibility.Collapsed;
        CustomPropsExpander.Visibility   = Visibility.Collapsed;
        HideSheetSetPanel();
    }

    private void ShowSheetSetProperties(DocumentContext ctx)
    {
        NoSelectionLabel.Visibility      = Visibility.Collapsed;
        SubsetPanel.Visibility           = Visibility.Collapsed;
        HoofdveldenPanel.Visibility      = Visibility.Collapsed;
        OverigeVeldenExpander.Visibility = Visibility.Collapsed;
        CustomPropsExpander.Visibility   = Visibility.Collapsed;
        ShowSheetSetPanel(ctx.Document.Info.Name ?? string.Empty);
    }

    private void ShowSubsetProperties(SubsetNode subset)
    {
        NoSelectionLabel.Visibility      = Visibility.Collapsed;
        SubsetPanel.Visibility           = Visibility.Visible;
        HoofdveldenPanel.Visibility      = Visibility.Collapsed;
        OverigeVeldenExpander.Visibility = Visibility.Collapsed;
        CustomPropsExpander.Visibility   = Visibility.Collapsed;
        HideSheetSetPanel();
        SubsetNameBox.Text = subset.Name;
    }

    private void ShowSheetProperties(List<SheetNode> sheets)
    {
        if (sheets.Count == 0) { ShowNoSelection(_vm.ActiveDocument); return; }

        NoSelectionLabel.Visibility      = Visibility.Collapsed;
        SubsetPanel.Visibility           = Visibility.Collapsed;
        HoofdveldenPanel.Visibility      = Visibility.Visible;
        OverigeVeldenExpander.Visibility = Visibility.Visible;
        CustomPropsExpander.Visibility   = Visibility.Visible;
        HideSheetSetPanel();

        if (sheets.Count == 1)
        {
            var info = sheets[0].Info;
            TbNumber.Text       = info.Number         ?? string.Empty;
            TbTitle.Text        = info.Title          ?? string.Empty;
            TbDesc.Text         = info.Description    ?? string.Empty;
            TbRevNr.Text        = info.RevisionNumber ?? string.Empty;
            TbRevDate.Text      = info.RevisionDate   ?? string.Empty;
            TbIssuePurpose.Text = info.IssuePurpose   ?? string.Empty;
            TbCategory.Text     = info.Category       ?? string.Empty;
            CbDoNotPlot.IsChecked = info.DoNotPlot;
            CustomPropsPanel.ItemsSource = info.CustomProperties
                .Select(kv => new CustomPropRow(kv.Key, kv.Value)).ToList();
        }
        else
        {
            TbNumber.Text       = Mixed(sheets, i => i.Number);
            TbTitle.Text        = Mixed(sheets, i => i.Title);
            TbDesc.Text         = Mixed(sheets, i => i.Description);
            TbRevNr.Text        = Mixed(sheets, i => i.RevisionNumber);
            TbRevDate.Text      = Mixed(sheets, i => i.RevisionDate);
            TbIssuePurpose.Text = Mixed(sheets, i => i.IssuePurpose);
            TbCategory.Text     = Mixed(sheets, i => i.Category);
            CbDoNotPlot.IsChecked = MixedBool(sheets, i => i.DoNotPlot);

            // Show custom props common to ALL selected sheets; [gedeeld] when values differ
            var keysets = sheets
                .Select(s => s.Info.CustomProperties.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase))
                .ToList();
            var commonKeys = keysets.Aggregate((a, b) => { a.IntersectWith(b); return a; });
            var multiRows = commonKeys.OrderBy(k => k).Select(key =>
            {
                var distinct = sheets
                    .Select(s => s.Info.CustomProperties.TryGetValue(key, out var v) ? v : null)
                    .Distinct().ToList();
                return new CustomPropRow(key, distinct.Count == 1 ? distinct[0] ?? string.Empty : "[gedeeld]");
            }).ToList();
            CustomPropsPanel.ItemsSource = multiRows.Count > 0 ? multiRows : null;
        }
    }

    // Dynamic SheetSet panel (injected into ScrollViewer StackPanel)
    private Border? _sheetSetPanelBorder;

    private void ShowSheetSetPanel(string currentName)
    {
        HideSheetSetPanel();
        var tb = new TextBox
        {
            Text = currentName,
            Height = 24, Padding = new Thickness(3, 0, 3, 0),
            VerticalContentAlignment = VerticalAlignment.Center
        };
        SheetSetNameBox = tb;
        var grid = new Grid { Margin = new Thickness(8) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        var label = new TextBlock
        {
            Text = "Naam:", VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 6, 0), Foreground = new SolidColorBrush(Color.FromRgb(0x44, 0x44, 0x44))
        };
        Grid.SetColumn(label, 0); Grid.SetColumn(tb, 1);
        grid.Children.Add(label); grid.Children.Add(tb);

        var header = new GroupBox
        {
            Header = "Sheet Set",
            Margin = new Thickness(12, 0, 12, 10),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3)),
            BorderThickness = new Thickness(1.5),
            Padding = new Thickness(8),
            Content = grid
        };
        _sheetSetPanelBorder = new Border { Child = header };

        // Insert before NoSelectionLabel in PropertyStack
        PropertyStack.Children.Insert(0, _sheetSetPanelBorder);
        NoSelectionLabel.Visibility = Visibility.Collapsed;
    }

    private void HideSheetSetPanel()
    {
        if (_sheetSetPanelBorder != null)
        {
            PropertyStack.Children.Remove(_sheetSetPanelBorder);
            _sheetSetPanelBorder = null;
            SheetSetNameBox = null;
        }
    }

    private void ShowSelectionForDoc(DocumentContext ctx)
    {
        if (ctx.SelectedRoot == ctx.Document)   ShowSheetSetProperties(ctx);
        else if (ctx.SelectedSubset != null)    ShowSubsetProperties(ctx.SelectedSubset);
        else if (ctx.SelectedSheets.Count > 0)  ShowSheetProperties(ctx.SelectedSheets.ToList());
        else                                    ShowNoSelection(ctx);
    }

    private void RefreshPropertyPanel(DocumentContext ctx) => ShowSelectionForDoc(ctx);

    // ═══════════════════════════════════════════════════════════
    // Build updates from editor
    // ═══════════════════════════════════════════════════════════

    private List<SheetUpdateModel> BuildSheetUpdates(List<SheetNode> sheets)
    {
        var customProps = CustomPropsPanel.ItemsSource?
            .OfType<CustomPropRow>()
            .ToDictionary(r => r.Key, r => r.Value);

        return sheets.Select(s => new SheetUpdateModel
        {
            Sheet          = s.Info,
            NewNumber      = NullIfMixed(TbNumber.Text),
            Title          = NullIfMixed(TbTitle.Text),
            Description    = NullIfMixed(TbDesc.Text),
            RevisionNumber = NullIfMixed(TbRevNr.Text),
            RevisionDate   = NullIfMixed(TbRevDate.Text),
            IssuePurpose   = NullIfMixed(TbIssuePurpose.Text),
            Category       = NullIfMixed(TbCategory.Text),
            DoNotPlot      = CbDoNotPlot.IsChecked,
            CustomProperties = customProps?
                .Where(kv => kv.Value != "[gedeeld]")
                .ToDictionary(kv => kv.Key, kv => kv.Value)
        }).ToList();
    }

    private static string? NullIfMixed(string value)
        => value == "[gedeeld]" ? null : value;

    // ═══════════════════════════════════════════════════════════
    // Helpers
    // ═══════════════════════════════════════════════════════════

    private static string Mixed(List<SheetNode> sheets, Func<SheetInfo, string?> getter)
    {
        var dist = sheets.Select(s => getter(s.Info)).Distinct().ToList();
        return dist.Count == 1 ? dist[0] ?? string.Empty : "[gedeeld]";
    }

    private static bool? MixedBool(List<SheetNode> sheets, Func<SheetInfo, bool?> getter)
    {
        var dist = sheets.Select(s => getter(s.Info)).Distinct().ToList();
        return dist.Count == 1 ? dist[0] : null;
    }

    private static IEnumerable<SheetNode> GetAllSheetNodesInOrder(SheetSetDocument doc)
    {
        foreach (var s in doc.RootSheets) yield return s;
        foreach (var sub in doc.RootSubsets)
            foreach (var s in GetSheetsFromSubset(sub)) yield return s;
    }

    private static IEnumerable<SheetNode> GetSheetsFromSubset(SubsetNode subset)
    {
        foreach (var s in subset.Sheets) yield return s;
        foreach (var sub in subset.Subsets)
            foreach (var s in GetSheetsFromSubset(sub)) yield return s;
    }

    private static SubsetNode? FindParentSubset(SheetSetDocument doc, SheetNode sheet)
    {
        foreach (var sub in doc.RootSubsets)
        {
            var found = FindParentSubsetInSubset(sub, sheet);
            if (found != null) return found;
        }
        return null;
    }

    private static SubsetNode? FindParentSubsetInSubset(SubsetNode parent, SheetNode sheet)
    {
        if (parent.Sheets.Contains(sheet)) return parent;
        foreach (var sub in parent.Subsets)
        {
            var found = FindParentSubsetInSubset(sub, sheet);
            if (found != null) return found;
        }
        return null;
    }

    private static List<SheetNode>? FindSheetList(SheetSetDocument doc, SheetNode sheet)
    {
        if (doc.RootSheets.Contains(sheet)) return doc.RootSheets;
        foreach (var sub in doc.RootSubsets)
        {
            var found = FindSheetListInSubset(sub, sheet);
            if (found != null) return found;
        }
        return null;
    }

    private static List<SheetNode>? FindSheetListInSubset(SubsetNode subset, SheetNode sheet)
    {
        if (subset.Sheets.Contains(sheet)) return subset.Sheets;
        foreach (var sub in subset.Subsets)
        {
            var found = FindSheetListInSubset(sub, sheet);
            if (found != null) return found;
        }
        return null;
    }

    private static List<SubsetNode>? FindSubsetList(SheetSetDocument doc, SubsetNode subset)
    {
        if (doc.RootSubsets.Contains(subset)) return doc.RootSubsets;
        foreach (var sub in doc.RootSubsets)
        {
            var found = FindSubsetListInSubset(sub, subset);
            if (found != null) return found;
        }
        return null;
    }

    private static List<SubsetNode>? FindSubsetListInSubset(SubsetNode parent, SubsetNode target)
    {
        if (parent.Subsets.Contains(target)) return parent.Subsets;
        foreach (var sub in parent.Subsets)
        {
            var found = FindSubsetListInSubset(sub, target);
            if (found != null) return found;
        }
        return null;
    }

    private static void RemoveSubsetFromModel(SheetSetDocument doc, SubsetNode target)
    {
        if (doc.RootSubsets.Remove(target)) return;
        foreach (var sub in doc.RootSubsets)
            if (RemoveSubsetFromSubset(sub, target)) return;
    }

    private static bool RemoveSubsetFromSubset(SubsetNode parent, SubsetNode target)
    {
        if (parent.Subsets.Remove(target)) return true;
        foreach (var sub in parent.Subsets)
            if (RemoveSubsetFromSubset(sub, target)) return true;
        return false;
    }

    private static List<TreeViewItem> FlattenTree(ItemsControl parent)
    {
        var result = new List<TreeViewItem>();
        foreach (var item in parent.Items)
        {
            if (parent.ItemContainerGenerator.ContainerFromItem(item) is TreeViewItem tvi)
            {
                result.Add(tvi);
                result.AddRange(FlattenTree(tvi));
            }
        }
        return result;
    }

    private TreeViewItem? GetItemFromPoint(Point point)
    {
        var hit = SheetTree.InputHitTest(point) as DependencyObject;
        while (hit != null && hit is not TreeViewItem)
            hit = VisualTreeHelper.GetParent(hit);
        return hit as TreeViewItem;
    }

    private static bool IsOnToggleButton(DependencyObject? el)
    {
        while (el != null)
        {
            if (el is System.Windows.Controls.Primitives.ToggleButton) return true;
            el = VisualTreeHelper.GetParent(el);
        }
        return false;
    }

    private string? PromptInput(string title, string prompt, string defaultValue = "")
    {
        var dlg = new InputDialog(title, prompt, defaultValue) { Owner = this };
        return dlg.ShowDialog() == true ? dlg.Value : null;
    }
}

// ─── CustomPropRow ────────────────────────────────────────────────────────────

public sealed class CustomPropRow(string key, string value)
{
    public string Key   { get; } = key;
    public string Value { get; set; } = value;
}
