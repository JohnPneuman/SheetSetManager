using System.Collections;
using System.Collections.ObjectModel;
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

    // Undo stack for custom-property deletions
    private sealed record DeletedPropUndo(
        DocumentContext Ctx,
        CustomPropertyDefinition Def,
        Dictionary<SheetInfo, string> SheetValues);
    private readonly Stack<DeletedPropUndo> _deletedPropStack = new();

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
        AddMenuItem(menu, "Nieuwe sheet set…",    (_, _) => NewSheetSet());
        AddMenuItem(menu, "Openen…",              (_, _) => OpenFile());
        AddMenuItem(menu, "CSV/TSV importeren…",  (_, _) => OpenImportWizard(), _vm.IsDocumentLoaded);
        AddMenuItem(menu, "Exporteer naar CSV…", (_, _) => ExportToCsv(),     _vm.IsDocumentLoaded);
        AddMenuItem(menu, "Ongedaan maken (import)", (_, _) => UndoLastImport(),
            _vm.IsDocumentLoaded && _vm.ActiveDocument?.LastImportSnapshot != null);
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

    private void NewSheetSet()
    {
        var wizard = new NewSheetSetWizard { Owner = this };
        if (wizard.ShowDialog() == true)
            LoadFile(wizard.OutputPath);
    }

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

    private void OpenImportWizard()
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;
        OpenImportWizardForContext(ctx);
    }

    private void OpenImportWizardForContext(DocumentContext ctx)
    {
        var wizard = new ImportWizard.ImportWizardWindow(ctx) { Owner = this };
        wizard.ShowDialog();

        if (ctx.IsDirty)
        {
            _vm.MarkDirty();
            RebuildTree();
            RefreshPropertyPanel(ctx);
        }
    }

    // ═══════════════════════════════════════════════════════════
    // Export / Undo
    // ═══════════════════════════════════════════════════════════

    private void ExportToCsv()
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        var dlg = new Microsoft.Win32.SaveFileDialog
        {
            Title            = "Exporteer naar CSV",
            Filter           = "CSV-bestanden (*.csv)|*.csv",
            DefaultExt       = ".csv",
            FileName         = ctx.DisplayName
        };
        if (dlg.ShowDialog() != true) return;

        try
        {
            SheetSet.Core.Import.Export.CsvExporter.ExportToFile(ctx.Document, dlg.FileName);
            _vm.StatusText = $"Geëxporteerd naar {System.IO.Path.GetFileName(dlg.FileName)}";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fout bij exporteren:\n{ex.Message}", "Fout",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void UndoLastImport()
    {
        var ctx  = _vm.ActiveDocument;
        var snap = ctx?.LastImportSnapshot;
        if (snap == null) return;

        var confirm = MessageBox.Show(
            $"Import van \"{snap.SourceDescription}\" ({snap.Timestamp:HH:mm:ss}) ongedaan maken?",
            "Ongedaan maken", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;

        foreach (var s in snap.Sheets)
        {
            if (s.Sheet.Element == null) continue;

            // Restore XML content without replacing the element reference
            s.Sheet.Element.RemoveAll();
            foreach (var attr in s.OriginalElement.Attributes())
                s.Sheet.Element.Add(new System.Xml.Linq.XAttribute(attr));
            foreach (var child in s.OriginalElement.Elements())
                s.Sheet.Element.Add(new System.Xml.Linq.XElement(child));

            s.Sheet.CustomProperties.Clear();
            foreach (var kv in s.OriginalCustomProperties)
                s.Sheet.CustomProperties[kv.Key] = kv.Value;
        }

        ctx!.LastImportSnapshot = null;
        ctx.IsDirty             = true;
        _vm.MarkDirty();
        RefreshPropertyPanel(ctx);
        _vm.StatusText = $"Import ongedaan gemaakt ({snap.SourceDescription})";
    }

    private void UndoPropDelete_Click(object sender, RoutedEventArgs e)
    {
        if (!_deletedPropStack.TryPop(out var entry)) return;
        var doc = entry.Ctx.Document;
        var def = entry.Def;

        // Restore definition in the XML bag and the in-memory model
        if (doc.Info.Element != null)
            XmlUtil.SetCustomProperty(doc.Info.Element, def.Name, def.Value, def.Flags);

        if (!doc.Info.CustomPropertyDefinitions.Any(d =>
                d.Name.Equals(def.Name, StringComparison.OrdinalIgnoreCase)))
            doc.Info.CustomPropertyDefinitions.Add(def);

        doc.Info.CustomProperties[def.Name] = def.Value;

        // Restore per-sheet values (Flags=2 properties)
        foreach (var (sheet, value) in entry.SheetValues)
        {
            if (sheet.Element != null)
                XmlUtil.SetCustomProperty(sheet.Element, def.Name, value);
            sheet.CustomProperties[def.Name] = value;
        }

        entry.Ctx.IsDirty = true;
        _vm.MarkDirty();
        RefreshPropertyPanel(entry.Ctx);
        UndoPropDeleteButton.IsEnabled = _deletedPropStack.Count > 0;
        _vm.StatusText = $"Verwijdering van '{def.Name}' ongedaan gemaakt.";
    }

    private void PushDeletedProp(DocumentContext ctx, CustomPropertyDefinition def)
    {
        var sheetValues = new Dictionary<SheetInfo, string>();
        if (def.Flags == 2)
        {
            foreach (var sheet in ctx.Document.GetAllSheets())
            {
                if (sheet.Info.CustomProperties.TryGetValue(def.Name, out var v))
                    sheetValues[sheet.Info] = v;
            }
        }
        _deletedPropStack.Push(new DeletedPropUndo(ctx, def, sheetValues));
        UndoPropDeleteButton.IsEnabled = true;
    }

    // ═══════════════════════════════════════════════════════════
    // Apply / Save
    // ═══════════════════════════════════════════════════════════

    private void ApplyButton_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        // Apply SheetSet properties
        if (ctx.SelectedRoot == ctx.Document)
        {
            var newName = SheetSetNameBox?.Text.Trim() ?? string.Empty;
            bool nameChanged = newName.Length > 0 && newName != ctx.Document.Info.Name;
            if (nameChanged)
                SheetSetWriter.RenameSheetSet(ctx.Document, newName);

            SheetSetWriter.UpdateSheetSetInfo(ctx.Document,
                SheetSetDescBox?.Text.Trim(),
                SheetSetProjectNameBox?.Text.Trim(),
                SheetSetProjectNumberBox?.Text.Trim(),
                SheetSetProjectPhaseBox?.Text.Trim(),
                SheetSetProjectMilestoneBox?.Text.Trim());

            var sheetSetEl2 = ctx.Document.Info.Element!;
            if (_sheetSetPropsCtrl?.ItemsSource is IEnumerable<SheetSetCustomPropRow> setRows)
            {
                foreach (var row in setRows)
                {
                    XmlUtil.SetCustomProperty(sheetSetEl2, row.Key, row.Value, flags: 1);
                    ctx.Document.Info.CustomProperties[row.Key] = row.Value;
                    var def = ctx.Document.Info.CustomPropertyDefinitions
                        .FirstOrDefault(d => string.Equals(d.Name, row.Key, StringComparison.OrdinalIgnoreCase));
                    if (def != null) def.Value = row.Value;
                }
            }
            if (_bladVeldenCtrl?.ItemsSource is IEnumerable<SheetSetCustomPropRow> bladRows)
            {
                foreach (var row in bladRows)
                {
                    XmlUtil.SetCustomProperty(sheetSetEl2, row.Key, row.Value, flags: 2);
                    ctx.Document.Info.CustomProperties[row.Key] = row.Value;
                    var def = ctx.Document.Info.CustomPropertyDefinitions
                        .FirstOrDefault(d => string.Equals(d.Name, row.Key, StringComparison.OrdinalIgnoreCase));
                    if (def != null) def.Value = row.Value;
                }
            }

            ctx.IsDirty = true;
            _vm.MarkDirty();
            if (nameChanged)
            {
                RebuildTree();
                FileDropdownButton.Content = _vm.FileTitle;
            }
            _vm.StatusText = "Sheet set bijgewerkt";
            return;
        }

        // Apply subset properties
        if (ctx.SelectedSubset != null)
        {
            var newName = SubsetNameBox?.Text.Trim() ?? string.Empty;
            bool nameChanged = newName.Length > 0 && newName != ctx.SelectedSubset.Name;
            if (nameChanged)
                SheetSetWriter.RenameSubset(ctx.SelectedSubset, newName);

            SheetSetWriter.UpdateSubsetInfo(ctx.SelectedSubset, SubsetDescBox?.Text.Trim());

            ctx.IsDirty = true;
            _vm.MarkDirty();
            if (nameChanged)
                RebuildTree();
            _vm.StatusText = $"Subset '{ctx.SelectedSubset.Name}' bijgewerkt";
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

            if (IsAutoCADRunning())
            {
                MessageBox.Show(
                    $"'{Path.GetFileName(ctx.FilePath)}' is opgeslagen.\n\n" +
                    "AutoCAD staat aan en heeft deze Sheet Set mogelijk in het geheugen.\n" +
                    "Herlaad de Sheet Set in AutoCAD om overschrijven te voorkomen:\n\n" +
                    "→ Rechtermuisknop op de Sheet Set → 'Sheet Set sluiten'\n" +
                    "→ Daarna opnieuw openen",
                    "Herlaad in AutoCAD",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fout bij opslaan:\n{ex.Message}", "Opslaan mislukt", MessageBoxButton.OK, MessageBoxImage.Error);
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

        var ctx = _vm.ActiveDocument;
        var selectedCount = ctx?.SelectedSheets.Count ?? 0;

        CtxNewSubset.IsEnabled  = isAddTarget;
        CtxNewSheet.IsEnabled   = isAddTarget;
        CtxRenameItem.IsEnabled = isSubset || isSheet;
        CtxDeleteItem.IsEnabled = isSubset || isSheet;

        // Enable import when at least one sheet is selected (either the right-clicked one or the multi-selection)
        CtxImportCsv.IsEnabled  = ctx != null && (selectedCount > 0 || isSheet);
        CtxImportCsv.Header     = selectedCount > 1
            ? $"CSV/TSV importeren op {selectedCount} geselecteerde sheets…"
            : "CSV/TSV importeren op selectie…";
    }

    private void CtxImportCsv_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        // If the right-clicked item is a sheet that isn't part of the multi-selection,
        // treat it as the sole target so the user gets what they right-clicked on.
        if (_contextTargetItem?.Tag is SheetNode clickedSheet &&
            !ctx.SelectedSheets.Any(s => s.Info == clickedSheet.Info))
        {
            // Temporarily scope to just this sheet by using the wizard with a one-item list
            var tempCtx = new DocumentContext(ctx.Document) { IsDirty = ctx.IsDirty };
            tempCtx.SelectedSheets.Add(clickedSheet);
            OpenImportWizardForContext(tempCtx);
            if (tempCtx.IsDirty) ctx.IsDirty = true;
            return;
        }

        OpenImportWizardForContext(ctx);
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

    private void CtxNewSheet_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        string? number = PromptInput("Nieuw blad", "Bladnummer:");
        if (string.IsNullOrWhiteSpace(number)) return;
        string? title = PromptInput("Nieuw blad", "Bladtitel:");
        if (title == null) return;

        SubsetNode? parentSubset = _contextTargetItem?.Tag switch
        {
            SubsetNode s    => s,
            SheetNode sheet => FindParentSubset(ctx.Document, sheet),
            _               => null
        };

        SheetSetWriter.AddSheet(ctx.Document, parentSubset, number, title);
        ctx.IsDirty = true;
        _vm.MarkDirty();
        RebuildTree();
        _vm.StatusText = $"Blad '{number} – {title}' toegevoegd";
    }

    private void CtxDeleteItem_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        if (_contextTargetItem?.Tag is SubsetNode subset)
        {
            var r = MessageBox.Show($"Subset '{subset.Name}' verwijderen inclusief alle inhoud?",
                "Verwijderen", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (r != MessageBoxResult.Yes) return;

            subset.Info.Element?.Remove();
            RemoveSubsetFromModel(ctx.Document, subset);
            ctx.IsDirty = true;
            _vm.MarkDirty();
            RebuildTree();
            ShowNoSelection(ctx);
            _vm.StatusText = $"Subset '{subset.Name}' verwijderd";
        }
        else if (_contextTargetItem?.Tag is SheetNode sheet)
        {
            var label = $"{sheet.Info.Number} – {sheet.Info.Title}".Trim(' ', '–', ' ');
            var r = MessageBox.Show($"Blad '{label}' verwijderen?",
                "Verwijderen", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (r != MessageBoxResult.Yes) return;

            SheetSetWriter.DeleteSheet(ctx.Document, sheet);
            ctx.SelectedSheets.Remove(sheet);
            ctx.IsDirty = true;
            _vm.MarkDirty();
            RebuildTree();
            ShowNoSelection(ctx);
            _vm.StatusText = $"Blad '{label}' verwijderd";
        }
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
        HideSheetSetFieldsBorder();
    }

    private void ShowSheetSetProperties(DocumentContext ctx)
    {
        NoSelectionLabel.Visibility      = Visibility.Collapsed;
        SubsetPanel.Visibility           = Visibility.Collapsed;
        HoofdveldenPanel.Visibility      = Visibility.Collapsed;
        OverigeVeldenExpander.Visibility = Visibility.Collapsed;
        CustomPropsExpander.Visibility   = Visibility.Collapsed;
        HideSheetSetFieldsBorder();
        ShowSheetSetPanel(ctx);
    }

    private void ShowSubsetProperties(SubsetNode subset)
    {
        NoSelectionLabel.Visibility      = Visibility.Collapsed;
        SubsetPanel.Visibility           = Visibility.Visible;
        HoofdveldenPanel.Visibility      = Visibility.Collapsed;
        OverigeVeldenExpander.Visibility = Visibility.Collapsed;
        CustomPropsExpander.Visibility   = Visibility.Collapsed;
        HideSheetSetPanel();
        HideSheetSetFieldsBorder();
        SubsetNameBox.Text = subset.Name;
        SubsetDescBox.Text = subset.Info.Description ?? string.Empty;
    }

    private void ShowSheetProperties(List<SheetNode> sheets)
    {
        HideSheetSetFieldsBorder();
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
            TbLayout.Text       = info.LayoutName     ?? string.Empty;
            TbDwgFile.Text      = info.RelativeDwgFileName ?? info.DwgFileName ?? string.Empty;
            CbDoNotPlot.IsChecked = info.DoNotPlot;
            var sheetOnlyKeys = GetSheetOnlyKeys();
            CustomPropsPanel.ItemsSource = info.CustomProperties
                .Where(kv => sheetOnlyKeys.Contains(kv.Key))
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
            TbLayout.Text       = string.Empty;
            TbDwgFile.Text      = string.Empty;
            CbDoNotPlot.IsChecked = MixedBool(sheets, i => i.DoNotPlot);

            // Show custom props common to ALL selected sheets; [gedeeld] when values differ
            // Only show per-sheet (Flags=2) properties, not sheet-set-level (Flags=1) properties
            var sheetOnlyKeysMulti = GetSheetOnlyKeys();
            var keysets = sheets
                .Select(s => s.Info.CustomProperties.Keys
                    .Where(k => sheetOnlyKeysMulti.Contains(k))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase))
                .ToList();
            var commonKeys = keysets.Count > 0
                ? keysets.Aggregate((a, b) => { a.IntersectWith(b); return a; })
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var multiRows = commonKeys.OrderBy(k => k).Select(key =>
            {
                var distinct = sheets
                    .Select(s => s.Info.CustomProperties.TryGetValue(key, out var v) ? v : null)
                    .Distinct().ToList();
                return new CustomPropRow(key, distinct.Count == 1 ? distinct[0] ?? string.Empty : "[gedeeld]");
            }).ToList();
            CustomPropsPanel.ItemsSource = multiRows.Count > 0 ? multiRows : null;
        }

        // Inject read-only sheetset Flags=1 fields above the custom props section
        var activeDoc = _vm.ActiveDocument;
        if (activeDoc != null)
        {
            var setDefs = activeDoc.Document.Info.CustomPropertyDefinitions
                .Where(d => d.Flags == 1)
                .ToList();
            if (setDefs.Count > 0)
            {
                _sheetSetFieldsBorder = BuildSheetSetFieldsBorder(
                    setDefs, activeDoc.Document.Info.CustomProperties);
                int idx = PropertyStack.Children.IndexOf(CustomPropsExpander);
                if (idx >= 0)
                    PropertyStack.Children.Insert(idx, _sheetSetFieldsBorder);
            }
        }
    }

    // Dynamic SheetSet panel (injected into ScrollViewer StackPanel)
    private Border? _sheetSetPanelBorder;
    private TextBox? SheetSetDescBox;
    private TextBox? SheetSetProjectNameBox;
    private TextBox? SheetSetProjectNumberBox;
    private TextBox? SheetSetProjectPhaseBox;
    private TextBox? SheetSetProjectMilestoneBox;
    private ItemsControl? _sheetSetPropsCtrl;
    private ItemsControl? _bladVeldenCtrl;
    private Border?       _sheetSetFieldsBorder;

    private void ShowSheetSetPanel(DocumentContext ctx)
    {
        HideSheetSetPanel();
        var info = ctx.Document.Info;

        var grid = new Grid { Margin = new Thickness(8) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

        int rowIndex = 0;
        TextBox AddRow(string labelText, string value)
        {
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            var lbl = new TextBlock
            {
                Text = labelText,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 6, 4),
                Foreground = new SolidColorBrush(Color.FromRgb(0x44, 0x44, 0x44))
            };
            var tb = new TextBox
            {
                Text = value,
                Height = 24,
                Padding = new Thickness(3, 0, 3, 0),
                VerticalContentAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 4)
            };
            Grid.SetRow(lbl, rowIndex); Grid.SetColumn(lbl, 0);
            Grid.SetRow(tb,  rowIndex); Grid.SetColumn(tb,  1);
            grid.Children.Add(lbl);
            grid.Children.Add(tb);
            rowIndex++;
            return tb;
        }

        SheetSetNameBox          = AddRow("Naam:",         info.Name             ?? string.Empty);
        SheetSetDescBox          = AddRow("Omschrijving:",  info.Description      ?? string.Empty);
        SheetSetProjectNameBox   = AddRow("Projectnaam:",   info.ProjectName      ?? string.Empty);
        SheetSetProjectNumberBox = AddRow("Projectnummer:", info.ProjectNumber    ?? string.Empty);
        SheetSetProjectPhaseBox  = AddRow("Projectfase:",   info.ProjectPhase     ?? string.Empty);
        SheetSetProjectMilestoneBox = AddRow("Mijlpaal:",   info.ProjectMilestone ?? string.Empty);

        var sheetSetGroupBox = new GroupBox
        {
            Header = "Sheet Set",
            Margin = new Thickness(12, 0, 12, 6),
            BorderBrush = new SolidColorBrush(Color.FromRgb(0x21, 0x96, 0xF3)),
            BorderThickness = new Thickness(1.5),
            Padding = new Thickness(8),
            Content = grid
        };

        // ── Sheetset velden (Flags=1) ─────────────────────────────────────────
        _sheetSetPropsCtrl = new ItemsControl
        {
            ItemTemplate = BuildTwoColPropRowTemplate(readOnly: false, showDelete: true),
            Margin       = new Thickness(0)
        };
        var setRows = new ObservableCollection<SheetSetCustomPropRow>(
            info.CustomPropertyDefinitions
                .Where(d => d.Flags == 1)
                .Select(d => new SheetSetCustomPropRow(d.Name, d.Value, 1)));
        foreach (var row in setRows.ToList())
        {
            var r = row;
            r.DeleteCommand = new SimpleCommand(() =>
            {
                var confirm = MessageBox.Show(
                    $"Property \"{r.Key}\" verwijderen?",
                    "Weet u het zeker?", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (confirm != MessageBoxResult.Yes) return;

                var def = ctx.Document.Info.CustomPropertyDefinitions
                    .FirstOrDefault(d => d.Name.Equals(r.Key, StringComparison.OrdinalIgnoreCase));
                if (def != null) PushDeletedProp(ctx, def);

                SheetSetWriter.DeleteSheetSetCustomProperty(ctx.Document, r.Key);
                setRows.Remove(r);
                ctx.IsDirty = true;
                _vm.MarkDirty();
            });
        }
        _sheetSetPropsCtrl.ItemsSource = setRows;

        var setStack = new StackPanel();
        setStack.Children.Add(MakeHeaderRow("Property", "Waarde"));
        setStack.Children.Add(_sheetSetPropsCtrl);

        var setExpander = new Expander
        {
            Header          = "Sheetset velden",
            IsExpanded      = true,
            Margin          = new Thickness(0, 0, 0, 4),
            BorderBrush     = new SolidColorBrush(Color.FromRgb(0xBD, 0xBD, 0xBD)),
            BorderThickness = new Thickness(1),
            Padding         = new Thickness(4),
            Content         = setStack
        };

        // ── Blad velden (Flags=2) ─────────────────────────────────────────────
        _bladVeldenCtrl = new ItemsControl
        {
            ItemTemplate = BuildTwoColPropRowTemplate(readOnly: false, showDelete: true),
            Margin       = new Thickness(0)
        };
        var bladRows = new ObservableCollection<SheetSetCustomPropRow>(
            info.CustomPropertyDefinitions
                .Where(d => d.Flags == 2)
                .Select(d => new SheetSetCustomPropRow(d.Name, d.Value, 2)));
        foreach (var row in bladRows.ToList())
        {
            var r = row;
            r.DeleteCommand = new SimpleCommand(() =>
            {
                var confirm = MessageBox.Show(
                    $"Property \"{r.Key}\" verwijderen?\nDit verwijdert ook alle bladwaarden voor dit veld.",
                    "Weet u het zeker?", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (confirm != MessageBoxResult.Yes) return;

                var def = ctx.Document.Info.CustomPropertyDefinitions
                    .FirstOrDefault(d => d.Name.Equals(r.Key, StringComparison.OrdinalIgnoreCase));
                if (def != null) PushDeletedProp(ctx, def);

                SheetSetWriter.DeleteSheetSetCustomProperty(ctx.Document, r.Key);
                bladRows.Remove(r);
                ctx.IsDirty = true;
                _vm.MarkDirty();
            });
        }
        _bladVeldenCtrl.ItemsSource = bladRows;

        var bladStack = new StackPanel();
        bladStack.Children.Add(MakeHeaderRow("Property", "Standaard waarde"));
        bladStack.Children.Add(_bladVeldenCtrl);

        var bladExpander = new Expander
        {
            Header          = "Blad velden (standaard beginwaarden)",
            IsExpanded      = true,
            Margin          = new Thickness(0, 0, 0, 4),
            BorderBrush     = new SolidColorBrush(Color.FromRgb(0xBD, 0xBD, 0xBD)),
            BorderThickness = new Thickness(1),
            Padding         = new Thickness(4),
            Content         = bladStack
        };

        var addBtn = new Button
        {
            Content             = "+ Nieuwe property",
            HorizontalAlignment = HorizontalAlignment.Left,
            Padding             = new Thickness(8, 3, 8, 3),
            Margin              = new Thickness(4, 4, 4, 2),
            FontSize            = 11
        };
        addBtn.Click += AddSheetSetCustomProp_Click;

        var custStack = new StackPanel();
        custStack.Children.Add(setExpander);
        custStack.Children.Add(bladExpander);
        custStack.Children.Add(addBtn);

        var custExpander = new Expander
        {
            Header          = "Custom properties",
            IsExpanded      = true,
            Margin          = new Thickness(12, 0, 12, 10),
            BorderBrush     = new SolidColorBrush(Color.FromRgb(0xBD, 0xBD, 0xBD)),
            BorderThickness = new Thickness(1),
            Padding         = new Thickness(4),
            Content         = custStack
        };

        var outerStack = new StackPanel();
        outerStack.Children.Add(sheetSetGroupBox);
        outerStack.Children.Add(custExpander);

        _sheetSetPanelBorder = new Border { Child = outerStack };
        PropertyStack.Children.Insert(0, _sheetSetPanelBorder);
        NoSelectionLabel.Visibility = Visibility.Collapsed;
    }

    private static DataTemplate BuildTwoColPropRowTemplate(bool readOnly, bool showDelete = false)
    {
        var extraCol   = showDelete ? "<ColumnDefinition Width=\"24\"/>" : "";
        var deleteXaml = showDelete
            ? """
              <Button Grid.Column="2" Content="&#xD7;"
                      Command="{Binding DeleteCommand}"
                      Width="20" Height="20" Padding="0"
                      FontSize="13" VerticalAlignment="Center"
                      Background="Transparent" BorderThickness="0"
                      Foreground="#E53935"/>
              """
            : "";

        // WPF XAML compiler does not support {{ }} escaping in raw string literals — use variables instead.
        var bindKey   = "{Binding Key}";
        var bindValue = "{Binding Value, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}";

        var xaml = $"""
            <DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="160"/>
                        <ColumnDefinition Width="*"/>
                        {extraCol}
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="{bindKey}"
                               VerticalAlignment="Center" Padding="4,3" FontSize="12" Foreground="#444"/>
                    <TextBox   Grid.Column="1"
                               Text="{bindValue}"
                               IsReadOnly="__RO__"
                               Height="22" Padding="3,0" VerticalContentAlignment="Center"
                               Margin="0,0,0,4" BorderThickness="0,0,0,1" BorderBrush="#DDD"
                               Background="__BG__" Foreground="__FG__"/>
                    {deleteXaml}
                </Grid>
            </DataTemplate>
            """
            .Replace("__RO__", readOnly ? "True"    : "False")
            .Replace("__BG__", readOnly ? "#F0F0F0" : "Transparent")
            .Replace("__FG__", readOnly ? "#888"    : "#222");
        return (DataTemplate)System.Windows.Markup.XamlReader.Parse(xaml);
    }

    private static Grid MakeHeaderRow(string col1, string col2)
    {
        var row = new Grid
        {
            Background = new SolidColorBrush(Color.FromRgb(0xF5, 0xF5, 0xF5)),
            Margin     = new Thickness(0, 2, 0, 2)
        };
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
        row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        var h1 = new TextBlock { Text = col1, FontWeight = FontWeights.SemiBold, Padding = new Thickness(4, 3, 4, 3), FontSize = 12 };
        var h2 = new TextBlock { Text = col2, FontWeight = FontWeights.SemiBold, Padding = new Thickness(4, 3, 4, 3), FontSize = 12 };
        Grid.SetColumn(h1, 0); Grid.SetColumn(h2, 1);
        row.Children.Add(h1); row.Children.Add(h2);
        return row;
    }

    private void HideSheetSetPanel()
    {
        if (_sheetSetPanelBorder != null)
        {
            PropertyStack.Children.Remove(_sheetSetPanelBorder);
            _sheetSetPanelBorder        = null;
            SheetSetNameBox             = null;
            SheetSetDescBox             = null;
            SheetSetProjectNameBox      = null;
            SheetSetProjectNumberBox    = null;
            SheetSetProjectPhaseBox     = null;
            SheetSetProjectMilestoneBox = null;
            _sheetSetPropsCtrl          = null;
            _bladVeldenCtrl             = null;
        }
    }

    private void HideSheetSetFieldsBorder()
    {
        if (_sheetSetFieldsBorder != null)
        {
            PropertyStack.Children.Remove(_sheetSetFieldsBorder);
            _sheetSetFieldsBorder = null;
        }
    }

    private Border BuildSheetSetFieldsBorder(
        List<CustomPropertyDefinition> defs, Dictionary<string, string> sheetSetValues)
    {
        var items = new ItemsControl { ItemTemplate = BuildTwoColPropRowTemplate(readOnly: true) };
        items.ItemsSource = defs
            .Select(d =>
            {
                sheetSetValues.TryGetValue(d.Name, out var val);
                return new SheetSetCustomPropRow(d.Name, val ?? string.Empty, 1);
            })
            .ToList();

        var stack = new StackPanel();
        stack.Children.Add(MakeHeaderRow("Property", "Waarde"));
        stack.Children.Add(items);

        var expander = new Expander
        {
            Header          = "Sheetset velden (alleen-lezen)",
            IsExpanded      = true,
            Margin          = new Thickness(12, 0, 12, 6),
            BorderBrush     = new SolidColorBrush(Color.FromRgb(0xBD, 0xBD, 0xBD)),
            BorderThickness = new Thickness(1),
            Padding         = new Thickness(4),
            Content         = stack
        };
        return new Border { Child = expander };
    }

    // Returns the set of property names that have Flags=2 (per-sheet); falls back to all keys
    // when no definitions are present so that existing files without definitions still work.
    private HashSet<string> GetSheetOnlyKeys()
    {
        var defs = _vm.ActiveDocument?.Document.Info.CustomPropertyDefinitions;
        if (defs == null || defs.Count == 0)
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sheetKeys = defs
            .Where(d => d.Flags == 2)
            .Select(d => d.Name)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return sheetKeys.Count > 0 ? sheetKeys
            : defs.Select(d => d.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private void ShowSelectionForDoc(DocumentContext ctx)
    {
        if (ctx.SelectedRoot == ctx.Document)   ShowSheetSetProperties(ctx);
        else if (ctx.SelectedSubset != null)    ShowSubsetProperties(ctx.SelectedSubset);
        else if (ctx.SelectedSheets.Count > 0)  ShowSheetProperties(ctx.SelectedSheets.ToList());
        else                                    ShowNoSelection(ctx);
    }

    private void RefreshPropertyPanel(DocumentContext ctx) => ShowSelectionForDoc(ctx);

    private void AddCustomPropButton_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null || ctx.SelectedSheets.Count == 0) return;

        string? key = PromptInput("Nieuwe custom property", "Naam van de property:");
        if (string.IsNullOrWhiteSpace(key)) return;

        var rows = CustomPropsPanel.ItemsSource?.OfType<CustomPropRow>().ToList() ?? [];
        if (rows.Any(r => string.Equals(r.Key, key, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show($"Property '{key}' staat al in de lijst.",
                "Al aanwezig", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        rows.Add(new CustomPropRow(key, string.Empty));
        CustomPropsPanel.ItemsSource = rows;
    }

    private void AddSheetSetCustomProp_Click(object sender, RoutedEventArgs e)
    {
        var ctx = _vm.ActiveDocument;
        if (ctx == null) return;

        var dlg = new AddCustomPropertyDialog { Owner = this };
        if (dlg.ShowDialog() != true) return;

        var name = dlg.PropertyName;
        var defs = ctx.Document.Info.CustomPropertyDefinitions;
        if (defs.Any(d => string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show($"Property '{name}' bestaat al.", "Al aanwezig",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        SheetSetWriter.AddSheetSetCustomProperty(ctx.Document, name, dlg.Flags);
        ctx.IsDirty = true;
        _vm.MarkDirty();
        ShowSheetSetProperties(ctx);
    }

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

    private static bool IsAutoCADRunning()
        => System.Diagnostics.Process.GetProcessesByName("acad").Length > 0;
}

// ─── CustomPropRow ────────────────────────────────────────────────────────────

public sealed class CustomPropRow(string key, string value)
{
    public string Key   { get; } = key;
    public string Value { get; set; } = value;
}

// ─── SheetSetCustomPropRow ────────────────────────────────────────────────────

public sealed class SheetSetCustomPropRow(string key, string value, int flags)
{
    public string Key        { get; }      = key;
    public string Value      { get; set; } = value;
    public int    Flags      { get; }      = flags;
    // Sheetset-veld (1): één waarde voor de hele set — hier bewerkbaar.
    // Blad-veld (2): waarde verschilt per blad — hier read-only (standaard beginwaarde).
    public bool   IsReadOnly => Flags != 1;
    public string TypeLabel  => Flags == 1 ? "Set" : "Blad (standaard)";
    public ICommand? DeleteCommand { get; set; }
}

// ─── SimpleCommand ────────────────────────────────────────────────────────────

internal sealed class SimpleCommand(Action execute) : ICommand
{
    public event EventHandler? CanExecuteChanged { add { } remove { } }
    public bool CanExecute(object? parameter) => true;
    public void Execute(object? parameter) => execute();
}
