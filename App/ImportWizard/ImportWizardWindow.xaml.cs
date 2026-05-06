using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using Microsoft.Win32;
using SheetSet.Core.Import.Adapters;
using SheetSet.Core.Import.History;
using SheetSet.Core.Import.Mapping;
using SheetSet.Core.Import.Models;
using SheetSet.Core.Import.Parsing;
using SheetSet.Core.Import.Profiles;
using SheetSet.Core.Models;
using SheetSetEditor.Models;

namespace SheetSetEditor.ImportWizard;

public partial class ImportWizardWindow : Window
{
    private readonly DocumentContext _ctx;
    private readonly FlatFileImportSource _source = new();
    private readonly MappingEngine _engine = new();
    private readonly JsonProfileRepository _profileRepo = new();

    private ImportPreview? _preview;
    private List<MappingRowViewModel> _mappings = [];
    private List<string> _targetProperties = [];
    private List<SheetInfo> _preSelectedSheets = [];
    private string _selectedFile = string.Empty;
    private int _page = 1;
    private const int TotalPages = 4;

    public ImportWizardWindow(DocumentContext ctx)
    {
        _ctx = ctx;
        InitializeComponent();

        _targetProperties =
        [
            // Standard per-sheet fields
            "Nummer", "Titel", "Omschrijving", "Revisienummer", "Revisiedatum", "Doel", "Categorie",
            // Standard sheet-set level fields
            "SheetSet.Naam", "SheetSet.Omschrijving", "SheetSet.Projectnaam",
            "SheetSet.Projectnummer", "SheetSet.Fase", "SheetSet.Mijlpaal",
            // Custom properties (Flags=2 per-sheet, Flags=1 set-level)
            .. ctx.Document.Info.CustomPropertyDefinitions.Select(d => d.Name)
        ];
        _preSelectedSheets = ctx.SelectedSheets.Select(n => n.Info).ToList();

        BuildDryRunGridColumns();
        RefreshProfileCombo();
        UpdateSheetSelectionLabel();
        GoToPage(1);
    }

    // ── Navigatie ──────────────────────────────────────────────────────────────

    private void GoToPage(int page)
    {
        _page = page;
        Page1.Visibility = page == 1 ? Visibility.Visible : Visibility.Collapsed;
        Page2.Visibility = page == 2 ? Visibility.Visible : Visibility.Collapsed;
        Page3.Visibility = page == 3 ? Visibility.Visible : Visibility.Collapsed;
        Page4.Visibility = page == 4 ? Visibility.Visible : Visibility.Collapsed;

        StepLabel.Text = page switch
        {
            1 => $"Stap 1 van {TotalPages}: Bestand(en) kiezen",
            2 => $"Stap 2 van {TotalPages}: Koppelen",
            3 => $"Stap 3 van {TotalPages}: Voorbeeld",
            4 => $"Stap 4 van {TotalPages}: Resultaat",
            _ => string.Empty
        };

        BackButton.IsEnabled = page > 1 && page < TotalPages;
        NextButton.IsEnabled = page < TotalPages;
        NextButton.Content   = page switch
        {
            2 => "Voorbeeld →",
            3 => "Importeren ✓",
            _ => "Volgende →"
        };
        CancelButton.Content = page == TotalPages ? "Sluiten" : "Annuleren";
    }

    private void Back_Click(object sender, RoutedEventArgs e)   => GoToPage(_page - 1);
    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();

    private void Next_Click(object sender, RoutedEventArgs e)
    {
        if (_page == 1) { GoToPage(2); return; }
        if (_page == 2) { RunDryRun(); return; }
        if (_page == 3) { RunImport(); }
    }

    // ── Stap 1: bestand kiezen ────────────────────────────────────────────────

    private void BrowseFile_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title  = "Kies importbestand",
            Filter = "Importbestanden (*.csv;*.tsv;*.txt)|*.csv;*.tsv;*.txt|Alle bestanden (*.*)|*.*"
        };
        if (dlg.ShowDialog() != true) return;

        _selectedFile    = dlg.FileName;
        FilePathBox.Text = dlg.FileName;
        LoadPreview(dlg.FileName);
    }

    private void LoadPreview(string filePath)
    {
        try
        {
            _preview = _source.GetPreview(filePath);

            var h         = _preview.DetectedHints;
            var layout    = h.Layout == FieldLayout.Vertical ? "Verticaal" : "Horizontaal (CSV/TSV)";
            var delimiter = h.Delimiter.HasValue ? $"'{h.Delimiter}'" : "geen";
            FormatLabel.Text = $"Layout: {layout}  |  Scheidingsteken: {delimiter}  |  Encoding: {h.Encoding}  |  Kolommen/velden: {_preview.Headers.Count}";

            var errors = _preview.ValidationIssues.Where(i => i.Severity == ValidationSeverity.Error).ToList();
            ValidationLabel.Text       = errors.Count > 0 ? string.Join("\n", errors.Select(i => i.Message)) : string.Empty;
            ValidationLabel.Visibility = errors.Count > 0 ? Visibility.Visible : Visibility.Collapsed;

            FormatPanel.Visibility  = Visibility.Visible;
            PreviewPanel.Visibility = Visibility.Visible;

            BuildPreviewGrid();
            BuildMappingRows();
            TryAutoSelectProfileForFile(filePath);

            NextButton.IsEnabled = _preview.Headers.Count > 0 && !_preview.HasErrors;
        }
        catch (Exception ex)
        {
            FormatPanel.Visibility  = Visibility.Visible;
            PreviewPanel.Visibility = Visibility.Collapsed;
            FormatLabel.Text        = $"Fout bij lezen: {ex.Message}";
            NextButton.IsEnabled    = false;
        }
    }

    private void TryAutoSelectProfileForFile(string filePath)
    {
        var match = _profileRepo.FindByFileName(Path.GetFileName(filePath));
        if (match == null) { ProfileHintLabel.Visibility = Visibility.Collapsed; return; }

        var item = ProfileCombo.Items.Cast<ImportProfile>().FirstOrDefault(p => p.Id == match.Id);
        if (item == null) return;

        ProfileCombo.SelectedItem       = item;
        ProfileHintLabel.Text           = $"Profiel \"{match.Name}\" automatisch herkend op basis van bestandsnaam.";
        ProfileHintLabel.Visibility     = Visibility.Visible;
        ApplyProfile(match);
    }

    private void BuildPreviewGrid()
    {
        PreviewGrid.Columns.Clear();
        if (_preview == null) return;

        if (_preview.DetectedHints.Layout == FieldLayout.Vertical)
        {
            PreviewGrid.Columns.Add(new DataGridTextColumn { Header = "Veld",   Binding = new Binding("Key"),   Width = new DataGridLength(100) });
            PreviewGrid.Columns.Add(new DataGridTextColumn { Header = "Waarde", Binding = new Binding("Value"), Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
            var row = _preview.SampleRows.FirstOrDefault();
            PreviewGrid.ItemsSource = row?.Fields.Select(kv => new KeyValuePair<string, string>(kv.Key, kv.Value)).ToList();
            return;
        }

        foreach (var header in _preview.Headers.Take(10))
            PreviewGrid.Columns.Add(new DataGridTextColumn
            {
                Header  = header,
                Binding = new Binding($"Fields[{header}]"),
                Width   = new DataGridLength(1, DataGridLengthUnitType.Star)
            });
        PreviewGrid.ItemsSource = _preview.SampleRows;
    }

    // ── Stap 2: koppelen ──────────────────────────────────────────────────────

    private void BuildMappingRows()
    {
        if (_preview == null) return;
        var withEmpty = new[] { string.Empty }.Concat(_targetProperties).ToList();

        _mappings = _preview.Headers.Select(h =>
        {
            var sample = _preview.SampleRows
                .Select(r => r.Fields.TryGetValue(h, out var v) ? v : string.Empty)
                .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? string.Empty;

            return new MappingRowViewModel
            {
                SourceField      = h,
                SourceSample     = Truncate(sample, 30),
                AvailableTargets = withEmpty,
                TargetProperty   = AutoMatch(h)
            };
        }).ToList();

        MappingList.ItemsSource = _mappings;
    }

    private void ClearMapping_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: string field })
        {
            var row = _mappings.FirstOrDefault(m => m.SourceField == field && !m.IsDerived);
            if (row != null) row.TargetProperty = string.Empty;
        }
    }

    private void AutoMap_Click(object sender, RoutedEventArgs e)
    {
        var matched = 0;
        foreach (var row in _mappings.Where(m => string.IsNullOrEmpty(m.TargetProperty) && !m.IsDerived))
        {
            var match = AutoMatch(row.SourceField);
            if (!string.IsNullOrEmpty(match)) { row.TargetProperty = match; matched++; }
        }
        AutoMapLabel.Text = matched > 0
            ? $"{matched} koppeling(en) automatisch ingevuld."
            : "Geen nieuwe koppelingen gevonden.";
    }

    // ── Auto-match met fuzzy fallback ─────────────────────────────────────────

    private string AutoMatch(string sourceField)
    {
        if (string.IsNullOrEmpty(sourceField)) return string.Empty;

        // 1. Exact
        var exact = _targetProperties.FirstOrDefault(t =>
            t.Equals(sourceField, StringComparison.OrdinalIgnoreCase));
        if (exact != null) return exact;

        // 2. Normalize (remove spaces, underscores, hyphens) and compare
        var norm = Normalize(sourceField);
        var byNorm = _targetProperties.FirstOrDefault(t =>
            Normalize(t).Equals(norm, StringComparison.OrdinalIgnoreCase));
        if (byNorm != null) return byNorm;

        // 3. Contains (either direction)
        var contains = _targetProperties.FirstOrDefault(t =>
            t.Contains(sourceField, StringComparison.OrdinalIgnoreCase) ||
            sourceField.Contains(t, StringComparison.OrdinalIgnoreCase));
        if (contains != null) return contains;

        // 4. Fuzzy (Levenshtein similarity ≥ 0.75)
        var best = _targetProperties
            .Select(t => (target: t, score: LevenshteinSimilarity(norm, Normalize(t))))
            .Where(x => x.score >= 0.75)
            .OrderByDescending(x => x.score)
            .FirstOrDefault();
        return best.target ?? string.Empty;
    }

    private static string Normalize(string s)
        => new string(s.ToLowerInvariant().Where(c => char.IsLetterOrDigit(c)).ToArray());

    private static double LevenshteinSimilarity(string a, string b)
    {
        if (a == b) return 1.0;
        if (a.Length == 0 || b.Length == 0) return 0.0;
        var d = LevenshteinDistance(a, b);
        return 1.0 - (double)d / Math.Max(a.Length, b.Length);
    }

    private static int LevenshteinDistance(string a, string b)
    {
        var m = a.Length; var n = b.Length;
        var dp = new int[m + 1, n + 1];
        for (var i = 0; i <= m; i++) dp[i, 0] = i;
        for (var j = 0; j <= n; j++) dp[0, j] = j;
        for (var i = 1; i <= m; i++)
            for (var j = 1; j <= n; j++)
                dp[i, j] = a[i - 1] == b[j - 1]
                    ? dp[i - 1, j - 1]
                    : 1 + Math.Min(dp[i - 1, j - 1], Math.Min(dp[i - 1, j], dp[i, j - 1]));
        return dp[m, n];
    }

    // ── Splits-transformatie ──────────────────────────────────────────────────

    private void AddSplit_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: string field }) return;
        var sourceRow = _mappings.FirstOrDefault(m => m.SourceField == field && !m.IsDerived);
        if (sourceRow == null) return;

        var results = ShowMultiSplitDialog(field, sourceRow.SourceSample);
        if (results == null || results.Count == 0) return;

        // Replace any existing derived rows for this field
        _mappings.RemoveAll(m => m.IsDerived && m.SourceField == field);

        var insertIdx = _mappings.IndexOf(sourceRow);
        foreach (var (sep, idx, target) in results)
        {
            var partLabel = idx >= 0 ? $"deel {idx + 1}" : "laatste deel";
            _mappings.Insert(++insertIdx, new MappingRowViewModel
            {
                SourceField      = field,
                SourceSample     = PreviewSplit(sourceRow.SourceSample, sep, idx),
                AvailableTargets = sourceRow.AvailableTargets,
                TargetProperty   = target,
                IsDerived        = true,
                DerivedLabel     = $"↳ {partLabel}",
                Transformations  = [new TransformationRule
                {
                    Type       = TransformationType.SplitPart,
                    Parameters = new Dictionary<string, string> { ["separator"] = sep, ["index"] = idx.ToString() }
                }]
            });
        }

        MappingList.ItemsSource = null;
        MappingList.ItemsSource = _mappings;
    }

    // ── Validatie-instellingen ────────────────────────────────────────────────

    private void OpenValidation_Click(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: string field }) return;
        var row = _mappings.FirstOrDefault(m => m.SourceField == field && !m.IsDerived);
        if (row == null) return;

        var dlg = new Window
        {
            Title         = $"Validatie — veld '{field}'",
            SizeToContent = SizeToContent.Height,
            Width         = 380,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode    = ResizeMode.NoResize,
            ShowInTaskbar = false,
            Owner         = this
        };

        var grid = new Grid { Margin = new Thickness(16) };
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        for (var r = 0; r < 4; r++)
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        void Lbl(string text, int row)
        {
            var tb = new TextBlock
            {
                Text = text, VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 6),
                Foreground = new SolidColorBrush(Color.FromRgb(0x44, 0x44, 0x44))
            };
            Grid.SetRow(tb, row); Grid.SetColumn(tb, 0);
            grid.Children.Add(tb);
        }

        Lbl("Verplicht:", 0);
        var reqCheck = new CheckBox { IsChecked = row.IsRequired, Margin = new Thickness(0, 0, 0, 6), VerticalAlignment = VerticalAlignment.Center };
        Grid.SetRow(reqCheck, 0); Grid.SetColumn(reqCheck, 1);
        grid.Children.Add(reqCheck);

        Lbl("Max. lengte:", 1);
        var maxBox = new TextBox
        {
            Text = row.ValidationMaxLength?.ToString() ?? string.Empty,
            Height = 26, Padding = new Thickness(4, 2, 4, 2), Margin = new Thickness(0, 0, 0, 6)
        };
        Grid.SetRow(maxBox, 1); Grid.SetColumn(maxBox, 1);
        grid.Children.Add(maxBox);

        Lbl("Regex patroon:", 2);
        var patBox = new TextBox
        {
            Text = row.ValidationPattern, Height = 26,
            Padding = new Thickness(4, 2, 4, 2), Margin = new Thickness(0, 0, 0, 10)
        };
        Grid.SetRow(patBox, 2); Grid.SetColumn(patBox, 1);
        grid.Children.Add(patBox);

        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var ok     = new Button { Content = "OK",        Width = 80, Height = 26, Margin = new Thickness(0, 0, 6, 0), IsDefault = true };
        var cancel = new Button { Content = "Annuleren", Width = 80, Height = 26, IsCancel = true };
        ok.Click     += (_, _) => dlg.DialogResult = true;
        cancel.Click += (_, _) => dlg.DialogResult = false;
        btnRow.Children.Add(ok); btnRow.Children.Add(cancel);
        Grid.SetRow(btnRow, 3); Grid.SetColumnSpan(btnRow, 2);
        grid.Children.Add(btnRow);

        dlg.Content = grid;
        if (dlg.ShowDialog() != true) return;

        row.IsRequired        = reqCheck.IsChecked == true;
        row.ValidationPattern = patBox.Text.Trim();
        row.ValidationMaxLength = int.TryParse(maxBox.Text.Trim(), out var ml) ? ml : null;
    }

    // ── Sheet selectie ────────────────────────────────────────────────────────

    private void UpdateSheetSelectionLabel()
    {
        var count = _preSelectedSheets.Count;
        SelectedSheetsLabel.Text      = $"Geselecteerde sheets ({count})";
        SelectedSheetsRadio.IsEnabled = count > 0;
        if (count == 0) AllSheetsRadio.IsChecked = true;
    }

    private void PickSheets_Click(object sender, RoutedEventArgs e)
    {
        var picked = ShowSheetPickerDialog(_ctx.Document, _preSelectedSheets);
        if (picked == null) return;
        _preSelectedSheets = picked;
        UpdateSheetSelectionLabel();
        if (picked.Count > 0) SelectedSheetsRadio.IsChecked = true;
    }

    private List<SheetInfo>? ShowSheetPickerDialog(SheetSetDocument doc, List<SheetInfo> current)
    {
        var dlg = new Window
        {
            Title         = "Selecteer sheets",
            Width         = 440, MaxHeight = 620,
            SizeToContent = SizeToContent.Height,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode    = ResizeMode.CanResizeWithGrip,
            ShowInTaskbar = false,
            Owner         = this
        };

        var outer = new Grid { Margin = new Thickness(16) };
        outer.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        outer.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star), MaxHeight = 380 });
        outer.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        outer.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var hdr = new TextBlock { Text = "Vink de sheets aan waarop de import wordt toegepast:", TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 8) };
        Grid.SetRow(hdr, 0); outer.Children.Add(hdr);

        // Build pick-node tree
        var roots = BuildPickTree(doc, current);

        var tree = new TreeView { BorderBrush = new SolidColorBrush(Color.FromRgb(0xDD, 0xDD, 0xDD)), BorderThickness = new Thickness(1) };
        foreach (var node in roots)
            tree.Items.Add(BuildTreeItem(node));

        // Set initial checked state after all Controls are assigned
        SetInitialChecked(roots, current);
        Grid.SetRow(tree, 1); outer.Children.Add(tree);

        // Select-all / deselect-all
        var selRow   = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 8, 0, 10) };
        var selAll   = new Button { Content = "Alles selecteren",   Height = 24, Padding = new Thickness(8, 0, 8, 0), Margin = new Thickness(0, 0, 6, 0), FontSize = 11 };
        var deselAll = new Button { Content = "Alles deselecteren", Height = 24, Padding = new Thickness(8, 0, 8, 0), FontSize = 11 };
        selAll  .Click += (_, _) => { SetAllNodes(roots, true);  };
        deselAll.Click += (_, _) => { SetAllNodes(roots, false); };
        selRow.Children.Add(selAll); selRow.Children.Add(deselAll);
        Grid.SetRow(selRow, 2); outer.Children.Add(selRow);

        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
        var ok     = new Button { Content = "OK",        Width = 80, Height = 26, Margin = new Thickness(0, 0, 6, 0), IsDefault = true };
        var cancel = new Button { Content = "Annuleren", Width = 80, Height = 26, IsCancel = true };
        ok.Click     += (_, _) => dlg.DialogResult = true;
        cancel.Click += (_, _) => dlg.DialogResult = false;
        btnRow.Children.Add(ok); btnRow.Children.Add(cancel);
        Grid.SetRow(btnRow, 3); outer.Children.Add(btnRow);

        dlg.Content = outer;
        if (dlg.ShowDialog() != true) return null;

        return CollectSelected(roots);
    }

    // ── PickNode tree ─────────────────────────────────────────────────────────

    private sealed class PickNode
    {
        public string        Label    { get; set; } = string.Empty;
        public SheetInfo?    Sheet    { get; set; }
        public List<PickNode> Children { get; }     = [];
        public PickNode?     Parent   { get; set; }
        public CheckBox?     Control  { get; set; }
        public bool          Updating { get; set; }
        public bool IsGroup => Sheet == null;
    }

    private static List<PickNode> BuildPickTree(SheetSetDocument doc, List<SheetInfo> current)
    {
        var roots = new List<PickNode>();
        foreach (var s in doc.RootSheets)
            roots.Add(new PickNode { Label = FormatSheet(s.Info), Sheet = s.Info });
        foreach (var sub in doc.RootSubsets)
            roots.Add(BuildSubsetNode(sub, null));
        return roots;
    }

    private static PickNode BuildSubsetNode(SubsetNode sub, PickNode? parent)
    {
        var node = new PickNode { Label = sub.Name, Sheet = null, Parent = parent };
        foreach (var s in sub.Sheets)
            node.Children.Add(new PickNode { Label = FormatSheet(s.Info), Sheet = s.Info, Parent = node });
        foreach (var child in sub.Subsets)
        {
            var cn = BuildSubsetNode(child, node);
            node.Children.Add(cn);
        }
        return node;
    }


    private static TreeViewItem BuildTreeItem(PickNode node)
    {
        var check = new CheckBox
        {
            Content       = node.Label,
            IsChecked     = false,
            IsThreeState  = node.IsGroup,
            FontWeight    = node.IsGroup ? FontWeights.SemiBold : FontWeights.Normal,
            VerticalAlignment = VerticalAlignment.Center,
            Padding       = new Thickness(4, 2, 4, 2)
        };
        node.Control = check;

        var tvi = new TreeViewItem { Header = check, IsExpanded = true };
        foreach (var child in node.Children)
            tvi.Items.Add(BuildTreeItem(child));

        // Wire propagation after all children have Controls assigned
        check.Checked += (_, _) =>
        {
            if (node.Updating) return;
            if (node.IsGroup) PropagateDown(node, true);
            RefreshParent(node.Parent);
        };
        check.Unchecked += (_, _) =>
        {
            if (node.Updating) return;
            if (node.IsGroup) PropagateDown(node, false);
            RefreshParent(node.Parent);
        };

        return tvi;
    }

    private static void SetAllNodes(IEnumerable<PickNode> nodes, bool value)
    {
        foreach (var node in nodes)
        {
            if (node.Control != null)
            {
                node.Updating = true;
                node.Control.IsChecked = value;
                node.Updating = false;
            }
            SetAllNodes(node.Children, value);
        }
    }

    private static void PropagateDown(PickNode node, bool value)
    {
        foreach (var child in node.Children)
        {
            if (child.Control != null)
            {
                child.Updating = true;
                child.Control.IsChecked = value;
                child.Updating = false;
            }
            if (child.IsGroup) PropagateDown(child, value);
        }
    }

    private static void RefreshParent(PickNode? parent)
    {
        if (parent?.Control == null || parent.Updating) return;
        var all  = parent.Children.All(c => c.Control?.IsChecked == true);
        var none = parent.Children.All(c => c.Control?.IsChecked == false);
        parent.Updating = true;
        parent.Control.IsChecked = all ? true : none ? false : null;
        parent.Updating = false;
        RefreshParent(parent.Parent);
    }

    private static void SetInitialChecked(IEnumerable<PickNode> nodes, List<SheetInfo> current)
    {
        foreach (var node in nodes)
        {
            if (!node.IsGroup && node.Control != null && node.Sheet != null)
                node.Control.IsChecked = current.Contains(node.Sheet);
            SetInitialChecked(node.Children, current);
        }
        // Refresh group states bottom-up handled by RefreshParent; do it for each root group
        foreach (var node in nodes.Where(n => n.IsGroup))
            RefreshParent(node);
    }

    private static List<SheetInfo> CollectSelected(IEnumerable<PickNode> nodes)
    {
        var result = new List<SheetInfo>();
        foreach (var node in nodes)
        {
            if (!node.IsGroup && node.Control?.IsChecked == true && node.Sheet != null)
                result.Add(node.Sheet);
            result.AddRange(CollectSelected(node.Children));
        }
        return result;
    }

    private static string FormatSheet(SheetInfo info)
    {
        var parts = new[] { info.Number, info.Title }.Where(s => !string.IsNullOrEmpty(s));
        return string.Join("  –  ", parts);
    }

    private List<SheetInfo> GetTargetSheets()
    {
        if (SelectedSheetsRadio.IsChecked == true && _preSelectedSheets.Count > 0)
            return _preSelectedSheets;
        return _ctx.Document.GetAllSheets().Select(s => s.Info).ToList();
    }

    // ── Profielen ─────────────────────────────────────────────────────────────

    private void RefreshProfileCombo()
    {
        ProfileCombo.ItemsSource  = null;
        ProfileCombo.ItemsSource  = _profileRepo.List();
        ProfileCombo.SelectedItem = null;
    }

    private void ProfileCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        => ProfileHintLabel.Visibility = Visibility.Collapsed;

    private void LoadProfile_Click(object sender, RoutedEventArgs e)
    {
        if (ProfileCombo.SelectedItem is not ImportProfile profile) return;
        ApplyProfile(profile);
        AutoMapLabel.Text = $"Profiel '{profile.Name}' geladen.";
    }

    private void ApplyProfile(ImportProfile profile)
    {
        // Remove all previously-derived (split) rows so we can reconstruct them cleanly
        _mappings.RemoveAll(m => m.IsDerived);

        foreach (var row in _mappings.ToList())
        {
            // Empty source line (e.g. blank line in vertical file) — never map
            if (string.IsNullOrWhiteSpace(row.SourceSample))
            {
                row.TargetProperty = string.Empty;
                continue;
            }

            // Plain mapping (no SplitPart transform) → update target + validation
            var plain = profile.FieldMappings.FirstOrDefault(m =>
                m.SourceColumn.Equals(row.SourceField, StringComparison.OrdinalIgnoreCase) &&
                !m.Transformations.Any(t => t.Type == TransformationType.SplitPart));

            if (plain != null)
            {
                row.TargetProperty      = plain.TargetProperty;
                row.IsRequired          = plain.IsRequired;
                row.ValidationPattern   = plain.ValidationPattern ?? string.Empty;
                row.ValidationMaxLength = plain.MaxLength;
            }
            else
            {
                // Field is completely unknown to this profile — leave AutoMatch as-is.
                // But if it exists ONLY as split parts, the parent must be cleared.
                var isInProfileAsPlain = profile.FieldMappings.Any(m =>
                    m.SourceColumn.Equals(row.SourceField, StringComparison.OrdinalIgnoreCase) &&
                    !m.Transformations.Any(t => t.Type == TransformationType.SplitPart));
                var isInProfileAsSplit = profile.FieldMappings.Any(m =>
                    m.SourceColumn.Equals(row.SourceField, StringComparison.OrdinalIgnoreCase) &&
                    m.Transformations.Any(t => t.Type == TransformationType.SplitPart));
                if (!isInProfileAsPlain && isInProfileAsSplit)
                    row.TargetProperty = string.Empty;
            }

            // Split mappings → reconstruct derived rows after the parent
            var splits = profile.FieldMappings
                .Where(m => m.SourceColumn.Equals(row.SourceField, StringComparison.OrdinalIgnoreCase) &&
                            m.Transformations.Any(t => t.Type == TransformationType.SplitPart))
                .ToList();

            var insertIdx = _mappings.IndexOf(row);
            foreach (var sm in splits)
            {
                var rule = sm.Transformations.First(t => t.Type == TransformationType.SplitPart);
                int.TryParse(rule.Parameters.GetValueOrDefault("index", "0"), out var idx);
                var sep       = rule.Parameters.GetValueOrDefault("separator", " / ");
                var partLabel = idx >= 0 ? $"deel {idx + 1}" : "laatste deel";

                _mappings.Insert(++insertIdx, new MappingRowViewModel
                {
                    SourceField      = row.SourceField,
                    SourceSample     = PreviewSplit(row.SourceSample, sep, idx),
                    AvailableTargets = row.AvailableTargets,
                    TargetProperty   = sm.TargetProperty,
                    IsDerived        = true,
                    DerivedLabel     = $"↳ {partLabel}",
                    Transformations  = sm.Transformations
                });
            }
        }

        MappingList.ItemsSource = null;
        MappingList.ItemsSource = _mappings;
    }

    private void SaveProfile_Click(object sender, RoutedEventArgs e)
    {
        var name = ShowInputDialog("Profielnaam:", "Sla op als profiel");
        if (string.IsNullOrWhiteSpace(name)) return;

        var existing = _profileRepo.FindByName(name);
        if (existing != null)
        {
            var confirm = MessageBox.Show($"Profiel \"{name}\" bestaat al. Overschrijven?",
                "Profiel overschrijven", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;

            var updated = BuildProfile(name);
            updated.Id                  = existing.Id;
            updated.CreatedAt           = existing.CreatedAt;
            updated.AssociatedFileNames = existing.AssociatedFileNames;
            AssociateCurrentFile(updated);
            _profileRepo.Save(updated);
        }
        else
        {
            var profile = BuildProfile(name);
            AssociateCurrentFile(profile);
            _profileRepo.Save(profile);
        }

        RefreshProfileCombo();
        AutoMapLabel.Text = $"Profiel '{name}' opgeslagen.";
    }

    private void DeleteProfile_Click(object sender, RoutedEventArgs e)
    {
        if (ProfileCombo.SelectedItem is not ImportProfile profile) return;
        var confirm = MessageBox.Show($"Profiel \"{profile.Name}\" verwijderen?",
            "Verwijder profiel", MessageBoxButton.YesNo, MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes) return;
        _profileRepo.Delete(profile.Id);
        RefreshProfileCombo();
        AutoMapLabel.Text           = $"Profiel '{profile.Name}' verwijderd.";
        ProfileHintLabel.Visibility = Visibility.Collapsed;
    }

    private void AssociateCurrentFile(ImportProfile profile)
    {
        if (_preview == null) return;
        var fileName = Path.GetFileName(_preview.FilePath);
        if (!string.IsNullOrEmpty(fileName) &&
            !profile.AssociatedFileNames.Any(f => string.Equals(f, fileName, StringComparison.OrdinalIgnoreCase)))
            profile.AssociatedFileNames.Add(fileName);
    }

    // ── Stap 3: dry run ───────────────────────────────────────────────────────

    private void RunDryRun()
    {
        if (_preview == null || string.IsNullOrEmpty(_selectedFile)) return;
        try
        {
            var profile    = BuildProfile(string.Empty);
            var rows       = _source.ReadRows(_selectedFile, profile.SourceHints);
            var mapped     = _engine.Apply(rows, profile);
            var sheets     = GetTargetSheets();
            var adapter    = new SheetSetTargetAdapter(sheets, _ctx.Document);
            var changes    = adapter.PreviewChanges(mapped);

            var willChange = changes.Count(c => c.WillChange);
            var unchanged  = changes.Count(c => !c.WillChange);
            var fileDesc   = Path.GetFileName(_selectedFile);

            DryRunSummary.Text = willChange > 0
                ? $"{willChange} waarde(n) zullen worden bijgewerkt in {sheets.Count} sheet(s) vanuit {fileDesc}." +
                  (unchanged > 0 ? $"  {unchanged} zijn al gelijk." : string.Empty)
                : $"Geen wijzigingen gevonden — alle waarden zijn al up-to-date ({fileDesc}).";

            var style   = new Style(typeof(DataGridRow));
            var trigger = new DataTrigger { Binding = new Binding("WillChange"), Value = true };
            trigger.Setters.Add(new Setter(DataGridRow.BackgroundProperty,
                new SolidColorBrush(Color.FromRgb(0xE8, 0xF5, 0xE9))));
            style.Triggers.Add(trigger);
            DryRunGrid.RowStyle = style;

            DryRunGrid.ItemsSource = changes;
            GoToPage(3);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fout bij voorbeeld:\n{ex.Message}", "Fout",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BuildDryRunGridColumns()
    {
        DryRunGrid.Columns.Add(new DataGridTextColumn { Header = "Sheet",          Binding = new Binding("SheetLabel"), Width = new DataGridLength(180) });
        DryRunGrid.Columns.Add(new DataGridTextColumn { Header = "Property",       Binding = new Binding("Property"),   Width = new DataGridLength(130) });
        DryRunGrid.Columns.Add(new DataGridTextColumn { Header = "Huidige waarde", Binding = new Binding("OldValue"),   Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
        DryRunGrid.Columns.Add(new DataGridTextColumn { Header = "Nieuwe waarde",  Binding = new Binding("NewValue"),   Width = new DataGridLength(1, DataGridLengthUnitType.Star) });
    }

    // ── Stap 4: importeren ────────────────────────────────────────────────────

    private void RunImport()
    {
        if (string.IsNullOrEmpty(_selectedFile)) return;
        try
        {
            var profile = BuildProfile(string.Empty);
            var sheets  = GetTargetSheets();
            var snapshot = CreateSnapshot(sheets, Path.GetFileName(_selectedFile));

            var rows    = _source.ReadRows(_selectedFile, profile.SourceHints);
            var mapped  = _engine.Apply(rows, profile);
            var adapter = new SheetSetTargetAdapter(sheets, _ctx.Document);
            var run     = adapter.Apply(mapped, new ImportContext());

            var totalSucceeded = run.SucceededRows;
            var totalFailed    = run.FailedRows;
            var allLines       = new List<string>();

            foreach (var row in run.Results)
            {
                foreach (var (prop, val) in row.TargetValues)
                    allLines.Add($"✓  {prop} = {val}");
                foreach (var w in row.Warnings) allLines.Add($"⚠  {w}");
                foreach (var err in row.Errors)  allLines.Add($"✗  {err}");
            }

            _ctx.LastImportSnapshot = snapshot;
            _ctx.IsDirty            = true;

            GoToPage(4);
            ResultSummary.Text = $"Import klaar — {totalSucceeded} rij(en) geslaagd" +
                                 (totalFailed > 0 ? $", {totalFailed} mislukt" : "") + ".";
            ResultList.ItemsSource = allLines.Count > 0 ? allLines : new[] { "(geen details)" };
            BackButton.IsEnabled   = false;
            NextButton.IsEnabled   = false;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fout tijdens import:\n{ex.Message}", "Fout",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private ImportProfile BuildProfile(string name)
    {
        var hints = _preview?.DetectedHints ?? new SourceHints();
        return new ImportProfile
        {
            Name          = name,
            SourceHints   = hints,
            FieldMappings = _mappings
                .Where(m => !string.IsNullOrEmpty(m.TargetProperty))
                .Select(m => new FieldMapping
                {
                    SourceColumn      = m.SourceField,
                    TargetProperty    = m.TargetProperty,
                    Transformations   = m.Transformations,
                    IsRequired        = m.IsRequired,
                    ValidationPattern = string.IsNullOrEmpty(m.ValidationPattern) ? null : m.ValidationPattern,
                    MaxLength         = m.ValidationMaxLength
                })
                .ToList()
        };
    }

    // ── Snapshot (undo) ───────────────────────────────────────────────────────

    private static ImportSnapshot CreateSnapshot(List<SheetInfo> sheets, string description)
        => new()
        {
            SourceDescription = description,
            Sheets = sheets
                .Where(s => s.Element != null)
                .Select(s => new SheetStateSnapshot
                {
                    Sheet                    = s,
                    OriginalElement          = new System.Xml.Linq.XElement(s.Element!),
                    OriginalCustomProperties = new Dictionary<string, string>(s.CustomProperties)
                })
                .ToList()
        };

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string Truncate(string s, int max)
        => s.Length <= max ? s : s[..max] + "…";

    private static string PreviewSplit(string sample, string separator, int index)
    {
        if (string.IsNullOrEmpty(sample)) return string.Empty;
        var parts = sample.Split(new[] { separator }, StringSplitOptions.None);
        var i = index < 0 ? parts.Length + index : index;
        return i >= 0 && i < parts.Length ? parts[i].Trim() : string.Empty;
    }

    private List<(string sep, int index, string target)>? ShowMultiSplitDialog(string field, string sample)
    {
        var dlg = new Window
        {
            Title         = $"Splits veld '{field}'",
            SizeToContent = SizeToContent.Height,
            Width         = 540, MaxHeight = 600,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode    = ResizeMode.CanResizeWithGrip, ShowInTaskbar = false, Owner = this
        };

        var withEmpty  = new[] { string.Empty }.Concat(_targetProperties).ToList();
        var partCombos = new List<(int index, ComboBox combo)>();

        var outer = new StackPanel { Margin = new Thickness(16) };

        // Separator row
        var sepPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 12) };
        sepPanel.Children.Add(new TextBlock
        {
            Text = "Scheidingsteken:", Width = 140, VerticalAlignment = VerticalAlignment.Center,
            Foreground = new SolidColorBrush(Color.FromRgb(0x44, 0x44, 0x44))
        });
        var sepBox = new TextBox { Text = " / ", Width = 160, Height = 26, Padding = new Thickness(4, 2, 4, 2) };
        sepPanel.Children.Add(sepBox);
        outer.Children.Add(sepPanel);

        // Column headers
        var hdr = new Grid { Margin = new Thickness(0, 0, 0, 4) };
        hdr.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
        hdr.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
        hdr.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        foreach (var (text, col) in new[] { ("Deel", 0), ("Voorbeeld", 1), ("Doel-property", 2) })
        {
            var tb = new TextBlock { Text = text, FontWeight = FontWeights.SemiBold, FontSize = 11, Foreground = new SolidColorBrush(Color.FromRgb(0x88, 0x88, 0x88)) };
            Grid.SetColumn(tb, col); hdr.Children.Add(tb);
        }
        outer.Children.Add(hdr);

        // Parts container (rebuilt on separator change)
        var partsPanel = new StackPanel();
        outer.Children.Add(partsPanel);

        void RebuildParts()
        {
            partsPanel.Children.Clear();
            partCombos.Clear();
            var sep   = sepBox.Text;
            var parts = (!string.IsNullOrEmpty(sample) && !string.IsNullOrEmpty(sep))
                ? sample.Split(new[] { sep }, StringSplitOptions.None)
                : (string.IsNullOrEmpty(sample) ? Array.Empty<string>() : new[] { sample });

            for (var i = 0; i < parts.Length; i++)
            {
                var idx   = i;
                var row   = new Grid { Margin = new Thickness(0, 2, 0, 2) };
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(160) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var label   = new TextBlock { Text = $"↳ deel {idx + 1}", VerticalAlignment = VerticalAlignment.Center, FontStyle = FontStyles.Italic, Foreground = new SolidColorBrush(Color.FromRgb(0x19, 0x76, 0xD2)), FontSize = 12 };
                var preview = new TextBlock { Text = parts[idx].Trim(), VerticalAlignment = VerticalAlignment.Center, Foreground = new SolidColorBrush(Color.FromRgb(0x66, 0x66, 0x66)), FontSize = 12, TextTrimming = TextTrimming.CharacterEllipsis };
                var combo   = new ComboBox  { ItemsSource = withEmpty, SelectedItem = string.Empty, Height = 24, FontSize = 12, IsEditable = true };

                Grid.SetColumn(label,   0); row.Children.Add(label);
                Grid.SetColumn(preview, 1); row.Children.Add(preview);
                Grid.SetColumn(combo,   2); row.Children.Add(combo);
                partsPanel.Children.Add(row);
                partCombos.Add((idx, combo));
            }
        }

        sepBox.TextChanged += (_, _) => RebuildParts();
        RebuildParts();

        // Buttons
        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 14, 0, 0) };
        var ok     = new Button { Content = "OK",        Width = 80, Height = 26, Margin = new Thickness(0, 0, 6, 0), IsDefault = true };
        var cancel = new Button { Content = "Annuleren", Width = 80, Height = 26, IsCancel = true };
        ok.Click     += (_, _) => dlg.DialogResult = true;
        cancel.Click += (_, _) => dlg.DialogResult = false;
        btnRow.Children.Add(ok); btnRow.Children.Add(cancel);
        outer.Children.Add(btnRow);

        dlg.Content = outer;
        if (dlg.ShowDialog() != true) return null;

        var finalSep = sepBox.Text;
        return partCombos
            .Select(pc => (finalSep, pc.index, (pc.combo.SelectedItem as string ?? pc.combo.Text).Trim()))
            .Where(t => !string.IsNullOrEmpty(t.Item3))
            .ToList();
    }

    private static string? ShowInputDialog(string prompt, string title)
    {
        var dlg = new Window
        {
            Title = title, SizeToContent = SizeToContent.Height, Width = 380,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false
        };
        var stack  = new StackPanel { Margin = new Thickness(16) };
        var label  = new TextBlock { Text = prompt, Margin = new Thickness(0, 0, 0, 6) };
        var box    = new TextBox   { Height = 26, Padding = new Thickness(4, 2, 4, 2) };
        var btnRow = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
        var ok     = new Button { Content = "OK",        Width = 80, Height = 26, Margin = new Thickness(0, 0, 6, 0), IsDefault = true };
        var cancel = new Button { Content = "Annuleren", Width = 80, Height = 26, IsCancel = true };
        ok.Click     += (_, _) => dlg.DialogResult = true;
        cancel.Click += (_, _) => dlg.DialogResult = false;
        btnRow.Children.Add(ok); btnRow.Children.Add(cancel);
        stack.Children.Add(label); stack.Children.Add(box); stack.Children.Add(btnRow);
        dlg.Content = stack;
        dlg.Owner   = Application.Current.MainWindow;
        box.Focus();
        return dlg.ShowDialog() == true ? box.Text.Trim() : null;
    }
}
