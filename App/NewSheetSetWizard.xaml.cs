using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using SheetSet.Core.Interop;
using SheetSet.Core.Models;
using SheetSet.Core.Parsing;
using SheetSet.Core.Writing;
using SheetSetEditor.Services;

namespace SheetSetEditor;

public partial class NewSheetSetWizard : Window
{
    private int _currentPage = 1;
    private ObservableCollection<CustomPropertyDefinition> _previewProps = [];

    public string OutputPath { get; private set; } = string.Empty;

    public NewSheetSetWizard()
    {
        InitializeComponent();
        FolderBox.Text   = EditorSettingsService.GetLastSaveFolder()   ?? string.Empty;
        TemplateBox.Text = EditorSettingsService.GetLastTemplatePath() ?? string.Empty;
    }

    private void BrowseFolder_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFolderDialog { Title = "Kies de opslagmap voor de nieuwe sheet set" };
        if (!string.IsNullOrWhiteSpace(FolderBox.Text) && Directory.Exists(FolderBox.Text))
            dlg.InitialDirectory = FolderBox.Text;
        if (dlg.ShowDialog() == true)
            FolderBox.Text = dlg.FolderName;
    }

    private void BrowseTemplate_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title  = "Kies het template bestand",
            Filter = "Sheet Set bestanden (*.dst)|*.dst",
        };
        var initial = TemplateBox.Text.Trim();
        if (File.Exists(initial))
            dlg.InitialDirectory = Path.GetDirectoryName(initial);
        else if (!string.IsNullOrWhiteSpace(FolderBox.Text) && Directory.Exists(FolderBox.Text))
            dlg.InitialDirectory = FolderBox.Text;

        if (dlg.ShowDialog() == true)
            TemplateBox.Text = dlg.FileName;
    }

    private void NextButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentPage == 1)
            GoToPage2();
        else
            CreateSheetSet();
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
        => GoToPage1();

    private void GoToPage2()
    {
        var folder   = FolderBox.Text.Trim();
        var template = TemplateBox.Text.Trim();

        if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
        {
            MessageBox.Show("Kies een bestaande opslagmap.", "Opslagmap vereist",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            FolderBox.Focus();
            return;
        }
        if (string.IsNullOrWhiteSpace(template) || !File.Exists(template))
        {
            MessageBox.Show("Kies een bestaand template bestand (.dst).", "Template vereist",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            TemplateBox.Focus();
            return;
        }

        EditorSettingsService.SetLastSaveFolder(folder);
        EditorSettingsService.SetLastTemplatePath(template);

        // Laad template eigenschappen voor de preview
        try { LoadTemplatePreview(template); }
        catch (Exception ex)
        {
            MessageBox.Show($"Kan template niet lezen:\n{ex.Message}", "Fout",
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        Page1Grid.Visibility = Visibility.Collapsed;
        Page2Grid.Visibility = Visibility.Visible;
        BackButton.Visibility = Visibility.Visible;
        NextButton.Content    = "Aanmaken";
        StepLabel.Text        = "Stap 2 van 2: Naam && Details";
        _currentPage = 2;
        NameBox.Focus();
    }

    private void GoToPage1()
    {
        Page2Grid.Visibility  = Visibility.Collapsed;
        Page1Grid.Visibility  = Visibility.Visible;
        BackButton.Visibility = Visibility.Collapsed;
        NextButton.Content    = "Volgende >";
        StepLabel.Text        = "Stap 1 van 2: Opslaan & Template";
        _currentPage = 1;
    }

    private void LoadTemplatePreview(string templatePath)
    {
        var tempXml = DstCodec.DecodeDstToTempXml(templatePath);
        try
        {
            var parser = new SheetSetParser();
            var doc    = parser.Parse(tempXml);
            _previewProps = new ObservableCollection<CustomPropertyDefinition>(doc.Info.CustomPropertyDefinitions);
            PropsPreview.ItemsSource = _previewProps;
        }
        finally
        {
            if (File.Exists(tempXml)) File.Delete(tempXml);
        }
    }

    private void DeleteProp_Click(object sender, RoutedEventArgs e)
    {
        if (((Button)sender).Tag is CustomPropertyDefinition prop)
            _previewProps.Remove(prop);
    }

    private void CreateSheetSet()
    {
        var name = NameBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show("Voer een naam in voor de sheet set.", "Naam vereist",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            NameBox.Focus();
            return;
        }

        // Bouw bestandsnaam: vervang ongeldige tekens
        var fileName = name;
        foreach (var c in Path.GetInvalidFileNameChars())
            fileName = fileName.Replace(c, '_');
        fileName += ".dst";

        OutputPath = Path.Combine(FolderBox.Text.Trim(), fileName);

        if (File.Exists(OutputPath))
        {
            var r = MessageBox.Show(
                $"'{Path.GetFileName(OutputPath)}' bestaat al.\nOverschrijven?",
                "Bestand bestaat al", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (r != MessageBoxResult.Yes) return;
        }

        try
        {
            var valueOverrides = _previewProps.ToDictionary(d => d.Name, d => d.Value);
            var propertyNamesToKeep = _previewProps
                .Select(d => d.Name)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            SheetSetWriter.CreateFromTemplate(
                TemplateBox.Text.Trim(),
                OutputPath,
                name,
                DescBox.Text.Trim(),
                valueOverrides,
                propertyNamesToKeep);

            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Aanmaken mislukt:\n{ex.Message}", "Fout",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
