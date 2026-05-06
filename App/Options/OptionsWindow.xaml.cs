using System.Windows;
using System.Windows.Controls;
using SheetSetEditor.Services;

namespace SheetSetEditor.Options;

public partial class OptionsWindow : Window
{
    private AppLanguage _originalLanguage;
    private bool        _originalDarkMode;

    private static readonly (AppLanguage Lang, string Display)[] LanguageItems =
    [
        (AppLanguage.Dutch,       "🇳🇱 Nederlands"),
        (AppLanguage.English,     "🇬🇧 English"),
        (AppLanguage.German,      "🇩🇪 Deutsch"),
        (AppLanguage.French,      "🇫🇷 Français"),
        (AppLanguage.Spanish,     "🇪🇸 Español"),
        (AppLanguage.Italian,     "🇮🇹 Italiano"),
        (AppLanguage.Portuguese,  "🇵🇹 Português"),
        (AppLanguage.Afrikaans,   "🇿🇦 Afrikaans"),
        (AppLanguage.Frisian,     "🏴 Frysk"),
        (AppLanguage.Klingon,     "🖖 tlhIngan Hol"),
    ];

    public OptionsWindow()
    {
        InitializeComponent();

        _originalLanguage = EditorSettingsService.GetLanguage();
        _originalDarkMode = EditorSettingsService.GetDarkMode();

        // Populate language combo
        foreach (var (lang, display) in LanguageItems)
            LanguageCombo.Items.Add(new ComboBoxItem { Content = display, Tag = lang });

        // Select current language
        for (int i = 0; i < LanguageCombo.Items.Count; i++)
        {
            if (((ComboBoxItem)LanguageCombo.Items[i]).Tag is AppLanguage l && l == _originalLanguage)
            {
                LanguageCombo.SelectedIndex = i;
                break;
            }
        }

        // Set theme radios
        if (_originalDarkMode) DarkRadio.IsChecked = true;
        else                    LightRadio.IsChecked = true;

        // General settings
        AutoBackupCheck.IsChecked = EditorSettingsService.GetAutoBackup();
        MaxRecentBox.Text         = EditorSettingsService.GetMaxRecentFiles().ToString();
    }

    private void LanguageCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LanguageCombo.SelectedItem is ComboBoxItem item && item.Tag is AppLanguage lang)
            LocalizationService.Instance.Language = lang;
    }

    private void Theme_Changed(object sender, RoutedEventArgs e)
    {
        bool dark = DarkRadio.IsChecked == true;
        App.ApplyTheme(dark);
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        bool dark = DarkRadio.IsChecked == true;
        var lang  = LanguageCombo.SelectedItem is ComboBoxItem ci && ci.Tag is AppLanguage l ? l : AppLanguage.Dutch;

        EditorSettingsService.SetLanguage(lang);
        EditorSettingsService.SetDarkMode(dark);
        EditorSettingsService.SetAutoBackup(AutoBackupCheck.IsChecked == true);

        if (int.TryParse(MaxRecentBox.Text.Trim(), out int max) && max is >= 1 and <= 100)
            EditorSettingsService.SetMaxRecentFiles(max);

        DialogResult = true;
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        // Revert previewed changes
        LocalizationService.Instance.Language = _originalLanguage;
        App.ApplyTheme(_originalDarkMode);
        DialogResult = false;
    }
}
