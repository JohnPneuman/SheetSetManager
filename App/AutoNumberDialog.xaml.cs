using System.Windows;

namespace SheetSetEditor;

public partial class AutoNumberDialog : Window
{
    public int    StartNumber { get; private set; } = 1;
    public int    Increment   { get; private set; } = 1;
    public string Prefix      { get; private set; } = string.Empty;
    public string Suffix      { get; private set; } = string.Empty;

    public AutoNumberDialog()
    {
        InitializeComponent();
        Loaded += (_, _) => StartBox.Focus();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (!int.TryParse(StartBox.Text, out int start) || start < 0)
        {
            MessageBox.Show("Voer een geldig startnummer in.", "Ongeldig invoer",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (!int.TryParse(IncrementBox.Text, out int inc) || inc < 1)
        {
            MessageBox.Show("Voer een geldige stap in (≥ 1).", "Ongeldig invoer",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        StartNumber = start;
        Increment   = inc;
        Prefix      = PrefixBox.Text;
        Suffix      = SuffixBox.Text;
        DialogResult = true;
        Close();
    }
}
