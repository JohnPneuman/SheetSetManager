using System.Windows;
using SheetSetEditor.Services;

namespace SheetSetEditor;

public partial class AddCustomPropertyDialog : Window
{
    public string PropertyName { get; private set; } = string.Empty;
    public int    Flags        { get; private set; } = 2;

    public AddCustomPropertyDialog()
    {
        InitializeComponent();
        NameBox.Focus();
    }

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        var name = NameBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            MessageBox.Show(LocalizationService.T("EnterPropertyName"), LocalizationService.T("NameRequired"),
                MessageBoxButton.OK, MessageBoxImage.Warning);
            NameBox.Focus();
            return;
        }

        PropertyName = name;
        Flags        = RbBladVeld.IsChecked == true ? 2 : 1;
        DialogResult = true;
        Close();
    }
}
