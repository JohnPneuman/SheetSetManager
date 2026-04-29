using System;
using System.Windows;

namespace BoekSolutions.SheetSetEditor.UI.Dialogs
{
    public partial class AutoNumberOptionsDialog : Window
    {
        public int StartNumber { get; private set; } = 1;
        public int Increment { get; private set; } = 1;
        public string Prefix { get; private set; } = string.Empty;
        public string Suffix { get; private set; } = string.Empty;

        public AutoNumberOptionsDialog()
        {
            InitializeComponent();
        }

        private void OnOkClicked(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(StartNumberBox.Text, out int start))
            {
                MessageBox.Show("Start number must be a valid number.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(IncrementBox.Text, out int increment))
            {
                MessageBox.Show("Increment must be a valid number.", "Input Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            StartNumber = start;
            Increment = increment;
            Prefix = PrefixBox.Text;
            Suffix = SuffixBox.Text;

            DialogResult = true;
            Close();
        }
    }
}