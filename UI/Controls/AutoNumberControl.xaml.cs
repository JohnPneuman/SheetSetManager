using System;
using System.Windows;
using System.Windows.Controls;
using BoekSolutions.SheetSetEditor.UI.Dialogs;
using BoekSolutions.SheetSetEditor.Models;

namespace BoekSolutions.SheetSetEditor.UI.Controls
{
    public partial class AutoNumberControl : UserControl
    {
        public event EventHandler<AutoNumberOptions> AutoNumberRequested;

        public AutoNumberControl()
        {
            InitializeComponent();
        }

        private void OnAutoNumberClicked(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(StartNumberBox.Text, out int startNumber))
            {
                MessageBox.Show("Please enter a valid start number.");
                return;
            }

            AutoNumberRequested?.Invoke(this, new AutoNumberOptions
            {
                StartNumber = startNumber,
                Increment = 1,
                Prefix = string.Empty,
                Suffix = string.Empty
            });
        }

        private void OnOptionsClicked(object sender, RoutedEventArgs e)
        {
            var dialog = new AutoNumberOptionsDialog
            {
                ResizeMode = ResizeMode.NoResize,
                Width = 220,
                Height = 220,
                WindowStartupLocation = WindowStartupLocation.Manual
            };

            if (sender is Button btn)
            {
                var point = btn.PointToScreen(new Point(btn.ActualWidth, btn.ActualHeight));
                dialog.Left = point.X - dialog.Width;
                dialog.Top = point.Y;
            }

            if (dialog.ShowDialog() == true)
            {
                AutoNumberRequested?.Invoke(this, new AutoNumberOptions
                {
                    StartNumber = dialog.StartNumber,
                    Increment = dialog.Increment,
                    Prefix = dialog.Prefix,
                    Suffix = dialog.Suffix
                });
            }
        }
    }
}
