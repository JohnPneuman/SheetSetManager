using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using BoekSolutions.SheetSetEditor.Helpers;
using SheetSetEditor.Services;
using static BoekSolutions.SheetSetEditor.Aliases;

namespace BoekSolutions.SheetSetEditor.UI.Controls
{
    public partial class SheetSetMenuButton : UserControl
    {
        private string _activePath = null;

        public event Action<string> SheetSetOpenRequested;
        public event Action SheetSetCloseRequested;

        public SheetSetMenuButton()
        {
            InitializeComponent();
            UpdateButton("Open...", null);
        }

        public void SetActiveSheetSet(string path)
        {
            _activePath = path;
            var name = Path.GetFileNameWithoutExtension(path);
            UpdateButton(name, path);
        }

        public void ClearActiveSheetSet()
        {
            _activePath = null;
            UpdateButton("Open...", null);
        }

        private void UpdateButton(string label, string tooltip)
        {
            ButtonLabel.Text = label;
            MainButton.ToolTip = tooltip;
        }

        private void MainButton_Click(object sender, RoutedEventArgs e)
        {
            var menu = new ContextMenu();

            if (_activePath != null)
            {
                var activeItem = new MenuItem
                {
                    Header = Path.GetFileNameWithoutExtension(_activePath),
                    ToolTip = _activePath
                };

                var closeItem = new MenuItem
                {
                    Header = LocalizationService.T("CloseSheetSet")
                };
                closeItem.Click += (s, ev) => OnCloseActive();

                activeItem.Items.Add(closeItem); // 👈 Voeg toe als submenu

                menu.Items.Add(activeItem);
                menu.Items.Add(new Separator());
            }

            var recentMenu = new MenuItem { Header = LocalizationService.T("RecentFiles") };
            foreach (var path in RecentFilesHelper.GetRecentSheetSets())
            {
                var item = new MenuItem
                {
                    Header = Path.GetFileNameWithoutExtension(path),
                    ToolTip = path
                };
                item.Click += (s, ev) => OnOpen(path);
                recentMenu.Items.Add(item);
            }
            menu.Items.Add(recentMenu);

            menu.Items.Add(new MenuItem
            {
                Header = LocalizationService.T("New"),
                Command = new RelayCommand(() => Log.Info("[SSM] New Sheet Set clicked (nog niet geïmplementeerd)."))
            });

            var openItem = new MenuItem { Header = LocalizationService.T("Open") };
            openItem.Click += (s, ev) => OnOpenDialog();
            menu.Items.Add(openItem);

            menu.PlacementTarget = MainButton;
            menu.IsOpen = true;
        }

        private void OnOpen(string path)
        {
            Log.Info($"[SSM] Open Sheet Set: {path}");
            SheetSetOpenRequested?.Invoke(path);
        }

        private void OnOpenDialog()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = LocalizationService.T("SheetSetFilter"),
                Title  = LocalizationService.T("OpenSheetSetTitle")
            };

            if (dlg.ShowDialog() == true)
            {
                OnOpen(dlg.FileName);
            }
        }

        private void OnCloseActive()
        {
            Log.Info("[SSM] Close Sheet Set clicked");
            SheetSetCloseRequested?.Invoke();
        }

        public void ReloadRecent()
        {
            // eventueel leegmaken als je caching gebruikt
        }

    }
}
