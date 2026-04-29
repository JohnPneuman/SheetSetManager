using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using ACSMCOMPONENTS24Lib;
using BoekSolutions.SheetSetEditor.Helpers;
using BoekSolutions.SheetSetEditor.Models;
using BoekSolutions.SheetSetEditor.ViewModels;
using BoekSolutions.SheetSetEditor.UI.Controls;
using System.ComponentModel;

namespace BoekSolutions.SheetSetEditor.UI
{
    public partial class MainWindow : Window
    {
        private IAcSmDatabase _database;
        private string _lastOpenedPath;

        public MainWindow()
        {
            LoggerInitializer.Init();
            InitializeComponent();
            SheetEditorHelper.AttachTreeSelection(SheetTreeControl.SheetTree);


            SheetEditorHelper.SheetSelected += id =>
            {
                if (id != null)
                    SheetPropertyGrid.LoadSheet(id);
                else
                    SheetPropertyGrid.LoadSharedProperties(TreeViewSelectionHelper.GetSelectedSheetIds());
            };

            AutoNumberControl.AutoNumberRequested += (s, options) =>
            {
                SheetEditorHelper.AutonumberSelected(
                SheetTreeControl.SheetTree,
                SheetTreeControl,
                options
            );
                MessageBox.Show("AutoNumber complete.");
            };
        }
        public void OnApplyClicked(object sender, RoutedEventArgs e)
        {
            var selected = TreeViewSelectionHelper.GetSelectedSheetIds();
            if (selected == null || selected.Count == 0)
            {
                Log.Warn("[SSM] Apply afgebroken: geen sheets geselecteerd.");
                return;
            }

            var vm = SheetPropertyGrid.StandardViewModel as SheetEditViewModel;
            if (vm == null)
            {
                Log.Warn("[SSM] Apply afgebroken: geen geldige ViewModel.");
                return;
            }

            string Safe(string val)
            {
                return string.IsNullOrWhiteSpace(val) || val == "…" ? null : val;
            }

            var updates = selected.Select(sheetId => new SheetUpdateModel
            {
                SheetId = sheetId,
                NewNumber = Safe(vm.NewPageNumber),
                Title = Safe(vm.NewTitle),
                Description = Safe(vm.NewDescription),

                // >>>>> Overige standaardvelden <<<<<
                RevisionNumber = Safe(vm.NewRevisionNumber),
                RevisionDate = Safe(vm.NewRevisionDate),
                IssuePurpose = Safe(vm.NewIssuePurpose),
                Category = Safe(vm.NewCategory),
                DoNotPlot = vm.NewDoNotPlot, // let op: bool, geen Safe!

                CustomProperties = SheetPropertyGrid.CustomProperties?
                    .Where(p => Safe(p.PropertyValue) != null)
                    .ToDictionary(p => p.PropertyName, p => Safe(p.PropertyValue))
                    ?? new Dictionary<string, string>()
            }).ToList();

            SheetEditorHelper.ApplyUpdates(updates, SheetTreeControl.SheetTree);
            Log.Info($"[SSM] Apply uitgevoerd op {updates.Count} sheet(s).");
        }


        public void OnSaveClicked(object sender, RoutedEventArgs e)
        {
            if (_database == null)
            {
                Log.Info("[SSM] Save afgebroken: database is null.");
                return;
            }

            if (!DatabaseHelper.IsLocked(_database))
            {
                Log.Info("[SSM] Kan niet saven: database is niet gelocked.");
                return;
            }

            Log.Info("[SSM] >>> PRE-SAVE DEBUG START <<<");
            Log.Info("[SSM] >>> PRE-SAVE DEBUG END <<<");

            SaveHelper.SaveDatabase(_database);
            DatabaseHelper.UnlockDb(_database);
            _database = null;
            SheetTreeControl.ClearTree();
            OnSheetSetOpen(_lastOpenedPath);

            SheetPropertyGrid.Clear();
        }

        private void LoadTreeFromDatabase(IAcSmDatabase db)
        {
            TryHelper.Run(() =>
            {
                SheetTreeControl.ClearTree();

                var sheetSet = db.GetSheetSet();
                if (sheetSet == null)
                {
                    Log.Info("[SSM] db.GetSheetSet() returned null!");
                    return;
                }

                var sheetSetId = SheetHelper.GetObjectId(sheetSet);
                if (sheetSetId == null || !sheetSetId.IsValid())
                {
                    Log.Info("[SSM] SheetSetId is null of ongeldig!");
                    return;
                }

                Log.Info("[SSM] Boomstructuur opbouwen...");
                SheetTreeControl.LoadSheetSet(sheetSetId);
            });
        }

        public void OnSheetSetOpen(string path)
        {
            _lastOpenedPath = path;

            Log.Info($"[SSM] Sheet Set openen: {path}");

            _database = SheetSetMgr.OpenDatabase(path, false);
            if (_database == null)
            {
                Log.Info("[SSM] Openen mislukt: database is null");
                return;
            }

            var locked = DatabaseHelper.TryLock(_database);
            if (!locked)
            {
                Log.Info("[SSM] Database kon niet gelocked worden.");
                return;
            }

            RecentFilesHelper.AddToRecent(path);
            SheetSetMenu.SetActiveSheetSet(path);

            LoadTreeFromDatabase(_database);

            Log.Info("[SSM] Sheet Set geladen en gelocked.");
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            OnSheetSetClose();
        }

        public void OnSheetSetClose()
        {
            // Eerst de database netjes unlocken (en evt. saven)
            if (_database != null)
            {
                if (DatabaseHelper.IsLocked(_database))
                {
                    DatabaseHelper.UnlockDb(_database);
                    Log.Info("[SSM] Database ge-unlocked bij sluiten.");
                }
                _database = null;
                // Forceren van garbage collection kan soms helpen bij COM-lekken:
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            SheetPropertyGrid.Clear();
            SheetTreeControl.ClearTree();
            SheetSetMenu.ClearActiveSheetSet();
            Log.Info("[SSM] Sheet Set gesloten.");
        }
    }
}
