using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Controls;
using ACSMCOMPONENTS24Lib;
using BoekSolutions.SheetSetEditor.Helpers;
using BoekSolutions.SheetSetEditor.Models;
using BoekSolutions.SheetSetEditor.ViewModels;
using System.Collections.Generic;

namespace BoekSolutions.SheetSetEditor.UI.Controls
{
    public partial class SheetPropertyGrid : UserControl
    {
        private IAcSmObjectId _currentSheetId;
        private SheetEditViewModel _standardViewModel;
        private ObservableCollection<CustomPropertyViewModel> _customProperties;

        public ObservableCollection<CustomPropertyViewModel> CustomProperties => _customProperties;
        public SheetEditViewModel StandardViewModel => _standardViewModel;

        public SheetPropertyGrid()
        {
            InitializeComponent();
        }

        public void LoadSheet(IAcSmObjectId id)
        {
            _currentSheetId = id;

            TryHelper.Run(() =>
            {
                using (var scope = new ComScope())
                {
                    var sheet = scope.Track(ComHelper.GetObject<IAcSmSheet>(id));

                    if (sheet != null)
                    {
                        Log.Info("[SSM] Geselecteerd object is een SHEET.");

                        _standardViewModel = new SheetEditViewModel(id);
                        DataContext = _standardViewModel;

                        _customProperties = PropertyBagHelper.GetCustomProperties(id);
                        CustomPropertiesList.ItemsSource = _customProperties;

                        Log.Info($"[SSM] Custom props geladen: {_customProperties?.Count ?? 0}");
                    }
                    else
                    {
                        Log.Warn("[SSM] Geen geldige sheet gevonden voor geladen ID.");
                        Clear();
                    }
                }
            });
        }

        public void Clear()
        {
            _currentSheetId = null;
            _standardViewModel = null;
            _customProperties = null;

            DataContext = null;
            CustomPropertiesList.ItemsSource = null;
        }

        public void LoadSharedProperties(List<IAcSmObjectId> ids)
        {
            if (ids == null || ids.Count == 0) return;

            var pageNumbers = new List<string>();
            var titles = new List<string>();
            var descriptions = new List<string>();
            var revisionNumbers = new List<string>();
            var revisionDates = new List<string>();
            var issuePurposes = new List<string>();
            var categories = new List<string>();
            var doNotPlots = new List<bool>();

            foreach (var id in ids)
            {
                TryHelper.Run(() =>
                {
                    using (var scope = new ComScope())
                    {
                        var sheet = scope.Track(ComHelper.GetObject<IAcSmSheet>(id));
                        if (sheet == null) return;

                        pageNumbers.Add(sheet.GetNumber() ?? "");
                        titles.Add(sheet.GetTitle() ?? "");
                        descriptions.Add(sheet.GetDesc() ?? "");

                        var sheet2 = sheet as IAcSmSheet2;
                        if (sheet2 != null)
                        {
                            revisionNumbers.Add(sheet2.GetRevisionNumber() ?? "");
                            revisionDates.Add(sheet2.GetRevisionDate() ?? "");
                            issuePurposes.Add(sheet2.GetIssuePurpose() ?? "");
                            categories.Add(sheet2.GetCategory() ?? "");
                        }
                        else
                        {
                            revisionNumbers.Add("");
                            revisionDates.Add("");
                            issuePurposes.Add("");
                            categories.Add("");
                        }
                        doNotPlots.Add(sheet.GetDoNotPlot());
                    }
                });
            }

            _standardViewModel = new SheetEditViewModel(null)
            {
                NewPageNumber = SheetEditorHelper.GetSharedValue(pageNumbers),
                NewTitle = SheetEditorHelper.GetSharedValue(titles),
                NewDescription = SheetEditorHelper.GetSharedValue(descriptions),
                NewRevisionNumber = SheetEditorHelper.GetSharedValue(revisionNumbers),
                NewRevisionDate = SheetEditorHelper.GetSharedValue(revisionDates),
                NewIssuePurpose = SheetEditorHelper.GetSharedValue(issuePurposes),
                NewCategory = SheetEditorHelper.GetSharedValue(categories),
                NewDoNotPlot = SheetEditorHelper.GetSharedBoolValue(doNotPlots)
            };

            DataContext = _standardViewModel;

            var sharedProps = SheetEditorHelper.GetSharedCustomProperties(ids);
            _customProperties = new ObservableCollection<CustomPropertyViewModel>(
                sharedProps.Select(kvp => new CustomPropertyViewModel
                {
                    PropertyName = kvp.Key,
                    PropertyValue = kvp.Value
                })
            );
            CustomPropertiesList.ItemsSource = _customProperties;
        }
    }
}
