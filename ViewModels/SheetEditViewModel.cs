using ACSMCOMPONENTS24Lib;
using BoekSolutions.SheetSetEditor.Helpers;
using System;

namespace BoekSolutions.SheetSetEditor.ViewModels
{
    public class SheetEditViewModel
    {
        public IAcSmObjectId SheetId { get; set; }

        // Page Number
        public string OriginalPageNumber { get; set; }
        public string NewPageNumber { get; set; }

        // Title
        public string OriginalTitle { get; set; }
        public string NewTitle { get; set; }

        // Description
        public string OriginalDescription { get; set; }
        public string NewDescription { get; set; }

        // Revision Number
        public string OriginalRevisionNumber { get; set; }
        public string NewRevisionNumber { get; set; }

        // Revision Date
        public string OriginalRevisionDate { get; set; }
        public string NewRevisionDate { get; set; }

        // Issue Purpose
        public string OriginalIssuePurpose { get; set; }
        public string NewIssuePurpose { get; set; }

        // Category
        public string OriginalCategory { get; set; }
        public string NewCategory { get; set; }

        // Do Not Plot (boolean, meestal een checkbox)
        public bool OriginalDoNotPlot { get; set; }
        public bool NewDoNotPlot { get; set; }

        // FileName (readonly, maar kun je wel tonen)
        public string FileName { get; set; }

        // Layout Name (optioneel)
        public string LayoutName { get; set; }

        // Constructor
        public SheetEditViewModel(IAcSmObjectId id)
        {
            SheetId = id;
            TryHelper.Run(() =>
            {
                using (var scope = new ComScope())
                {
                    var sheet = scope.Track(ComHelper.GetObject<IAcSmSheet>(id));
                    if (sheet != null)
                    {
                        OriginalPageNumber = NewPageNumber = sheet.GetNumber() ?? string.Empty;
                        OriginalTitle = NewTitle = sheet.GetTitle() ?? string.Empty;
                        OriginalDescription = NewDescription = sheet.GetDesc() ?? string.Empty;
                        OriginalDoNotPlot = NewDoNotPlot = sheet.GetDoNotPlot();

                        FileName = sheet.GetLayout().GetFileName() ?? string.Empty;

                        // Layout name uitlezen (optioneel)
                        var layoutRef = sheet.GetLayout();
                        LayoutName = layoutRef != null ? layoutRef.GetName() : string.Empty;

                        // Sheet2 extra velden (revision, issue, category)
                        var sheet2 = sheet as IAcSmSheet2;
                        if (sheet2 != null)
                        {
                            OriginalRevisionNumber = NewRevisionNumber = sheet2.GetRevisionNumber() ?? string.Empty;
                            OriginalRevisionDate = NewRevisionDate = sheet2.GetRevisionDate() ?? string.Empty;
                            OriginalIssuePurpose = NewIssuePurpose = sheet2.GetIssuePurpose() ?? string.Empty;
                            OriginalCategory = NewCategory = sheet2.GetCategory() ?? string.Empty;
                        }
                        else
                        {
                            // Sheet2 niet beschikbaar? Leeg laten
                            OriginalRevisionNumber = NewRevisionNumber = string.Empty;
                            OriginalRevisionDate = NewRevisionDate = string.Empty;
                            OriginalIssuePurpose = NewIssuePurpose = string.Empty;
                            OriginalCategory = NewCategory = string.Empty;
                        }
                    }
                }
            });
        }

        public SheetEditViewModel() { }
    }
}