// Nieuwe ApplyUpdate helper + model

using System;
using System.Collections.Generic;
using ACSMCOMPONENTS24Lib;
using BoekSolutions.SheetSetEditor.Models;

namespace BoekSolutions.SheetSetEditor.Models
{
    public class SheetUpdateModel
    {
        public IAcSmObjectId SheetId { get; set; }
        public string NewNumber { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }

        // Nieuwe standaardvelden:
        public string RevisionNumber { get; set; }
        public string RevisionDate { get; set; }
        public string IssuePurpose { get; set; }
        public string Category { get; set; }
        public bool? DoNotPlot { get; set; } // Nullable zodat je 'geen wijziging' kan aanduiden

        public Dictionary<string, string> CustomProperties { get; set; } = new Dictionary<string, string>();
    }

}
