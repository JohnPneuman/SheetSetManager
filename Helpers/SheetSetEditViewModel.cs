using ACSMCOMPONENTS24Lib;
using BoekSolutions.SheetSetEditor.Helpers;

namespace BoekSolutions.SheetSetEditor.ViewModels
{
    public class SheetSetEditViewModel
    {
        public IAcSmObjectId SheetSetId { get; set; }

        public string OriginalName { get; set; }
        public string NewName { get; set; }

        public string OriginalDescription { get; set; }
        public string NewDescription { get; set; }

        public SheetSetEditViewModel(IAcSmObjectId id)
        {
            SheetSetId = id;
            TryHelper.Run(() =>
            {
                using (var scope = new ComScope())
                {
                    var set = scope.Track(ComHelper.GetObject<IAcSmSheetSet>(id));
                    if (set != null)
                    {
                        OriginalName = NewName = set.GetName() ?? string.Empty;
                        OriginalDescription = NewDescription = set.GetDesc() ?? string.Empty;
                    }
                }
            });
        }

        public SheetSetEditViewModel() { }
    }
}