using BoekSolutions.SheetSetEditor.Helpers;

namespace BoekSolutions.SheetSetEditor.Helpers
{
    public static class SheetDisplayHelper
    {
        /// <summary>
        /// Format: "nummer – titel", of alleen nummer als titel leeg is.
        /// </summary>
        public static string FormatSheetLabel(string number, string title)
        {
            string safeNumber = number ?? "";
            string safeTitle = title ?? "";

            if (string.IsNullOrWhiteSpace(safeTitle))
            {
                Log.Debug($"[Display] Alleen nummer getoond: {safeNumber}");
                return safeNumber;
            }

            var label = $"{safeNumber} – {safeTitle}";
            Log.Debug($"[Display] Combinatie: {label}");
            return label;
        }
    }
}
