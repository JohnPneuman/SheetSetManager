using SheetSet.Core.Models;

namespace SheetSet.Core.Import;

public static class CsvTsvImporter
{
    private static readonly HashSet<string> StandardColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "Number", "Paginanummer", "Nummer",
        "Title", "Titel",
        "Description", "Omschrijving", "Desc",
        "Category", "Categorie",
        "RevisionNumber", "Revisienummer",
        "RevisionDate", "Revisiedatum",
        "IssuePurpose", "Uitgiftedoel",
        "DoNotPlot", "NietPlotten"
    };

    /// <summary>
    /// Parse a CSV or TSV file and match rows to sheets by Number.
    /// Returns SheetUpdateModel list for sheets that have a matching row.
    /// </summary>
    public static List<SheetUpdateModel> Import(string filePath, IEnumerable<SheetNode> allSheets)
    {
        var lines = File.ReadAllLines(filePath);
        if (lines.Length < 2) return [];

        var sep = DetectSeparator(lines[0]);
        var headers = SplitLine(lines[0], sep);

        var sheetsByNumber = allSheets
            .Where(s => s.Number != null)
            .ToDictionary(s => s.Number!, s => s, StringComparer.OrdinalIgnoreCase);

        var results = new List<SheetUpdateModel>();

        for (var i = 1; i < lines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i])) continue;

            var values = SplitLine(lines[i], sep);
            var row = BuildRow(headers, values);

            var number = row.GetValueOrDefault("Number")
                ?? row.GetValueOrDefault("Paginanummer")
                ?? row.GetValueOrDefault("Nummer");

            if (number == null || !sheetsByNumber.TryGetValue(number, out var sheetNode))
                continue;

            var update = new SheetUpdateModel { Sheet = sheetNode.Info };

            update.Title = row.GetValueOrDefault("Title") ?? row.GetValueOrDefault("Titel");
            update.Description = row.GetValueOrDefault("Description")
                ?? row.GetValueOrDefault("Omschrijving")
                ?? row.GetValueOrDefault("Desc");
            update.Category = row.GetValueOrDefault("Category") ?? row.GetValueOrDefault("Categorie");
            update.RevisionNumber = row.GetValueOrDefault("RevisionNumber") ?? row.GetValueOrDefault("Revisienummer");
            update.RevisionDate = row.GetValueOrDefault("RevisionDate") ?? row.GetValueOrDefault("Revisiedatum");
            update.IssuePurpose = row.GetValueOrDefault("IssuePurpose") ?? row.GetValueOrDefault("Uitgiftedoel");

            var doNotPlotRaw = row.GetValueOrDefault("DoNotPlot") ?? row.GetValueOrDefault("NietPlotten");
            if (doNotPlotRaw != null)
                update.DoNotPlot = doNotPlotRaw is "1" or "true" or "ja" or "yes";

            // All non-standard columns are treated as custom properties
            var customProps = row
                .Where(kv => !StandardColumns.Contains(kv.Key) && !string.IsNullOrWhiteSpace(kv.Value))
                .ToDictionary(kv => kv.Key, kv => kv.Value);

            if (customProps.Count > 0)
                update.CustomProperties = customProps;

            results.Add(update);
        }

        return results;
    }

    private static char DetectSeparator(string header)
        => header.Contains('\t') ? '\t' : ',';

    private static string[] SplitLine(string line, char sep)
    {
        if (sep == '\t')
            return line.Split('\t');

        // Basic CSV: handle quoted fields
        var fields = new List<string>();
        var inQuotes = false;
        var current = new System.Text.StringBuilder();

        foreach (var ch in line)
        {
            if (ch == '"') { inQuotes = !inQuotes; continue; }
            if (ch == ',' && !inQuotes) { fields.Add(current.ToString()); current.Clear(); continue; }
            current.Append(ch);
        }
        fields.Add(current.ToString());
        return [.. fields];
    }

    private static Dictionary<string, string> BuildRow(string[] headers, string[] values)
    {
        var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < headers.Length; i++)
        {
            var key = headers[i].Trim();
            var value = i < values.Length ? values[i].Trim() : string.Empty;
            if (!string.IsNullOrEmpty(key))
                row[key] = value;
        }
        return row;
    }
}
