using System.Text;
using SheetSet.Core.Models;

namespace SheetSet.Core.Import.Export;

public static class CsvExporter
{
    public static void ExportToFile(SheetSetDocument doc, string outputPath)
        => File.WriteAllText(outputPath, Build(doc), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

    public static string Build(SheetSetDocument doc)
    {
        var sheetProps = doc.Info.CustomPropertyDefinitions.Where(d => d.Flags != 1).Select(d => d.Name).ToList();
        var setProps   = doc.Info.CustomPropertyDefinitions.Where(d => d.Flags == 1).Select(d => d.Name).ToList();

        var sb = new StringBuilder();

        // Header row
        var stdSheetHdrs = new[] { "Nummer", "Titel", "Omschrijving", "Revisienummer", "Revisiedatum", "Doel", "Categorie" };
        var stdSetHdrs   = new[] { "SheetSet.Naam", "SheetSet.Omschrijving", "SheetSet.Projectnaam", "SheetSet.Projectnummer", "SheetSet.Fase", "SheetSet.Mijlpaal" };
        sb.AppendLine(string.Join(",", stdSheetHdrs.Concat(sheetProps).Concat(stdSetHdrs).Concat(setProps).Select(QuoteCsv)));

        // SheetSet row (first data row)
        var setRow = new List<string>(new string[stdSheetHdrs.Length + sheetProps.Count]);
        for (var i = 0; i < setRow.Count; i++) setRow[i] = string.Empty;
        setRow.Add(doc.Info.Name             ?? string.Empty);
        setRow.Add(doc.Info.Description      ?? string.Empty);
        setRow.Add(doc.Info.ProjectName      ?? string.Empty);
        setRow.Add(doc.Info.ProjectNumber    ?? string.Empty);
        setRow.Add(doc.Info.ProjectPhase     ?? string.Empty);
        setRow.Add(doc.Info.ProjectMilestone ?? string.Empty);
        foreach (var prop in setProps)
            setRow.Add(doc.Info.CustomProperties.TryGetValue(prop, out var sv) ? sv : string.Empty);
        sb.AppendLine(string.Join(",", setRow.Select(QuoteCsv)));

        // Sheet rows
        var emptySetCols = new string[stdSetHdrs.Length + setProps.Count];
        for (var i = 0; i < emptySetCols.Length; i++) emptySetCols[i] = string.Empty;

        foreach (var sheet in doc.GetAllSheets())
        {
            var values = new List<string>
            {
                sheet.Info.Number         ?? string.Empty,
                sheet.Info.Title          ?? string.Empty,
                sheet.Info.Description    ?? string.Empty,
                sheet.Info.RevisionNumber ?? string.Empty,
                sheet.Info.RevisionDate   ?? string.Empty,
                sheet.Info.IssuePurpose   ?? string.Empty,
                sheet.Info.Category       ?? string.Empty
            };
            foreach (var prop in sheetProps)
                values.Add(sheet.Info.CustomProperties.TryGetValue(prop, out var v) ? v : string.Empty);
            values.AddRange(emptySetCols);
            sb.AppendLine(string.Join(",", values.Select(QuoteCsv)));
        }

        return sb.ToString();
    }

    private static string QuoteCsv(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }
}
