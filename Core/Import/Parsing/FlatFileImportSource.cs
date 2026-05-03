using System.Text;
using SheetSet.Core.Import.Abstractions;
using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Parsing;

/// <summary>
/// Handles .csv, .tsv, .txt, .tab in two layouts:
///   Horizontal — one record per line, first line = headers (traditional CSV/TSV)
///   Vertical   — one value per line, line position = field index (acad_oh/bh format)
///
/// Vertical detection: fewer than half the non-empty lines contain the best delimiter,
/// meaning each line is a single value without any column separator.
/// </summary>
public class FlatFileImportSource : IImportSource
{
    private static readonly HashSet<string> SupportedExtensions =
        new(StringComparer.OrdinalIgnoreCase) { ".csv", ".tsv", ".txt", ".tab" };

    public bool CanHandle(string filePath)
        => SupportedExtensions.Contains(Path.GetExtension(filePath));

    public ImportPreview GetPreview(string filePath, SourceHints? hints = null, int maxRows = 8)
    {
        var preview = new ImportPreview { FilePath = filePath };
        try
        {
            var encoding = ResolveEncoding(filePath, hints);

            // Read enough lines to detect format; for Vertical we may need all lines
            var sampleLines = ReadLines(filePath, encoding, hints?.SkipRows ?? 0, 20).ToList();

            if (sampleLines.Count == 0)
            {
                preview.ValidationIssues.Add(Issue(ValidationSeverity.Error, "Het bestand is leeg."));
                return preview;
            }

            var delimiter = ResolveDelimiter(hints, sampleLines);
            var layout    = ResolveLayout(hints, sampleLines, delimiter);

            preview.DetectedHints = new SourceHints
            {
                Delimiter = layout == FieldLayout.Horizontal ? delimiter : null,
                Encoding  = encoding.WebName,
                HasHeader = layout == FieldLayout.Horizontal,
                Layout    = layout,
                SkipRows  = hints?.SkipRows ?? 0
            };

            if (layout == FieldLayout.Vertical)
                PopulateVertical(preview, sampleLines);
            else
                PopulateHorizontal(preview, sampleLines, delimiter, maxRows);
        }
        catch (Exception ex)
        {
            preview.ValidationIssues.Add(Issue(ValidationSeverity.Error, $"Fout bij lezen: {ex.Message}"));
        }
        return preview;
    }

    public IEnumerable<NormalizedRow> ReadRows(string filePath, SourceHints? hints = null)
    {
        var encoding = ResolveEncoding(filePath, hints);
        var allLines = ReadLines(filePath, encoding, hints?.SkipRows ?? 0).ToList();

        if (allLines.Count == 0) yield break;

        var probe     = allLines.Take(8).ToList();
        var delimiter = ResolveDelimiter(hints, probe);
        var layout    = ResolveLayout(hints, probe, delimiter);

        if (layout == FieldLayout.Vertical)
        {
            // Keep all lines (including empty ones) — position = field index
            yield return BuildVerticalRow(allLines, filePath);
            yield break;
        }

        var nonEmpty = allLines.Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
        if (nonEmpty.Count < 2) yield break;

        var headers = SplitLine(nonEmpty[0], delimiter);
        for (var i = 1; i < nonEmpty.Count; i++)
            yield return BuildHorizontalRow(headers, nonEmpty[i], delimiter, i, filePath);
    }

    // ── Preview builders ──────────────────────────────────────────────────────

    private static void PopulateVertical(ImportPreview preview, List<string> lines)
    {
        var row = BuildVerticalRow(lines, preview.FilePath);

        // Headers are the field indices: "1", "2", ... matching the number of lines
        preview.Headers = Enumerable.Range(1, lines.Count).Select(i => i.ToString()).ToList();
        preview.SampleRows.Add(row);
    }

    private static void PopulateHorizontal(ImportPreview preview, List<string> lines,
        char delimiter, int maxRows)
    {
        var nonEmpty = lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
        if (nonEmpty.Count == 0) return;

        var headers = SplitLine(nonEmpty[0], delimiter);
        preview.Headers = [.. headers];
        ValidateHeaders(headers, preview.ValidationIssues);

        for (var i = 1; i < nonEmpty.Count && preview.SampleRows.Count < maxRows; i++)
        {
            var values = SplitLine(nonEmpty[i], delimiter);
            if (values.Length != headers.Length)
                preview.ValidationIssues.Add(Issue(ValidationSeverity.Warning,
                    $"Regel {i + 1} heeft {values.Length} kolom(men), verwacht {headers.Length}.",
                    rowIndex: i + 1));

            preview.SampleRows.Add(BuildHorizontalRow(headers, nonEmpty[i], delimiter, i, preview.FilePath));
        }
    }

    // ── Row builders ──────────────────────────────────────────────────────────

    /// <summary>
    /// Vertical: each line is one field value; the 1-based line position is the field key.
    /// Empty lines are kept so field positions remain correct.
    /// </summary>
    private static NormalizedRow BuildVerticalRow(List<string> lines, string filePath)
    {
        var fields = new Dictionary<string, string>(StringComparer.Ordinal);
        for (var i = 0; i < lines.Count; i++)
            fields[(i + 1).ToString()] = lines[i].Trim();

        return new NormalizedRow { RowIndex = 0, SourceFile = filePath, Fields = fields };
    }

    private static NormalizedRow BuildHorizontalRow(
        string[] headers, string line, char delimiter, int rowIndex, string filePath)
    {
        var values = SplitLine(line, delimiter);
        var fields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < headers.Length; i++)
        {
            var key   = headers[i].Trim();
            var value = i < values.Length ? values[i].Trim() : string.Empty;
            if (!string.IsNullOrEmpty(key))
                fields[key] = value;
        }
        return new NormalizedRow { RowIndex = rowIndex, SourceFile = filePath, Fields = fields };
    }

    // ── Detection helpers ─────────────────────────────────────────────────────

    private static Encoding ResolveEncoding(string filePath, SourceHints? hints)
    {
        if (!string.IsNullOrEmpty(hints?.Encoding))
            return Encoding.GetEncoding(hints.Encoding);
        return EncodingDetector.Detect(filePath);
    }

    private static char ResolveDelimiter(SourceHints? hints, IReadOnlyList<string> sampleLines)
    {
        if (hints?.Delimiter != null) return hints.Delimiter.Value;
        return DelimiterDetector.Detect(sampleLines);
    }

    private static FieldLayout ResolveLayout(SourceHints? hints, IReadOnlyList<string> sampleLines, char delimiter)
    {
        if (hints?.Layout != null) return hints.Layout.Value;
        return DelimiterDetector.DetectLayout(sampleLines, delimiter);
    }

    // ── CSV line splitter (handles quoted fields for comma-delimited files) ───

    private static string[] SplitLine(string line, char delimiter)
    {
        if (delimiter != ',') return line.Split(delimiter);

        var fields   = new List<string>();
        var inQuotes = false;
        var current  = new StringBuilder();

        foreach (var ch in line)
        {
            if (ch == '"')        { inQuotes = !inQuotes; continue; }
            if (ch == ',' && !inQuotes) { fields.Add(current.ToString()); current.Clear(); continue; }
            current.Append(ch);
        }
        fields.Add(current.ToString());
        return [.. fields];
    }

    // ── Validation ────────────────────────────────────────────────────────────

    private static void ValidateHeaders(string[] headers, List<ValidationIssue> issues)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < headers.Length; i++)
        {
            var h = headers[i].Trim();
            if (string.IsNullOrEmpty(h))
                issues.Add(Issue(ValidationSeverity.Warning,
                    $"Lege kolomnaam op positie {i + 1}.", column: $"Col_{i + 1}", rowIndex: 1));
            else if (!seen.Add(h))
                issues.Add(Issue(ValidationSeverity.Warning,
                    $"Dubbele kolomnaam '{h}'.", column: h, rowIndex: 1));
        }
    }

    // ── Line reader ───────────────────────────────────────────────────────────

    private static IEnumerable<string> ReadLines(string filePath, Encoding encoding,
        int skipRows, int? maxLines = null)
    {
        var yielded = 0;
        foreach (var line in File.ReadLines(filePath, encoding).Skip(skipRows))
        {
            if (maxLines.HasValue && yielded >= maxLines.Value) yield break;
            yield return line;
            yielded++;
        }
    }

    private static ValidationIssue Issue(ValidationSeverity severity, string message,
        int? rowIndex = null, string? column = null)
        => new() { Severity = severity, Message = message, RowIndex = rowIndex, Column = column };
}
