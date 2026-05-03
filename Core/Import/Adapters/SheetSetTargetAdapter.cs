using SheetSet.Core.Import.Abstractions;
using SheetSet.Core.Import.Models;
using SheetSet.Core.Models;
using SheetSet.Core.Parsing;

namespace SheetSet.Core.Import.Adapters;

/// <summary>
/// Writes MappedRow values to the sheet set's in-memory XElement tree — no AutoCAD COM required.
/// Flags=1 custom properties and SheetSet.* standard fields are written to the document root;
/// everything else (Flags=2 or standard per-sheet fields) is written to individual sheets.
/// </summary>
public class SheetSetTargetAdapter : ITargetAdapter
{
    private readonly List<SheetInfo> _targetSheets;
    private readonly SheetSetDocument _document;

    // "SheetSet.<display>" → XML propname used by XmlUtil.SetPropValue on the root element
    private static readonly Dictionary<string, string> StdSetFields = new(StringComparer.OrdinalIgnoreCase)
    {
        ["SheetSet.Naam"]         = "Name",
        ["SheetSet.Omschrijving"] = "Desc",
        ["SheetSet.Projectnaam"]  = "ProjectName",
        ["SheetSet.Projectnummer"] = "ProjectNumber",
        ["SheetSet.Fase"]         = "ProjectPhase",
        ["SheetSet.Mijlpaal"]     = "ProjectMilestone",
    };

    // "Nummer" etc. → lambda that sets the value on a SheetInfo
    private static readonly Dictionary<string, Action<SheetInfo, string>> StdSheetFields =
        new(StringComparer.OrdinalIgnoreCase)
    {
        ["Nummer"]       = (s, v) => s.Number         = v,
        ["Titel"]        = (s, v) => s.Title          = v,
        ["Omschrijving"] = (s, v) => s.Description    = v,
        ["Revisienummer"] = (s, v) => s.RevisionNumber = v,
        ["Revisiedatum"]  = (s, v) => s.RevisionDate   = v,
        ["Doel"]         = (s, v) => s.IssuePurpose   = v,
        ["Categorie"]    = (s, v) => s.Category       = v,
    };

    public SheetSetTargetAdapter(List<SheetInfo> targetSheets, SheetSetDocument document)
    {
        _targetSheets = targetSheets ?? throw new ArgumentNullException(nameof(targetSheets));
        _document     = document     ?? throw new ArgumentNullException(nameof(document));
    }

    public List<DryRunChange> PreviewChanges(IEnumerable<MappedRow> rows)
    {
        var changes = new List<DryRunChange>();
        var rowList = rows.ToList();

        // Set-level: collect from ALL rows, skip empty values (so a sheet row's blank SheetSet column doesn't show as a change)
        var seenSetProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rowList.Where(r => !r.HasErrors))
        {
            foreach (var (property, newValue) in row.TargetValues.Where(kv => !string.IsNullOrEmpty(kv.Value) && IsSetLevelProperty(kv.Key)))
            {
                if (!seenSetProps.Add(property)) continue; // first non-empty value wins
                if (StdSetFields.ContainsKey(property))
                {
                    var old = GetStdSetValue(property);
                    if (old != newValue) changes.Add(new DryRunChange { SheetLabel = "(SheetSet)", Property = property, OldValue = old, NewValue = newValue });
                }
                else
                {
                    var old = _document.Info.CustomProperties.TryGetValue(property, out var v) ? v : string.Empty;
                    if (old != newValue) changes.Add(new DryRunChange { SheetLabel = "(SheetSet)", Property = property, OldValue = old, NewValue = newValue });
                }
            }
        }

        // Sheet-level: only rows that have at least one non-empty sheet-level value
        var sheetRows = rowList.Where(r => !r.HasErrors && HasNonEmptySheetLevel(r)).ToList();
        foreach (var (row, sheets) in BuildAssignments(sheetRows, _targetSheets))
        {
            if (row.HasErrors) continue;
            foreach (var (property, newValue) in row.TargetValues.Where(kv => IsSheetLevelProperty(kv.Key)))
            {
                foreach (var sheet in sheets)
                {
                    var old = StdSheetFields.ContainsKey(property)
                        ? GetStdSheetValue(sheet, property)
                        : (sheet.CustomProperties.TryGetValue(property, out var v) ? v : string.Empty);
                    if (old != newValue)
                        changes.Add(new DryRunChange { SheetLabel = SheetLabel(sheet), Property = property, OldValue = old, NewValue = newValue });
                }
            }
        }

        return changes;
    }

    public IReadOnlyList<string> GetAvailableProperties()
        => _document.Info.CustomPropertyDefinitions.Select(d => d.Name).ToList();

    public ImportRun Apply(IEnumerable<MappedRow> rows, ImportContext context)
    {
        var run = new ImportRun { ExecutedAt = DateTime.UtcNow, WasDryRun = context.DryRun };
        var rowList = rows.ToList();
        run.TotalRows = rowList.Count;

        // Pass 1 — write set-level values from ALL rows (skip empty so sheet rows don't blank the set)
        if (!context.DryRun)
        {
            foreach (var row in rowList.Where(r => !r.HasErrors))
            {
                foreach (var (property, value) in row.TargetValues.Where(kv => !string.IsNullOrEmpty(kv.Value) && IsSetLevelProperty(kv.Key)))
                {
                    if (StdSetFields.TryGetValue(property, out var xmlProp))
                    {
                        if (_document.Info.Element != null)
                            XmlUtil.SetPropValue(_document.Info.Element, xmlProp, value);
                        ApplyStdSetValue(xmlProp, value);
                    }
                    else if (IsFlags1(property))
                    {
                        if (_document.Info.Element != null)
                            XmlUtil.SetCustomProperty(_document.Info.Element, property, value, flags: 1);
                        _document.Info.CustomProperties[property] = value;
                    }
                }
            }
        }

        // Pass 2 — zip/broadcast sheet-level rows to sheets
        // Only rows that have at least one non-empty sheet-level value qualify (this excludes the SheetSet header row)
        var sheetRows = rowList.Where(r => !r.HasErrors && HasNonEmptySheetLevel(r)).ToList();

        foreach (var (row, sheets) in BuildAssignments(sheetRows, _targetSheets))
        {
            run.Results.Add(row);
            if (row.HasErrors) { run.FailedRows++; continue; }

            if (!context.DryRun)
            {
                foreach (var (property, value) in row.TargetValues.Where(kv => IsSheetLevelProperty(kv.Key)))
                {
                    if (StdSheetFields.TryGetValue(property, out var setter))
                    {
                        foreach (var sheet in sheets)
                        {
                            if (sheet.Element == null) { row.Warnings.Add($"Sheet '{sheet.Number}' heeft geen XElement."); continue; }
                            XmlUtil.SetPropValue(sheet.Element, StdSheetXmlProp(property), value);
                            setter(sheet, value);
                        }
                    }
                    else
                    {
                        foreach (var sheet in sheets)
                        {
                            if (sheet.Element == null) { row.Warnings.Add($"Sheet '{sheet.Number}' heeft geen XElement."); continue; }
                            XmlUtil.SetCustomProperty(sheet.Element, property, value);
                            sheet.CustomProperties[property] = value;
                        }
                    }
                }
            }

            run.SucceededRows++;
        }

        // Set-only rows count as succeeded too
        run.SucceededRows += rowList.Count(r => !r.HasErrors && !HasNonEmptySheetLevel(r));
        run.FailedRows    += rowList.Count(r =>  r.HasErrors && !HasNonEmptySheetLevel(r));

        return run;
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private bool IsSetLevelProperty(string property)
        => StdSetFields.ContainsKey(property) || IsFlags1(property);

    private bool IsSheetLevelProperty(string property) => !IsSetLevelProperty(property);

    private bool HasNonEmptySheetLevel(MappedRow row)
        => row.TargetValues.Any(kv => !string.IsNullOrEmpty(kv.Value) && IsSheetLevelProperty(kv.Key));

    private bool IsFlags1(string property)
        => _document.Info.CustomPropertyDefinitions
            .Any(d => d.Name.Equals(property, StringComparison.OrdinalIgnoreCase) && d.Flags == 1);

    private string GetStdSetValue(string property) => property.ToLowerInvariant() switch
    {
        "sheetset.naam"          => _document.Info.Name             ?? string.Empty,
        "sheetset.omschrijving"  => _document.Info.Description      ?? string.Empty,
        "sheetset.projectnaam"   => _document.Info.ProjectName      ?? string.Empty,
        "sheetset.projectnummer" => _document.Info.ProjectNumber    ?? string.Empty,
        "sheetset.fase"          => _document.Info.ProjectPhase     ?? string.Empty,
        "sheetset.mijlpaal"      => _document.Info.ProjectMilestone ?? string.Empty,
        _ => string.Empty
    };

    private void ApplyStdSetValue(string xmlProp, string value)
    {
        switch (xmlProp)
        {
            case "Name":             _document.Info.Name             = value; break;
            case "Desc":             _document.Info.Description      = value; break;
            case "ProjectName":      _document.Info.ProjectName      = value; break;
            case "ProjectNumber":    _document.Info.ProjectNumber    = value; break;
            case "ProjectPhase":     _document.Info.ProjectPhase     = value; break;
            case "ProjectMilestone": _document.Info.ProjectMilestone = value; break;
        }
    }

    private static string GetStdSheetValue(SheetInfo sheet, string property) => property.ToLowerInvariant() switch
    {
        "nummer"       => sheet.Number         ?? string.Empty,
        "titel"        => sheet.Title          ?? string.Empty,
        "omschrijving" => sheet.Description    ?? string.Empty,
        "revisienummer" => sheet.RevisionNumber ?? string.Empty,
        "revisiedatum"  => sheet.RevisionDate   ?? string.Empty,
        "doel"         => sheet.IssuePurpose   ?? string.Empty,
        "categorie"    => sheet.Category       ?? string.Empty,
        _ => string.Empty
    };

    private static string StdSheetXmlProp(string property) => property.ToLowerInvariant() switch
    {
        "nummer"       => "Number",
        "titel"        => "Title",
        "omschrijving" => "Desc",
        "revisienummer" => "RevisionNumber",
        "revisiedatum"  => "RevisionDate",
        "doel"         => "IssuePurpose",
        "categorie"    => "Category",
        _ => property
    };

    private static string SheetLabel(SheetInfo sheet)
    {
        var label = string.Join(" – ", new[] { sheet.Number, sheet.Title }.Where(s => !string.IsNullOrEmpty(s)));
        return string.IsNullOrEmpty(label) ? "(onbekend)" : label;
    }

    /// <summary>1 row → broadcast to all sheets; N rows → zip row[i] to sheet[i].</summary>
    private static List<(MappedRow Row, List<SheetInfo> Sheets)> BuildAssignments(
        List<MappedRow> rows, List<SheetInfo> sheets)
    {
        if (rows.Count == 0) return [];
        if (rows.Count == 1) return [(rows[0], sheets)];
        return rows.Zip(sheets, (row, sheet) => (row, new List<SheetInfo> { sheet })).ToList();
    }
}
