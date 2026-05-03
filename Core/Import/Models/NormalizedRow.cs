namespace SheetSet.Core.Import.Models;

/// <summary>
/// Source-agnostic representation of one record.
/// In Horizontal mode: one NormalizedRow per data line.
/// In Vertical mode: one NormalizedRow per file; Fields keyed by field index ("1", "2", …).
/// </summary>
public class NormalizedRow
{
    public int RowIndex { get; set; }
    public string SourceFile { get; set; } = string.Empty;
    public IReadOnlyDictionary<string, string> Fields { get; set; } = new Dictionary<string, string>();
}
