namespace SheetSet.Core.Import.Models;

public class ImportPreview
{
    public string FilePath { get; set; } = string.Empty;

    /// <summary>What the parser detected; can differ from hints passed in.</summary>
    public SourceHints DetectedHints { get; set; } = new();

    /// <summary>Column names (Horizontal) or field indices as strings (Vertical).</summary>
    public List<string> Headers { get; set; } = new();

    public List<NormalizedRow> SampleRows { get; set; } = new();
    public List<ValidationIssue> ValidationIssues { get; set; } = new();

    public bool HasErrors => ValidationIssues.Any(i => i.Severity == ValidationSeverity.Error);
}
