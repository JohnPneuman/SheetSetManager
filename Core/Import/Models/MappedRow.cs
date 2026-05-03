namespace SheetSet.Core.Import.Models;

/// <summary>Output of the mapping engine: resolved target property values for one record.</summary>
public class MappedRow
{
    public int SourceRowIndex { get; set; }

    /// <summary>property name → transformed value, ready for the target adapter.</summary>
    public Dictionary<string, string> TargetValues { get; set; } = new();

    public List<string> Warnings { get; set; } = new();
    public List<string> Errors { get; set; } = new();
    public bool HasErrors => Errors.Count > 0;
}
