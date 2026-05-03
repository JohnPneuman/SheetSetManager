namespace SheetSet.Core.Import.Models;

public class SourceHints
{
    /// <summary>null = auto-detect.</summary>
    public char? Delimiter { get; set; }

    /// <summary>null = auto-detect (tries UTF-8 BOM, then Windows-1252).</summary>
    public string? Encoding { get; set; }

    public bool HasHeader { get; set; } = true;
    public int SkipRows { get; set; }

    /// <summary>null = auto-detect; set explicitly to override detection.</summary>
    public FieldLayout? Layout { get; set; }
}
