namespace SheetSet.Core.Import.Models;

public class FieldMapping
{
    /// <summary>
    /// Column header name (Horizontal) or field index as string e.g. "3" (Vertical).
    /// </summary>
    public string SourceColumn { get; set; } = string.Empty;

    /// <summary>Name of the target Sheet Set custom property.</summary>
    public string TargetProperty { get; set; } = string.Empty;

    public bool IsRequired { get; set; }

    /// <summary>Used when source is empty and no DefaultIfEmpty transformation is defined.</summary>
    public string? FallbackValue { get; set; }

    public List<TransformationRule> Transformations { get; set; } = new();

    /// <summary>Regex pattern the transformed value must match; empty = no check.</summary>
    public string? ValidationPattern { get; set; }

    /// <summary>Maximum allowed length; null = no limit.</summary>
    public int? MaxLength { get; set; }
}
