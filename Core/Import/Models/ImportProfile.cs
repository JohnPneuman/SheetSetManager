namespace SheetSet.Core.Import.Models;

public class ImportProfile
{
    public Guid Id { get; set; } = Guid.NewGuid();

    /// <summary>Incremented when the schema changes; used by ProfileLoader for migration.</summary>
    public int Version { get; set; } = 1;

    public string Name { get; set; } = string.Empty;
    public string? Description { get; set; }
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;

    public SourceHints SourceHints { get; set; } = new();
    public List<FieldMapping> FieldMappings { get; set; } = new();

    /// <summary>File names (without path) that this profile has been successfully used with.</summary>
    public List<string> AssociatedFileNames { get; set; } = new();
}
