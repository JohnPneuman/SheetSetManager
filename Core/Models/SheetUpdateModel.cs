namespace SheetSet.Core.Models;

public sealed class SheetUpdateModel
{
    public required SheetInfo Sheet { get; init; }

    public string? NewNumber { get; set; }
    public string? Title { get; set; }
    public string? Description { get; set; }
    public string? Category { get; set; }
    public string? RevisionNumber { get; set; }
    public string? RevisionDate { get; set; }
    public string? IssuePurpose { get; set; }
    public bool? DoNotPlot { get; set; }

    // Only non-null entries are applied; null dict = no custom property changes
    public Dictionary<string, string>? CustomProperties { get; set; }
}
