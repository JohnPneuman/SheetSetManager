using System.Xml.Linq;
using SheetSet.Core.Models;

namespace SheetSet.Core.Import.History;

public class ImportSnapshot
{
    public DateTime Timestamp         { get; init; } = DateTime.UtcNow;
    public string   SourceDescription { get; init; } = string.Empty;
    public List<SheetStateSnapshot> Sheets { get; init; } = [];
}

public class SheetStateSnapshot
{
    public SheetInfo  Sheet                    { get; init; } = null!;
    public XElement   OriginalElement          { get; init; } = null!;
    public Dictionary<string, string> OriginalCustomProperties { get; init; } = new();
}
