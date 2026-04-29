using System.Xml.Linq;

namespace SheetSet.Core.Models;

public sealed class SheetSetDocument
{
    public SheetSetInfo Info { get; }
    public List<SubsetNode> RootSubsets { get; }
    public List<SheetNode> RootSheets { get; }

    // Set when loaded via SheetSetParser.ParseForEditing()
    public XDocument? SourceDocument { get; set; }
    public string? SourceDstPath { get; set; }

    public SheetSetDocument(SheetSetInfo info, List<SubsetNode> rootSubsets, List<SheetNode> rootSheets)
    {
        Info = info;
        RootSubsets = rootSubsets;
        RootSheets = rootSheets;
    }

    public IEnumerable<SheetNode> GetAllSheets()
    {
        foreach (var sheet in RootSheets)
            yield return sheet;

        foreach (var subset in RootSubsets)
            foreach (var sheet in subset.GetAllSheets())
                yield return sheet;
    }
}

public sealed class SubsetNode
{
    public SubsetInfo Info { get; }
    public string Name => Info.Name ?? "(zonder naam)";
    public List<SubsetNode> Subsets { get; }
    public List<SheetNode> Sheets { get; }

    public SubsetNode(SubsetInfo info, List<SubsetNode> subsets, List<SheetNode> sheets)
    {
        Info = info;
        Subsets = subsets;
        Sheets = sheets;
    }

    public IEnumerable<SheetNode> GetAllSheets()
    {
        foreach (var sheet in Sheets)
            yield return sheet;

        foreach (var subset in Subsets)
            foreach (var sheet in subset.GetAllSheets())
                yield return sheet;
    }
}

public sealed class SheetNode
{
    public SheetInfo Info { get; }
    public string? Number => Info.Number;
    public string? Title => Info.Title;
    public string? LayoutName => Info.LayoutName;
    public string? BestDwgPath => Info.ResolvedDwgPath;
    public string? BestFolderPath => Info.FolderPath;

    public SheetNode(SheetInfo info) => Info = info;
}

public sealed class SheetSetInfo
{
    public string? Name { get; set; }
    public string? Description { get; set; }
    public string? ProjectName { get; set; }
    public string? ProjectNumber { get; set; }
    public string? ProjectPhase { get; set; }
    public string? ProjectMilestone { get; set; }
    public string? BaseFolder { get; set; }
    public Dictionary<string, string> CustomProperties { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public XElement? Element { get; set; }
}

public sealed class SubsetInfo
{
    public string? Name { get; set; }
    public string? Description { get; set; }
    public FileReferenceInfo? NewSheetLocation { get; set; }
    public Dictionary<string, string> CustomProperties { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public XElement? Element { get; set; }
}

public sealed class SheetInfo
{
    public string? Number { get; set; }
    public string? Title { get; set; }
    public string? Description { get; set; }
    public string? Category { get; set; }
    public string? RevisionNumber { get; set; }
    public string? RevisionDate { get; set; }
    public string? IssuePurpose { get; set; }
    public bool? DoNotPlot { get; set; }
    public string? LayoutName { get; set; }
    public string? DwgFileName { get; set; }
    public string? RelativeDwgFileName { get; set; }
    public string? ResolvedDwgPath { get; set; }
    public string? FolderPath { get; set; }
    public Dictionary<string, string> CustomProperties { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public XElement? Element { get; set; }
}

public sealed class FileReferenceInfo
{
    public string? FileName { get; set; }
    public string? RelativeFileName { get; set; }
    public string? ResolvedPath { get; set; }
    public override string ToString() => ResolvedPath ?? FileName ?? RelativeFileName ?? string.Empty;
}

public sealed class LayoutReferenceInfo
{
    public string? Name { get; set; }
    public string? FileName { get; set; }
    public string? RelativeFileName { get; set; }
    public string? ResolvedPath { get; set; }
}
