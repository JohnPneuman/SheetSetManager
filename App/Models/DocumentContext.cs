using System.IO;
using System.Windows.Controls;
using SheetSet.Core.Models;

namespace SheetSetEditor.Models;

/// <summary>One open .dst document with its per-document selection state.</summary>
public sealed class DocumentContext
{
    public SheetSetDocument Document { get; }
    public bool IsDirty { get; set; }

    // Per-document selection state (preserved when switching active doc)
    public HashSet<SheetNode> SelectedSheets { get; } = [];
    public SubsetNode? SelectedSubset { get; set; }
    public SheetSetDocument? SelectedRoot { get; set; } // the root node itself
    public TreeViewItem? LastClickedItem { get; set; }

    public DocumentContext(SheetSetDocument doc) => Document = doc;

    public string DisplayName =>
        Document.Info.Name ??
        Path.GetFileNameWithoutExtension(Document.SourceDstPath ?? "Onbekend");

    public string FilePath => Document.SourceDstPath ?? string.Empty;
}
