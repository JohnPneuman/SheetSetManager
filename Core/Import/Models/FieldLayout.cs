namespace SheetSet.Core.Import.Models;

public enum FieldLayout
{
    /// <summary>One record per line, columns as fields (traditional CSV/TSV).</summary>
    Horizontal,

    /// <summary>One field per line: col1 = field index, col2 = value (acad_oh/bh format).</summary>
    Vertical
}
