using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Abstractions;

public interface IImportSource
{
    bool CanHandle(string filePath);

    ImportPreview GetPreview(string filePath, SourceHints? hints = null, int maxRows = 8);

    IEnumerable<NormalizedRow> ReadRows(string filePath, SourceHints? hints = null);
}
