using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Abstractions;

public interface ITargetAdapter
{
    /// <summary>Returns all property names the adapter can write to (for the mapping UI).</summary>
    IReadOnlyList<string> GetAvailableProperties();

    ImportRun Apply(IEnumerable<MappedRow> rows, ImportContext context);
}
