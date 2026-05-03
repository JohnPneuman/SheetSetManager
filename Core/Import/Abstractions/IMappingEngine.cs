using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Abstractions;

public interface IMappingEngine
{
    IReadOnlyList<MappedRow> Apply(IEnumerable<NormalizedRow> rows, ImportProfile profile);
}
