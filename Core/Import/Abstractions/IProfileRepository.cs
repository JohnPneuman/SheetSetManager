using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Abstractions;

public interface IProfileRepository
{
    void Save(ImportProfile profile);

    /// <summary>Returns null when not found.</summary>
    ImportProfile Load(Guid id);

    IReadOnlyList<ImportProfile> List();

    void Delete(Guid id);
}
