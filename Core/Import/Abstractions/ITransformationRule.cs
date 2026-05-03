using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Abstractions;

public interface ITransformationRule
{
    TransformationType Type { get; }

    /// <param name="rowContext">All fields of the current row — needed for CombineFields.</param>
    string Apply(string input, IReadOnlyDictionary<string, string> rowContext);
}
