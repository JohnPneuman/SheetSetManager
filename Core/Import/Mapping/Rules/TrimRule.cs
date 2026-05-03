using SheetSet.Core.Import.Abstractions;
using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Mapping.Rules;

public class TrimRule : ITransformationRule
{
    public TransformationType Type => TransformationType.Trim;
    public string Apply(string input, IReadOnlyDictionary<string, string> _) => input?.Trim() ?? string.Empty;
}
