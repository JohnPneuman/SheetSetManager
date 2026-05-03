using System.Globalization;
using SheetSet.Core.Import.Abstractions;
using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Mapping.Rules;

public class UppercaseRule : ITransformationRule
{
    public TransformationType Type => TransformationType.Uppercase;
    public string Apply(string input, IReadOnlyDictionary<string, string> _) => input?.ToUpperInvariant() ?? string.Empty;
}

public class LowercaseRule : ITransformationRule
{
    public TransformationType Type => TransformationType.Lowercase;
    public string Apply(string input, IReadOnlyDictionary<string, string> _) => input?.ToLowerInvariant() ?? string.Empty;
}

public class TitleCaseRule : ITransformationRule
{
    public TransformationType Type => TransformationType.TitleCase;
    public string Apply(string input, IReadOnlyDictionary<string, string> _)
        => string.IsNullOrEmpty(input) ? string.Empty
            : CultureInfo.CurrentCulture.TextInfo.ToTitleCase(input.ToLowerInvariant());
}
