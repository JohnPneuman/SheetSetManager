using SheetSet.Core.Import.Mapping.Rules;
using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Mapping;

public static class TransformationPipeline
{
    public static string Apply(string value, IEnumerable<TransformationRule> rules,
        IReadOnlyDictionary<string, string> rowContext)
    {
        foreach (var rule in rules)
            value = ApplyRule(value ?? string.Empty, rule, rowContext);
        return value ?? string.Empty;
    }

    private static string ApplyRule(string value, TransformationRule rule,
        IReadOnlyDictionary<string, string> rowContext)
        => rule.Type switch
        {
            TransformationType.Trim             => value.Trim(),
            TransformationType.Uppercase        => value.ToUpperInvariant(),
            TransformationType.Lowercase        => value.ToLowerInvariant(),
            TransformationType.TitleCase        => System.Globalization.CultureInfo.CurrentCulture
                                                       .TextInfo.ToTitleCase(value.ToLowerInvariant()),
            TransformationType.RemoveWhitespace => RemoveWhitespaceRule.Apply(value),
            TransformationType.Replace          => ReplaceRule.Apply(value, rule.Parameters),
            TransformationType.Format           => FormatRule.Apply(value, rule.Parameters),
            TransformationType.CombineFields    => CombineFieldsRule.Apply(rule.Parameters, rowContext),
            TransformationType.DefaultIfEmpty   => DefaultIfEmptyRule.Apply(value, rule.Parameters),
            TransformationType.Truncate         => TruncateRule.Apply(value, rule.Parameters),
            TransformationType.PadLeft          => PadLeftRule.Apply(value, rule.Parameters),
            TransformationType.PadRight         => PadRightRule.Apply(value, rule.Parameters),
            TransformationType.RegexExtract     => RegexExtractRule.Apply(value, rule.Parameters),
            TransformationType.SplitPart        => SplitPartRule.Apply(value, rule.Parameters),
            _                                   => value
        };
}
