using System.Text.RegularExpressions;

namespace SheetSet.Core.Import.Mapping.Rules;

public static class DefaultIfEmptyRule
{
    public static string Apply(string input, IReadOnlyDictionary<string, string> parameters)
        => string.IsNullOrWhiteSpace(input)
            ? parameters.GetValueOrDefault("default", string.Empty)
            : input;
}

public static class TruncateRule
{
    public static string Apply(string input, IReadOnlyDictionary<string, string> parameters)
    {
        if (string.IsNullOrEmpty(input)) return string.Empty;
        if (!int.TryParse(parameters.GetValueOrDefault("maxLength", "255"), out var max)) return input;
        return input.Length <= max ? input : input[..max];
    }
}

public static class PadLeftRule
{
    public static string Apply(string input, IReadOnlyDictionary<string, string> parameters)
    {
        if (!int.TryParse(parameters.GetValueOrDefault("length", "0"), out var len))
            return input ?? string.Empty;
        var padChar = (parameters.GetValueOrDefault("char", "0") + "0")[0];
        return (input ?? string.Empty).PadLeft(len, padChar);
    }
}

public static class PadRightRule
{
    public static string Apply(string input, IReadOnlyDictionary<string, string> parameters)
    {
        if (!int.TryParse(parameters.GetValueOrDefault("length", "0"), out var len))
            return input ?? string.Empty;
        var padChar = (parameters.GetValueOrDefault("char", " ") + " ")[0];
        return (input ?? string.Empty).PadRight(len, padChar);
    }
}

public static class RemoveWhitespaceRule
{
    public static string Apply(string input)
        => string.IsNullOrEmpty(input) ? string.Empty : Regex.Replace(input, @"\s+", string.Empty);
}

/// <summary>Parameters: separator (default " / "), index (0-based; -1 = last part).</summary>
public static class SplitPartRule
{
    public static string Apply(string input, IReadOnlyDictionary<string, string> parameters)
    {
        if (string.IsNullOrEmpty(input)) return string.Empty;
        var separator = parameters.GetValueOrDefault("separator", " / ");
        if (!int.TryParse(parameters.GetValueOrDefault("index", "0"), out var index))
            return string.Empty;

        var parts = input.Split(new[] { separator }, StringSplitOptions.None);
        if (index < 0) index = parts.Length + index;
        return index >= 0 && index < parts.Length ? parts[index].Trim() : string.Empty;
    }
}

public static class RegexExtractRule
{
    public static string Apply(string input, IReadOnlyDictionary<string, string> parameters)
    {
        if (string.IsNullOrEmpty(input)) return string.Empty;
        var pattern = parameters.GetValueOrDefault("pattern", "(.*)");
        int.TryParse(parameters.GetValueOrDefault("group", "1"), out var group);
        var match = Regex.Match(input, pattern);
        return match.Success && match.Groups.Count > group ? match.Groups[group].Value : string.Empty;
    }
}
