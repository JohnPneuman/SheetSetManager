using System.Text.RegularExpressions;

namespace SheetSet.Core.Import.Mapping.Rules;

/// <summary>Parameters: find, replace, regex (bool, default false).</summary>
public static class ReplaceRule
{
    public static string Apply(string input, IReadOnlyDictionary<string, string> parameters)
    {
        if (string.IsNullOrEmpty(input)) return input ?? string.Empty;
        var find    = parameters.GetValueOrDefault("find",    string.Empty);
        var replace = parameters.GetValueOrDefault("replace", string.Empty);
        var useRegex = parameters.GetValueOrDefault("regex", "false")
            .Equals("true", StringComparison.OrdinalIgnoreCase);

        if (string.IsNullOrEmpty(find)) return input;
        return useRegex
            ? Regex.Replace(input, find, replace)
            : input.Replace(find, replace);
    }
}
