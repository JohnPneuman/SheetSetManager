namespace SheetSet.Core.Import.Mapping.Rules;

/// <summary>Parameters: pattern e.g. "PRJ-{0}" or "{0:D5}".</summary>
public static class FormatRule
{
    public static string Apply(string input, IReadOnlyDictionary<string, string> parameters)
    {
        var pattern = parameters.GetValueOrDefault("pattern");
        if (string.IsNullOrEmpty(pattern)) return input ?? string.Empty;
        try { return string.Format(pattern, input); }
        catch { return input ?? string.Empty; }
    }
}
