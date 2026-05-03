namespace SheetSet.Core.Import.Mapping.Rules;

/// <summary>Parameters: fields (comma-separated column names or indices), separator (default " ").</summary>
public static class CombineFieldsRule
{
    public static string Apply(IReadOnlyDictionary<string, string> parameters,
        IReadOnlyDictionary<string, string> rowContext)
    {
        var fieldsCsv = parameters.GetValueOrDefault("fields", string.Empty);
        var separator = parameters.GetValueOrDefault("separator", " ");

        var parts = fieldsCsv
            .Split(',')
            .Select(f => f.Trim())
            .Select(f => rowContext.TryGetValue(f, out var v) ? v.Trim() : string.Empty)
            .Where(v => !string.IsNullOrEmpty(v));

        return string.Join(separator, parts);
    }
}
