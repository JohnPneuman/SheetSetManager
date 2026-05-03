namespace SheetSet.Core.Import.Models;

public enum TransformationType
{
    Trim,
    Uppercase,
    Lowercase,
    TitleCase,
    Replace,         // parameters: find, replace, regex (bool)
    Format,          // parameters: pattern e.g. "PRJ-{0}"
    CombineFields,   // parameters: fields (comma-separated), separator
    DefaultIfEmpty,  // parameters: default
    Truncate,        // parameters: maxLength
    PadLeft,         // parameters: length, char
    PadRight,        // parameters: length, char
    RegexExtract,    // parameters: pattern, group (int)
    RemoveWhitespace,
    SplitPart        // parameters: separator, index (0-based; -1 = laatste deel)
}

public class TransformationRule
{
    public TransformationType Type { get; set; }
    public Dictionary<string, string> Parameters { get; set; } = new();
}
