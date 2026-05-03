namespace SheetSet.Core.Import.Models;

public class DryRunChange
{
    public string SheetLabel { get; set; } = string.Empty;
    public string Property   { get; set; } = string.Empty;
    public string OldValue   { get; set; } = string.Empty;
    public string NewValue   { get; set; } = string.Empty;
    public bool   WillChange => !string.Equals(OldValue, NewValue, StringComparison.Ordinal);
}
