namespace SheetSet.Core.Import.Models;

public enum ValidationSeverity { Info, Warning, Error }

public class ValidationIssue
{
    public ValidationSeverity Severity { get; set; }
    public string Message { get; set; } = string.Empty;
    public int? RowIndex { get; set; }
    public string? Column { get; set; }
}
