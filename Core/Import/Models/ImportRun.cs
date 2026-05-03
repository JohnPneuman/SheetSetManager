namespace SheetSet.Core.Import.Models;

public class ImportRun
{
    public Guid RunId { get; set; } = Guid.NewGuid();
    public Guid ProfileId { get; set; }
    public string SourceFile { get; set; } = string.Empty;
    public DateTime ExecutedAt { get; set; } = DateTime.UtcNow;
    public bool WasDryRun { get; set; }

    public int TotalRows { get; set; }
    public int SucceededRows { get; set; }
    public int SkippedRows { get; set; }
    public int FailedRows { get; set; }

    public List<MappedRow> Results { get; set; } = new();
}
