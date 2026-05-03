namespace SheetSet.Core.Import.Models;

public class ImportContext
{
    /// <summary>When true the adapter validates and collects results without writing anything.</summary>
    public bool DryRun { get; set; }

    /// <summary>Skip records whose target values are already equal to the current values.</summary>
    public bool SkipUnchanged { get; set; }
}
