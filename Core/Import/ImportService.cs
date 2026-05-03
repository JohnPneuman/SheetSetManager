using SheetSet.Core.Import.Abstractions;
using SheetSet.Core.Import.Models;
using SheetSet.Core.Import.Parsing;

namespace SheetSet.Core.Import;

/// <summary>
/// Orchestrates the full import pipeline:
/// IImportSource → NormalizedRow[] → IMappingEngine → ITargetAdapter
/// </summary>
public class ImportService
{
    private readonly IImportSource _source;
    private readonly IMappingEngine _engine;
    private readonly ITargetAdapter _adapter;

    public ImportService(IImportSource source, IMappingEngine engine, ITargetAdapter adapter)
    {
        _source  = source;
        _engine  = engine;
        _adapter = adapter;
    }

    public ImportPreview Preview(string filePath, SourceHints? hints = null, int maxRows = 8)
        => _source.GetPreview(filePath, hints, maxRows);

    public ImportRun Run(string filePath, ImportProfile profile, ImportContext? context = null)
    {
        context ??= new ImportContext();

        var rows    = _source.ReadRows(filePath, profile.SourceHints);
        var mapped  = _engine.Apply(rows, profile);
        return _adapter.Apply(mapped, context);
    }

    /// <summary>Dry-run: volledig pipeline-pad zonder wegschrijven. Bruikbaar voor preview-validatie.</summary>
    public ImportRun DryRun(string filePath, ImportProfile profile)
        => Run(filePath, profile, new ImportContext { DryRun = true });

    public IReadOnlyList<string> GetAvailableTargetProperties()
        => _adapter.GetAvailableProperties();
}
