using System.ComponentModel;
using System.Runtime.CompilerServices;
using SheetSet.Core.Import.Models;

namespace SheetSetEditor.ImportWizard;

public class MappingRowViewModel : INotifyPropertyChanged
{
    private string _targetProperty = string.Empty;

    /// <summary>Veldnaam of index uit de bronrij ("1", "2", "Klantnaam", ...).</summary>
    public string SourceField { get; init; } = string.Empty;

    /// <summary>Eerste niet-lege voorbeeldwaarde uit de preview.</summary>
    public string SourceSample { get; init; } = string.Empty;

    /// <summary>Beschikbare target-properties (inclusief lege eerste optie).</summary>
    public List<string> AvailableTargets { get; init; } = [];

    public string TargetProperty
    {
        get => _targetProperty;
        set { _targetProperty = value; OnPropertyChanged(); }
    }

    /// <summary>Transformations to apply before writing to target property.</summary>
    public List<TransformationRule> Transformations { get; set; } = [];

    /// <summary>True when this row was created as a derived split from another row.</summary>
    public bool IsDerived { get; init; }

    /// <summary>Display label shown in the source column for derived rows (e.g. "↳ deel 1").</summary>
    public string DerivedLabel { get; init; } = string.Empty;

    // ── Validatie ─────────────────────────────────────────────────────────────

    public bool   IsRequired          { get; set; }
    public string ValidationPattern   { get; set; } = string.Empty;
    public int?   ValidationMaxLength { get; set; }

    /// <summary>Human-readable summary of active validation rules for the tooltip/indicator.</summary>
    public string ValidationSummary
    {
        get
        {
            var parts = new List<string>();
            if (IsRequired)                               parts.Add("Verplicht");
            if (ValidationMaxLength.HasValue)             parts.Add($"Max {ValidationMaxLength} tekens");
            if (!string.IsNullOrEmpty(ValidationPattern)) parts.Add($"Patroon: {ValidationPattern}");
            return parts.Count > 0 ? string.Join(", ", parts) : string.Empty;
        }
    }

    public bool HasValidation => IsRequired || ValidationMaxLength.HasValue || !string.IsNullOrEmpty(ValidationPattern);

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
