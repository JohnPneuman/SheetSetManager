using System.Text.RegularExpressions;
using SheetSet.Core.Import.Abstractions;
using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Mapping;

public class MappingEngine : IMappingEngine
{
    public IReadOnlyList<MappedRow> Apply(IEnumerable<NormalizedRow> rows, ImportProfile profile)
    {
        var results = new List<MappedRow>();

        foreach (var row in rows)
        {
            var mapped = new MappedRow { SourceRowIndex = row.RowIndex };

            foreach (var mapping in profile.FieldMappings)
            {
                if (!row.Fields.TryGetValue(mapping.SourceColumn, out var rawValue))
                    rawValue = null;

                // Apply fallback before transformations if source is missing/empty
                if (string.IsNullOrWhiteSpace(rawValue))
                    rawValue = mapping.FallbackValue ?? string.Empty;

                var transformed = TransformationPipeline.Apply(rawValue, mapping.Transformations, row.Fields);

                if (string.IsNullOrWhiteSpace(transformed))
                {
                    if (mapping.IsRequired)
                        mapped.Errors.Add(
                            $"Verplicht veld '{mapping.TargetProperty}' heeft geen waarde (bronkolom: '{mapping.SourceColumn}').");
                    continue;
                }

                if (mapping.MaxLength.HasValue && transformed.Length > mapping.MaxLength.Value)
                    mapped.Warnings.Add(
                        $"'{mapping.TargetProperty}': waarde te lang ({transformed.Length} tekens, max {mapping.MaxLength.Value}).");

                if (!string.IsNullOrEmpty(mapping.ValidationPattern) &&
                    !Regex.IsMatch(transformed, mapping.ValidationPattern))
                    mapped.Warnings.Add(
                        $"'{mapping.TargetProperty}': waarde \"{transformed}\" voldoet niet aan patroon '{mapping.ValidationPattern}'.");

                mapped.TargetValues[mapping.TargetProperty] = transformed;
            }

            results.Add(mapped);
        }

        return results;
    }
}
