using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Parsing;

public static class DelimiterDetector
{
    private static readonly char[] Candidates = ['\t', ';', ',', '|'];

    public static char Detect(IReadOnlyList<string> sampleLines)
    {
        if (sampleLines.Count == 0) return ',';

        var best = Candidates
            .Select(d => (delimiter: d, score: Score(sampleLines, d)))
            .OrderByDescending(x => x.score)
            .First();

        return best.score > 0 ? best.delimiter : ',';
    }

    /// <summary>
    /// Detects Vertical (one value per line) vs Horizontal (multiple columns per line).
    /// Vertical: fewer than half the non-empty lines contain the delimiter.
    /// </summary>
    public static FieldLayout DetectLayout(IReadOnlyList<string> sampleLines, char delimiter)
    {
        var nonEmpty = sampleLines.Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
        if (nonEmpty.Count == 0) return FieldLayout.Horizontal;

        var withDelimiter = nonEmpty.Count(l => l.Contains(delimiter));
        return withDelimiter < nonEmpty.Count / 2.0
            ? FieldLayout.Vertical
            : FieldLayout.Horizontal;
    }

    private static double Score(IReadOnlyList<string> lines, char delimiter)
    {
        var counts = lines
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Select(l => CountUnquoted(l, delimiter))
            .ToList();

        if (counts.Count == 0 || counts.All(c => c == 0)) return 0;

        var mean = counts.Average();
        if (mean < 1) return 0;

        var variance = counts.Select(c => Math.Pow(c - mean, 2)).Average();
        return mean / (1 + Math.Sqrt(variance));
    }

    private static int CountUnquoted(string line, char delimiter)
    {
        var count = 0;
        var inQuotes = false;
        foreach (var ch in line)
        {
            if (ch == '"') inQuotes = !inQuotes;
            else if (ch == delimiter && !inQuotes) count++;
        }
        return count;
    }
}
