using SheetSet.Core.Import.Abstractions;
using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Profiles;

/// <summary>
/// Persists ImportProfiles as individual JSON files under %APPDATA%\BoekSolutions\ImportProfiles\.
/// </summary>
public class JsonProfileRepository : IProfileRepository
{
    private readonly string _directory;

    public JsonProfileRepository()
        : this(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "BoekSolutions", "ImportProfiles"))
    { }

    public JsonProfileRepository(string directory)
    {
        _directory = directory;
        Directory.CreateDirectory(_directory);
    }

    public void Save(ImportProfile profile)
    {
        profile.UpdatedAt = DateTime.UtcNow;
        var path = FilePath(profile.Id);
        File.WriteAllText(path, ProfileLoader.ToJson(profile));
    }

    public ImportProfile Load(Guid id)
    {
        var path = FilePath(id);
        if (!File.Exists(path)) return null;
        return ProfileLoader.FromJson(File.ReadAllText(path));
    }

    public IReadOnlyList<ImportProfile> List()
    {
        return Directory
            .EnumerateFiles(_directory, "*.json")
            .Select(f =>
            {
                try { return ProfileLoader.FromJson(File.ReadAllText(f)); }
                catch { return null; }
            })
            .OfType<ImportProfile>()
            .OrderByDescending(p => p.UpdatedAt)
            .ToList();
    }

    public void Delete(Guid id)
    {
        var path = FilePath(id);
        if (File.Exists(path))
            File.Delete(path);
    }

    /// <summary>Returns the most-recently-updated profile whose AssociatedFileNames contains fileName (case-insensitive).</summary>
    public ImportProfile? FindByFileName(string fileName)
        => List().FirstOrDefault(p =>
            p.AssociatedFileNames.Any(f => string.Equals(f, fileName, StringComparison.OrdinalIgnoreCase)));

    /// <summary>Returns the existing profile with the given name, or null.</summary>
    public ImportProfile? FindByName(string name)
        => List().FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));

    private string FilePath(Guid id) => Path.Combine(_directory, $"{id:N}.json");
}
