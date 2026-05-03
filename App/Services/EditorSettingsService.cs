using System.IO;
using System.Text.Json;

namespace SheetSetEditor.Services;

public static class EditorSettingsService
{
    private static readonly string _filePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "BoekSolutions", "SheetSetEditor", "settings.json");

    private sealed class Settings
    {
        public string? LastTemplatePath { get; set; }
        public string? LastSaveFolder   { get; set; }
    }

    public static string? GetLastTemplatePath() => Load().LastTemplatePath;
    public static string? GetLastSaveFolder()   => Load().LastSaveFolder;

    public static void SetLastTemplatePath(string? path)
    {
        var s = Load(); s.LastTemplatePath = path; Save(s);
    }

    public static void SetLastSaveFolder(string? path)
    {
        var s = Load(); s.LastSaveFolder = path; Save(s);
    }

    private static Settings Load()
    {
        try
        {
            if (!File.Exists(_filePath)) return new Settings();
            return JsonSerializer.Deserialize<Settings>(File.ReadAllText(_filePath)) ?? new Settings();
        }
        catch { return new Settings(); }
    }

    private static void Save(Settings s)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_filePath)!);
            File.WriteAllText(_filePath, JsonSerializer.Serialize(s,
                new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }
}
