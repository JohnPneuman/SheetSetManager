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
        public string? LastTemplatePath  { get; set; }
        public string? LastSaveFolder    { get; set; }
        public string  Language          { get; set; } = "Dutch";
        public bool    DarkMode          { get; set; } = false;
        public bool    AutoBackup        { get; set; } = true;
        public int     MaxRecentFiles    { get; set; } = 20;
    }

    public static string?      GetLastTemplatePath() => Load().LastTemplatePath;
    public static string?      GetLastSaveFolder()   => Load().LastSaveFolder;
    public static AppLanguage  GetLanguage()          => Enum.TryParse<AppLanguage>(Load().Language, out var l) ? l : AppLanguage.Dutch;
    public static bool         GetDarkMode()          => Load().DarkMode;
    public static bool         GetAutoBackup()        => Load().AutoBackup;
    public static int          GetMaxRecentFiles()    => Load().MaxRecentFiles;

    public static void SetLastTemplatePath(string? path)    { var s = Load(); s.LastTemplatePath = path; Save(s); }
    public static void SetLastSaveFolder(string? path)      { var s = Load(); s.LastSaveFolder   = path; Save(s); }
    public static void SetLanguage(AppLanguage lang)        { var s = Load(); s.Language          = lang.ToString(); Save(s); }
    public static void SetDarkMode(bool dark)               { var s = Load(); s.DarkMode          = dark; Save(s); }
    public static void SetAutoBackup(bool value)            { var s = Load(); s.AutoBackup        = value; Save(s); }
    public static void SetMaxRecentFiles(int count)         { var s = Load(); s.MaxRecentFiles    = count; Save(s); }

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
