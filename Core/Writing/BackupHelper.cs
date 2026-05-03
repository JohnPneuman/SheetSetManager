namespace SheetSet.Core.Writing;

public static class BackupHelper
{
    private const int MaxBackups = 5;

    private static string BackupDir
    {
        get
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BoekSolutions", "Backups");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    public static void BackupFile(string dstPath)
    {
        if (string.IsNullOrWhiteSpace(dstPath) || !File.Exists(dstPath)) return;

        try
        {
            var name = Path.GetFileNameWithoutExtension(dstPath);
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var backupPath = Path.Combine(BackupDir, $"{name}.backup_{stamp}.dst");

            File.Copy(dstPath, backupPath, overwrite: false);

            PruneOldBackups(name);
        }
        catch
        {
            // Backup mislukken mag de opslag niet blokkeren
        }
    }

    private static void PruneOldBackups(string name)
    {
        var backups = Directory.GetFiles(BackupDir, $"{name}.backup_*.dst");
        Array.Sort(backups);
        for (int i = 0; i < backups.Length - MaxBackups; i++)
            File.Delete(backups[i]);
    }
}
