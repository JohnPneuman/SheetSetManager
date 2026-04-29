using System;
using System.Collections.Generic;
using System.IO;
using BoekSolutions.SheetSetEditor.Helpers;
using Newtonsoft.Json;

namespace BoekSolutions.SheetSetEditor
{
    public static class RecentFilesHelper
    {
        private static string GetPath()
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string versionPart = Aliases.AcadVer?.Split('.')?[0] ?? "24"; // bijv. "24" voor AutoCAD 2024
            string profile = Aliases.Profile ?? "Unnamed Profile";

            string folder = Path.Combine(
                appDataPath,
                "Autodesk",
                $"AutoCAD {versionPart}",
                $"R{versionPart}",
                "enu",
                "Support",
                "Profiles",
                profile
            );

            Directory.CreateDirectory(folder);
            return Path.Combine(folder, "ssm_recent.json");
        }

        public static List<string> Load()
        {
            return TryHelper.Run(() =>
            {
                var path = GetPath();
                if (!File.Exists(path))
                    return new List<string>();

                var json = File.ReadAllText(path);
                return JsonConvert.DeserializeObject<List<string>>(json) ?? new List<string>();
            }, "RecentFilesHelper.Load", new List<string>());
        }

        public static void Save(List<string> items)
        {
            TryHelper.Run(() =>
            {
                var json = JsonConvert.SerializeObject(items, Formatting.Indented);
                File.WriteAllText(GetPath(), json);
            }, "RecentFilesHelper.Save");
        }

        public static void AddToRecent(string filePath)
        {
            TryHelper.Run(() =>
            {
                var recent = Load();

                // Voorkom dubbele entries
                recent.Remove(filePath);
                recent.Insert(0, filePath);

                // Beperk de lijst tot 15 items
                while (recent.Count > 15)
                    recent.RemoveAt(recent.Count - 1);

                Save(recent);
            }, "RecentFilesHelper.AddToRecent");
        }
        public static List<string> GetRecentSheetSets()
        {
            return Load();
        }


    }
}
