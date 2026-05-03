using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SheetSet.Core.Import.Models;

namespace SheetSet.Core.Import.Profiles;

/// <summary>
/// Deserializes an ImportProfile from JSON and applies any pending migrations.
/// </summary>
public static class ProfileLoader
{
    private const int CurrentVersion = 1;

    // Register migrations here in order: (fromVersion, migrate func)
    private static readonly List<(int FromVersion, Func<JObject, JObject> Migrate)> Migrations = [];

    public static ImportProfile FromJson(string json)
    {
        var obj = JObject.Parse(json);
        obj = ApplyMigrations(obj);
        return obj.ToObject<ImportProfile>(JsonSerializer.Create(SerializerSettings()))!;
    }

    public static string ToJson(ImportProfile profile)
    {
        profile.Version = CurrentVersion;
        return JsonConvert.SerializeObject(profile, SerializerSettings());
    }

    private static JObject ApplyMigrations(JObject obj)
    {
        var version = obj["version"]?.Value<int>() ?? 1;
        foreach (var (fromVersion, migrate) in Migrations)
        {
            if (version == fromVersion)
            {
                obj = migrate(obj);
                version++;
                obj["version"] = version;
            }
        }
        return obj;
    }

    private static JsonSerializerSettings SerializerSettings() => new()
    {
        Formatting = Formatting.Indented,
        NullValueHandling = NullValueHandling.Ignore,
        DefaultValueHandling = DefaultValueHandling.Ignore
    };
}
