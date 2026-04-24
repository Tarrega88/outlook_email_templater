using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using Envoy.Models;

namespace Envoy.Services;

public static class ConfigService
{
    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };

    public static string DefaultPath
    {
        get => Path.Combine(AppDataPaths.Dir, "config.json");
    }

    public static AppConfig Load(string path)
    {
        if (!File.Exists(path)) return new AppConfig();
        try
        {
            var json = File.ReadAllText(path);

            // Backward compatibility: older configs stored the list under "Groups".
            // Rename the key to "Emails" before deserializing so the collection loads.
            if (JsonNode.Parse(json) is JsonObject root)
            {
                if (!root.ContainsKey("Emails") && root["Groups"] is JsonNode legacy)
                {
                    root["Emails"] = legacy.DeepClone();
                    root.Remove("Groups");
                    json = root.ToJsonString();
                }
            }

            return JsonSerializer.Deserialize<AppConfig>(json, Options) ?? new AppConfig();
        }
        catch
        {
            return new AppConfig();
        }
    }

    public static void Save(AppConfig config, string path)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        var json = JsonSerializer.Serialize(config, Options);
        File.WriteAllText(path, json);
    }
}
