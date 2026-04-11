using System.Text.Json;
using System.Text.Json.Nodes;

namespace XRai.Studio;

/// <summary>
/// Persisted per-user dashboard preferences. Stored as a tiny JSON file at
/// %LOCALAPPDATA%\XRai\studio\preferences.json so a second Studio launch
/// remembers the user's choices without asking again.
///
/// Schema (versioned via "schemaVersion"):
///   {
///     "schemaVersion": 1,
///     "followMode": true,
///     "preferredIde": "VSCode",     // "VSCode" | "VisualStudio" | "Rider" | null
///     "onboarded": true,             // true once the user has seen the startup overlay
///     "theme": "dark"
///   }
/// </summary>
public sealed class StudioPreferences
{
    public int SchemaVersion { get; set; } = 1;
    public bool FollowMode { get; set; } = true;
    public string? PreferredIde { get; set; }
    public bool Onboarded { get; set; }
    public string Theme { get; set; } = "dark";

    private static string PreferencesPath
    {
        get
        {
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(local, "XRai", "studio", "preferences.json");
        }
    }

    public static StudioPreferences Load()
    {
        var path = PreferencesPath;
        if (!File.Exists(path)) return new StudioPreferences();

        try
        {
            var text = File.ReadAllText(path);
            var prefs = JsonSerializer.Deserialize<StudioPreferences>(text, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
            });
            return prefs ?? new StudioPreferences();
        }
        catch (Exception ex)
        {
            // Corrupt file — log to Debug (which goes to the daemon's
            // attached debugger / DebugView) and return defaults so the
            // user isn't stuck in a broken state. The next Save will
            // overwrite the corrupt file with valid JSON.
            System.Diagnostics.Debug.WriteLine(
                $"StudioPreferences.Load: corrupt file at {path} ({ex.GetType().Name}: {ex.Message}). Resetting to defaults.");
            return new StudioPreferences();
        }
    }

    public void Save()
    {
        try
        {
            var path = PreferencesPath;
            var dir = Path.GetDirectoryName(path)!;
            Directory.CreateDirectory(dir);
            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
        catch
        {
            // Preference save failures are non-fatal — Studio still works
            // with in-memory defaults.
        }
    }

    public JsonObject ToJson() => new()
    {
        ["schemaVersion"] = SchemaVersion,
        ["followMode"] = FollowMode,
        ["preferredIde"] = PreferredIde,
        ["onboarded"] = Onboarded,
        ["theme"] = Theme,
    };

    public static StudioPreferences FromJson(JsonNode? node)
    {
        var prefs = new StudioPreferences();
        if (node is not JsonObject obj) return prefs;
        try
        {
            if (obj["followMode"]?.GetValue<bool?>() is bool fm) prefs.FollowMode = fm;
            if (obj["preferredIde"]?.GetValue<string?>() is string pi) prefs.PreferredIde = pi;
            if (obj["onboarded"]?.GetValue<bool?>() is bool ob) prefs.Onboarded = ob;
            if (obj["theme"]?.GetValue<string?>() is string th) prefs.Theme = th;
        }
        catch { }
        return prefs;
    }
}
