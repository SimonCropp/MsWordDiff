using System.Text.Json;

public class SettingsManager(string settingsPath)
{
    static readonly JsonSerializerOptions jsonOptions = new()
    {
        WriteIndented = true
    };

    public static string DefaultSettingsPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        ".config",
        "MsWordDiff",
        "settings.json");

    public async Task<Settings> ReadAsync()
    {
        if (!File.Exists(settingsPath))
        {
            return new();
        }

        try
        {
            await using var stream = File.OpenRead(settingsPath);
            return await JsonSerializer.DeserializeAsync<Settings>(stream) ?? new();
        }
        catch (Exception exception)
        {
            Log.Warning(exception, "Failed to read settings from {Path}", settingsPath);
            return new();
        }
    }

    public async Task WriteAsync(Settings settings)
    {
        var directory = Path.GetDirectoryName(settingsPath);
        if (directory != null)
        {
            Directory.CreateDirectory(directory);
        }

        await using var stream = File.Create(settingsPath);
        await JsonSerializer.SerializeAsync(stream, settings, jsonOptions);
    }

    public async Task SetQuietAsync(bool value)
    {
        var settings = await ReadAsync();
        settings.Quiet = value;
        await WriteAsync(settings);
    }
}
