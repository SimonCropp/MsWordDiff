public class SettingsManager(string settingsPath)
{
    static JsonSerializerOptions jsonOptions = new()
    {
        WriteIndented = true
    };

    public static string DefaultSettingsPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
        ".config",
        "MsWordDiff",
        "settings.json");

    public string SettingsPath => settingsPath;

    public async Task<Settings> Read()
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

    public async Task Write(Settings settings)
    {
        var directory = Path.GetDirectoryName(settingsPath);
        if (directory != null)
        {
            Directory.CreateDirectory(directory);
        }

        await using var stream = File.Create(settingsPath);
        await JsonSerializer.SerializeAsync(stream, settings, jsonOptions);
    }

    public async Task SetQuiet(bool value)
    {
        var settings = await Read();
        settings.Quiet = value;
        await Write(settings);
    }

    public async Task Setup()
    {
        if (File.Exists(settingsPath))
        {
            return;
        }

        await File.WriteAllTextAsync(settingsPath, "{}");
        var result = MessageBox.Show(
            """
            Threre are two UX modes. Standard and Quiet.
            
            Standard shows the comparsion, both documents being compared, and the list of differencs.
            
            Quiet shows on the comparison.
            
            Should Quiet be used?
            
            Note: This can be changed later using the set-quiet command line option.
            """,
            "Standard or Quiet UX?",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);
        var quiet = result == DialogResult.Yes;
        await SetQuiet(quiet);
    }
}