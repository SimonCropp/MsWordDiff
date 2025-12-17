[Command("settings", Description = "Display the settings")]
public class SettingsPathCommand : ICommand
{
    public string SettingsPath { get; init; } = SettingsManager.DefaultSettingsPath;

    public async ValueTask ExecuteAsync(IConsole console)
    {
        await console.Output.WriteLineAsync(SettingsPath);
        if (File.Exists(SettingsPath))
        {
            await console.Output.WriteLineAsync(await File.ReadAllTextAsync(SettingsPath));
            return;
        }

        await console.Output.WriteLineAsync("No settings file found.");
    }
}