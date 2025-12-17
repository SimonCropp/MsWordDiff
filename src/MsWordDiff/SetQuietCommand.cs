[Command("set-quiet", Description = "Set the default value for the Quiet option. Settings are saved in the user configuration file.")]
public class SetQuietCommand : ICommand
{
    [CommandParameter(0, Name = "value", Description = "true to enable Quiet mode by default, false to disable")]
    public required bool Value { get; init; }

    public string SettingsPath { get; init; } = SettingsManager.DefaultSettingsPath;

    public async ValueTask ExecuteAsync(IConsole console)
    {
        var settingsManager = new SettingsManager(SettingsPath);
        await settingsManager.SetQuietAsync(Value);

        await console.Output.WriteLineAsync($"Quiet mode default set to: {Value}");
        await console.Output.WriteLineAsync($"Settings saved to: {SettingsPath}");
    }
}