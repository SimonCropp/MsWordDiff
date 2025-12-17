[Command("set-quiet", Description = "Set the default value for the Quiet option. Settings are saved in the user configuration file.")]
public class SetQuietCommand(SettingsManager settingsManager) : ICommand
{
    [CommandParameter(0, Name = "value", Description = "true to enable Quiet mode by default, false to disable")]
    public required bool Value { get; init; }

    public async ValueTask ExecuteAsync(IConsole console)
    {
        await settingsManager.SetQuiet(Value);

        await console.Output.WriteLineAsync($"Quiet mode default set to: {Value}");
        await console.Output.WriteLineAsync($"Settings saved to: {settingsManager.SettingsPath}");
    }
}