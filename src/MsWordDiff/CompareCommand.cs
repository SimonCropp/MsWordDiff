[Command(Description = "Compare two Word documents side by side using Microsoft Word")]
public class CompareCommand(SettingsManager settingsManager) : ICommand
{
    [CommandParameter(0, Name = "path1", Description = "Path to the first Word document")]
    public required FileInfo Path1 { get; init; }

    [CommandParameter(1, Name = "path2", Description = "Path to the second Word document")]
    public required FileInfo Path2 { get; init; }

    [CommandOption("quiet", Description = "Hide source documents in the comparison view. Default can be set via 'set-quiet' command")]
    public bool? Quiet { get; init; }

    [CommandOption("watch", Description = "Watch source files and auto-refresh comparison when they change")]
    public bool Watch { get; init; }

    public async ValueTask ExecuteAsync(IConsole console)
    {
        if (!Path1.Exists)
        {
            throw new CommandException($"File does not exist: {Path1.FullName}");
        }

        if (!Path2.Exists)
        {
            throw new CommandException($"File does not exist: {Path2.FullName}");
        }

        var settings = await settingsManager.Read();

        var quiet = Quiet ?? settings.Quiet;

        try
        {
            Word.Launch(Path1.FullName, Path2.FullName, quiet, Watch);
        }
        catch (Exception exception)
        {
            Log.Error(exception, "Failed to launch");
            throw new CommandException("Failed to launch", 1, innerException: exception);
        }
    }
}
