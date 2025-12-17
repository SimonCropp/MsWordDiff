using CliFx.Attributes;
using CliFx.Infrastructure;

[Command(Description = "Compare two Word documents side by side using Microsoft Word")]
public class CompareCommand : ICommand
{
    [CommandParameter(0, Name = "path1", Description = "Path to the first Word document")]
    public required FileInfo Path1 { get; init; }

    [CommandParameter(1, Name = "path2", Description = "Path to the second Word document")]
    public required FileInfo Path2 { get; init; }

    public ValueTask ExecuteAsync(IConsole console)
    {
        if (!Path1.Exists)
        {
            console.Error.WriteLine($"File does not exist: {Path1.FullName}");
            return ValueTask.CompletedTask;
        }

        if (!Path2.Exists)
        {
            console.Error.WriteLine($"File does not exist: {Path2.FullName}");
            return ValueTask.CompletedTask;
        }

        try
        {
            Word.Launch(Path1.FullName, Path2.FullName);
            return ValueTask.CompletedTask;
        }
        catch (Exception exception)
        {
            Log.Error(exception, "Failed to launch");
            return ValueTask.FromException(new CommandException("Failed to launch", 1, innerException: exception));
        }
    }
}
