public static class Program
{
    public static int Main(string[] args)
    {
        Logging.Init();
        var command = BuildCommand(Word.Launch);
        return command.Invoke(args);
    }

    public static RootCommand BuildCommand(Action<string, string> launchAction)
    {
        var path1Argument = new Argument<FileInfo>("path1", "Path to the first Word document")
        {
            Arity = ArgumentArity.ExactlyOne
        }.ExistingOnly();

        var path2Argument = new Argument<FileInfo>("path2", "Path to the second Word document")
        {
            Arity = ArgumentArity.ExactlyOne
        }.ExistingOnly();

        var rootCommand = new RootCommand("Compare two Word documents side by side using Microsoft Word")
        {
            path1Argument,
            path2Argument
        };
        rootCommand.Name = "diff";

        rootCommand.SetHandler((path1, path2) =>
        {
            try
            {
                launchAction(path1.FullName, path2.FullName);
            }
            catch (Exception exception)
            {
                Log.Error(exception, "Failed to launch");
                throw;
            }
        }, path1Argument, path2Argument);

        return rootCommand;
    }
}
