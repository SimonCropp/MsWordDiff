public static class Program
{
    public static int Main(string[] args)
    {
        Logging.Init();
        var command = BuildCommand(Word.Launch);
        return command.Parse(args).Invoke();
    }

    public static RootCommand BuildCommand(Action<string, string> launchAction, TextWriter? errorOutput = null)
    {
        errorOutput ??= Console.Error;

        var path1Argument = new Argument<FileInfo?>("path1")
        {
            Description = "Path to the first Word document",
            Arity = ArgumentArity.ZeroOrOne
        };

        var path2Argument = new Argument<FileInfo?>("path2")
        {
            Description = "Path to the second Word document",
            Arity = ArgumentArity.ZeroOrOne
        };

        var rootCommand = new RootCommand("Compare two Word documents side by side using Microsoft Word");
        rootCommand.Arguments.Add(path1Argument);
        rootCommand.Arguments.Add(path2Argument);

        rootCommand.SetAction(parseResult =>
        {
            var path1 = parseResult.GetValue(path1Argument);
            var path2 = parseResult.GetValue(path2Argument);

            var errors = new List<string>();

            if (path1 is null)
            {
                errors.Add("Required argument missing: <path1>");
            }
            else if (!path1.Exists)
            {
                errors.Add($"File not found: {path1.FullName}");
            }

            if (path2 is null)
            {
                errors.Add("Required argument missing: <path2>");
            }
            else if (!path2.Exists)
            {
                errors.Add($"File not found: {path2.FullName}");
            }

            if (errors.Count > 0)
            {
                errorOutput.WriteLine("Usage: diffword <path1> <path2>");
                errorOutput.WriteLine();
                foreach (var error in errors)
                {
                    errorOutput.WriteLine(error);
                }
                return 1;
            }

            try
            {
                launchAction(path1!.FullName, path2!.FullName);
                return 0;
            }
            catch (Exception exception)
            {
                Log.Error(exception, "Failed to launch");
                return 1;
            }
        });

        return rootCommand;
    }
}
