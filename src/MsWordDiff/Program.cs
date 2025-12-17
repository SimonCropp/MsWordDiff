public static class Program
{
    public static int Main(string[] args)
    {
        Logging.Init();
        var command = BuildCommand(Word.Launch);
        var parseResult = command.Parse(args);

        if (parseResult.Errors.Count > 0)
        {
            Console.Error.WriteLine("Usage: diffword <path1> <path2>");
            Console.Error.WriteLine();
            foreach (var error in parseResult.Errors)
            {
                Console.Error.WriteLine(error.Message);
            }
            return 1;
        }

        return parseResult.Invoke();
    }

    public static RootCommand BuildCommand(Action<string, string> launchAction)
    {
        var path1Argument = new Argument<FileInfo>("path1")
        {
            Description = "Path to the first Word document",
            Arity = ArgumentArity.ExactlyOne
        };

        var path2Argument = new Argument<FileInfo>("path2")
        {
            Description = "Path to the second Word document",
            Arity = ArgumentArity.ExactlyOne
        };

        var rootCommand = new RootCommand("Compare two Word documents side by side using Microsoft Word");
        rootCommand.Arguments.Add(path1Argument);
        rootCommand.Arguments.Add(path2Argument);

        rootCommand.SetAction(parseResult =>
        {
            var path1 = parseResult.GetValue(path1Argument)!;
            var path2 = parseResult.GetValue(path2Argument)!;

            if (!path1.Exists)
            {
                Console.Error.WriteLine($"File not found: {path1.FullName}");
                return 1;
            }

            if (!path2.Exists)
            {
                Console.Error.WriteLine($"File not found: {path2.FullName}");
                return 1;
            }

            try
            {
                launchAction(path1.FullName, path2.FullName);
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
