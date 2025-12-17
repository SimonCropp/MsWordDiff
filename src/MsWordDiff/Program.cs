public static class Program
{
    public static int Main(string[] args)
    {
        Logging.Init();
        var command = BuildCommand(Word.Launch);
        var parseResult = command.Parse(args);

        if (parseResult.Errors.Count > 0)
        {
            var builder = new StringBuilder();
            builder.AppendLine("Usage: diffword <path1> <path2>");
            builder.AppendLine();
            foreach (var error in parseResult.Errors)
            {
                // Improve generic "Required argument missing" messages
                var message = error.Message;
                if (message.Contains("Required argument missing"))
                {
                    // Extract argument name from SymbolResult if available
                    if (error.SymbolResult is ArgumentResult argResult)
                    {
                        message = $"Required argument missing: <{argResult.Argument.Name}>";
                    }
                }

                builder.AppendLine(message);
            }
            Log.Error("Failed to launch: {result}", builder.ToString());

            return 1;
        }

        return parseResult.Invoke();
    }

    public static RootCommand BuildCommand(Action<string, string> launchAction)
    {
        var path1Argument = new Argument<FileInfo>("path1")
        {
            Description = "Path to the first Word document",
            Arity = ArgumentArity.ExactlyOne,
            Validators =
            {
                fileExists
            }
        };

        var path2Argument = new Argument<FileInfo>("path2")
        {
            Description = "Path to the second Word document",
            Arity = ArgumentArity.ExactlyOne,
            Validators =
            {
                fileExists
            }
        };

        var rootCommand = new RootCommand("Compare two Word documents side by side using Microsoft Word");
        rootCommand.Arguments.Add(path1Argument);
        rootCommand.Arguments.Add(path2Argument);

        rootCommand.SetAction(parseResult =>
        {
            var path1 = parseResult.GetValue(path1Argument);
            var path2 = parseResult.GetValue(path2Argument);

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

    static Action<ArgumentResult> fileExists =
        result =>
        {
            var file = result.GetValueOrDefault<FileInfo>();
            if (!file.Exists)
            {
                result.AddError($"File does not exist: {file.FullName}");
            }
        };
}