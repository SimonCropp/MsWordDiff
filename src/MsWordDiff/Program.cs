public static class Program
{
    static int Main(string[] args)
    {
        Logging.Init();
        if (args.Length != 2)
        {
            Console.Error.WriteLine("Usage: diffword <path1> <path2>");
            return 1;
        }

        var path1 = Path.GetFullPath(args[0]);
        var path2 = Path.GetFullPath(args[1]);

        if (!File.Exists(path1))
        {
            Log.Error("File not found: {Path1}", path1);
            return 1;
        }

        if (!File.Exists(path2))
        {
            Log.Error("File not found: {Path2}", path2);
            return 1;
        }

        try
        {
            Word.Launch(path1, path2);
            return 0;
        }
        catch (Exception exception)
        {
            Log.Error(exception, "Failed to launch");
            return 1;
        }
    }
}
