static class TempFiles
{
    public static readonly string TempDirectory = Path.Combine(Path.GetTempPath(), "MsExcelDiff");

    static TempFiles()
    {
        Directory.CreateDirectory(TempDirectory);
        CleanOld(TempDirectory);
    }

    public static string Create(string content) =>
        Create(TempDirectory, content);

    internal static string Create(string directory, string content)
    {
        var path = Path.Combine(directory, $"{Guid.NewGuid()}.txt");
        File.WriteAllText(path, content);
        return path;
    }

    public static bool TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch
        {
            // Best effort; CleanOld will get it next time
        }

        return false;
    }

    internal static void CleanOld(string directory)
    {
        try
        {
            var cutoff = DateTime.UtcNow.AddDays(-1);
            foreach (var file in Directory.EnumerateFiles(directory))
            {
                try
                {
                    if (File.GetLastWriteTimeUtc(file) < cutoff)
                    {
                        File.Delete(file);
                    }
                }
                catch
                {
                    // File may be in use
                }
            }
        }
        catch
        {
            // Best effort cleanup
        }
    }
}
