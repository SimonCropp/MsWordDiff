class FileWatcherManager : IDisposable
{
    FileSystemWatcher watcher1;
    FileSystemWatcher watcher2;
    System.Timers.Timer debounceTimer;

    public FileWatcherManager(string path1, string path2, Action onChanged)
    {
        watcher1 = CreateWatcher(path1);
        watcher2 = CreateWatcher(path2);

        debounceTimer = new(500)
        {
            AutoReset = false
        };
        debounceTimer.Elapsed += (_, _) => onChanged();
    }

    FileSystemWatcher CreateWatcher(string filePath)
    {
        var directory = Path.GetDirectoryName(filePath);
        if (directory == null)
        {
            throw new InvalidOperationException($"Could not determine directory for {filePath}");
        }

        var fileName = Path.GetFileName(filePath);

        var watcher = new FileSystemWatcher(directory, fileName)
        {
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size,
            EnableRaisingEvents = true
        };

        watcher.Changed += OnFileChanged;
        return watcher;
    }

    void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        debounceTimer.Stop();
        debounceTimer.Start();
    }

    public void Dispose()
    {
        watcher1.Dispose();
        watcher2.Dispose();
        debounceTimer.Dispose();
    }
}