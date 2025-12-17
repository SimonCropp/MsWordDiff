public class FileWatcherManagerTests
{
    [Test]
    public async Task FileChanged_TriggersCallbackAfterDebounce()
    {
        var tempFile1 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");
        var tempFile2 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");

        try
        {
            // Create initial files
            await File.WriteAllTextAsync(tempFile1, "initial");
            await File.WriteAllTextAsync(tempFile2, "initial");

            var callbackCount = 0;
            using var watcher = new FileWatcherManager(tempFile1, tempFile2, () => callbackCount++);

            // Modify file1
            await File.WriteAllTextAsync(tempFile1, "modified");

            // Wait for debounce (500ms) + buffer
            await Task.Delay(700);

            await Assert.That(callbackCount).IsEqualTo(1);
        }
        finally
        {
            if (File.Exists(tempFile1)) File.Delete(tempFile1);
            if (File.Exists(tempFile2)) File.Delete(tempFile2);
        }
    }

    [Test]
    public async Task FileCreated_TriggersCallback()
    {
        var tempFile1 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");
        var tempFile2 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");

        try
        {
            // Create initial files
            await File.WriteAllTextAsync(tempFile1, "initial");
            await File.WriteAllTextAsync(tempFile2, "initial");

            var callbackCount = 0;
            using var watcher = new FileWatcherManager(tempFile1, tempFile2, () => callbackCount++);

            // Simulate Word's save pattern: delete and recreate
            File.Delete(tempFile1);
            await Task.Delay(100);
            await File.WriteAllTextAsync(tempFile1, "recreated");

            // Wait for debounce
            await Task.Delay(700);

            await Assert.That(callbackCount).IsGreaterThanOrEqualTo(1);
        }
        finally
        {
            if (File.Exists(tempFile1)) File.Delete(tempFile1);
            if (File.Exists(tempFile2)) File.Delete(tempFile2);
        }
    }

    [Test]
    public async Task MultipleRapidChanges_DebouncesToSingleCallback()
    {
        var tempFile1 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");
        var tempFile2 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");

        try
        {
            // Create initial files
            await File.WriteAllTextAsync(tempFile1, "initial");
            await File.WriteAllTextAsync(tempFile2, "initial");

            var callbackCount = 0;
            using var watcher = new FileWatcherManager(tempFile1, tempFile2, () => callbackCount++);

            // Make multiple rapid changes
            for (var i = 0; i < 5; i++)
            {
                await File.WriteAllTextAsync(tempFile1, $"change {i}");
                await Task.Delay(50); // Less than debounce time
            }

            // Wait for debounce
            await Task.Delay(700);

            // Should only trigger once due to debouncing
            await Assert.That(callbackCount).IsEqualTo(1);
        }
        finally
        {
            if (File.Exists(tempFile1)) File.Delete(tempFile1);
            if (File.Exists(tempFile2)) File.Delete(tempFile2);
        }
    }

    [Test]
    public async Task BothFiles_TriggerCallbackIndependently()
    {
        var tempFile1 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");
        var tempFile2 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");

        try
        {
            // Create initial files
            await File.WriteAllTextAsync(tempFile1, "initial");
            await File.WriteAllTextAsync(tempFile2, "initial");

            var callbackCount = 0;
            using var watcher = new FileWatcherManager(tempFile1, tempFile2, () => callbackCount++);

            // Modify file1
            await File.WriteAllTextAsync(tempFile1, "modified1");
            await Task.Delay(700);

            var firstCount = callbackCount;
            await Assert.That(firstCount).IsEqualTo(1);

            // Modify file2
            await File.WriteAllTextAsync(tempFile2, "modified2");
            await Task.Delay(700);

            await Assert.That(callbackCount).IsEqualTo(2);
        }
        finally
        {
            if (File.Exists(tempFile1)) File.Delete(tempFile1);
            if (File.Exists(tempFile2)) File.Delete(tempFile2);
        }
    }

    [Test]
    public async Task Dispose_StopsWatching()
    {
        var tempFile1 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");
        var tempFile2 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");

        try
        {
            // Create initial files
            await File.WriteAllTextAsync(tempFile1, "initial");
            await File.WriteAllTextAsync(tempFile2, "initial");

            var callbackCount = 0;
            var watcher = new FileWatcherManager(tempFile1, tempFile2, () => callbackCount++);

            // Dispose watcher
            watcher.Dispose();

            // Modify file after disposal
            await File.WriteAllTextAsync(tempFile1, "modified after dispose");
            await Task.Delay(700);

            // Should not have triggered callback
            await Assert.That(callbackCount).IsEqualTo(0);
        }
        finally
        {
            if (File.Exists(tempFile1)) File.Delete(tempFile1);
            if (File.Exists(tempFile2)) File.Delete(tempFile2);
        }
    }

    [Test]
    public async Task SimulateWordSavePattern_TriggersCallback()
    {
        var tempFile1 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.docx");
        var tempFile2 = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.docx");
        var tempFileBackup = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.tmp");

        try
        {
            // Create initial files
            await File.WriteAllTextAsync(tempFile1, "original content");
            await File.WriteAllTextAsync(tempFile2, "original content");

            var callbackCount = 0;
            using var watcher = new FileWatcherManager(tempFile1, tempFile2, () => callbackCount++);

            // Simulate Word's atomic save pattern:
            // 1. Create temp file with new content
            await File.WriteAllTextAsync(tempFileBackup, "new content");
            await Task.Delay(50);

            // 2. Delete original
            File.Delete(tempFile1);
            await Task.Delay(50);

            // 3. Rename temp to original
            File.Move(tempFileBackup, tempFile1);

            // Wait for debounce
            await Task.Delay(700);

            // Should have triggered callback
            await Assert.That(callbackCount).IsGreaterThanOrEqualTo(1);
        }
        finally
        {
            if (File.Exists(tempFile1)) File.Delete(tempFile1);
            if (File.Exists(tempFile2)) File.Delete(tempFile2);
            if (File.Exists(tempFileBackup)) File.Delete(tempFileBackup);
        }
    }

    [Test]
    public void InvalidFilePath_ThrowsException()
    {
        var invalidPath = "Q:\\NonExistent\\Path\\file.txt";
        var tempFile = Path.Combine(Path.GetTempPath(), $"test-watch-{Guid.NewGuid()}.txt");

        try
        {
            File.WriteAllText(tempFile, "test");

            Assert.Throws<ArgumentException>(() =>
            {
                using var watcher = new FileWatcherManager(invalidPath, tempFile, () => { });
            });
        }
        finally
        {
            if (File.Exists(tempFile)) File.Delete(tempFile);
        }
    }
}
