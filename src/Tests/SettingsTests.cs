public class SettingsTests
{
    [Test]
    public async Task ReadNonExistentSettings_ReturnsDefaultSettings()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"msworddiff-test-{Guid.NewGuid()}.json");

        try
        {
            var settingsManager = new SettingsManager(tempPath);
            var settings = await settingsManager.Read();

            await Assert.That(settings.Quiet).IsFalse();
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    [Test]
    public async Task WriteAndReadSettings_PreservesValues()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"msworddiff-test-{Guid.NewGuid()}.json");

        try
        {
            var settingsManager = new SettingsManager(tempPath);
            var settings = new Settings { Quiet = true };

            await settingsManager.Write(settings);

            var readSettings = await settingsManager.Read();

            await Assert.That(readSettings.Quiet).IsTrue();
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }

    [Test]
    public async Task ReadCorruptedSettings_ReturnsDefaultSettings()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"msworddiff-test-{Guid.NewGuid()}.json");

        try
        {
            await File.WriteAllTextAsync(tempPath, "{ invalid json }");

            var settingsManager = new SettingsManager(tempPath);
            var settings = await settingsManager.Read();

            await Assert.That(settings.Quiet).IsFalse();
        }
        finally
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
    }
}
