public class QuietOptionTests
{
    [Test]
    public async Task QuietOption_ShowsInHelp()
    {
        using var console = new FakeInMemoryConsole();

        var app = new CliApplicationBuilder()
            .AddCommand<CompareCommand>()
            .UseConsole(console)
            .Build();

        var exitCode = await app.RunAsync(["--help"]);

        var output = console.ReadOutputString();

        await Assert.That(exitCode).IsEqualTo(0);
        await Assert.That(output).Contains("--quiet");
        await Verify(output);
    }

    [Test]
    public async Task SetQuietCommand_UpdatesSettings()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"msworddiff-test-{Guid.NewGuid()}.json");

        try
        {
            using var console = new FakeInMemoryConsole();

            var app = new CliApplicationBuilder()
                .AddCommand<SetQuietCommand>()
                .UseConsole(console)
                .Build();

            var exitCode = await app.RunAsync(["set-quiet", "true"]);

            await Assert.That(exitCode).IsEqualTo(0);

            var output = console.ReadOutputString();
            await Assert.That(output).Contains("Quiet mode default set to: True");
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
    public async Task SettingsPathCommand_DisplaysPath()
    {
        using var console = new FakeInMemoryConsole();

        var app = new CliApplicationBuilder()
            .AddCommand<SettingsPathCommand>()
            .UseConsole(console)
            .Build();

        var exitCode = await app.RunAsync(["settings", "path"]);

        await Assert.That(exitCode).IsEqualTo(0);

        var output = console.ReadOutputString();
        await Assert.That(output).Contains(SettingsManager.DefaultSettingsPath);
    }
}
