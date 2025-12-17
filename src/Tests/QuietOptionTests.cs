public class QuietOptionTests
{
    [Test]
    public async Task QuietOption_ShowsInHelp()
    {
        using var console = new FakeInMemoryConsole();

        var services = new ServiceCollection();
        services.AddSingleton(new SettingsManager(SettingsManager.DefaultSettingsPath));

        var serviceProvider = services.BuildServiceProvider();
        var typeActivator = new DependencyInjectionTypeActivator(serviceProvider);

        var app = new CliApplicationBuilder()
            .AddCommand<CompareCommand>()
            .UseConsole(console)
            .UseTypeActivator(typeActivator)
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
        using var tempPath = TempFile.Create();

        using var console = new FakeInMemoryConsole();

        var services = new ServiceCollection();
        services.AddSingleton(new SettingsManager(tempPath));

        var serviceProvider = services.BuildServiceProvider();
        var typeActivator = new DependencyInjectionTypeActivator(serviceProvider);

        var app = new CliApplicationBuilder()
            .AddCommand<SetQuietCommand>()
            .UseConsole(console)
            .UseTypeActivator(typeActivator)
            .Build();

        var exitCode = await app.RunAsync(["set-quiet", "true"]);

        await Assert.That(exitCode).IsEqualTo(0);

        var output = console.ReadOutputString();
        await Assert.That(output).Contains("Quiet mode default set to: True");
    }

    [Test]
    public async Task SettingsPathCommand_DisplaysPath()
    {
        using var console = new FakeInMemoryConsole();

        var services = new ServiceCollection();
        services.AddSingleton(new SettingsManager(SettingsManager.DefaultSettingsPath));

        var serviceProvider = services.BuildServiceProvider();
        var typeActivator = new DependencyInjectionTypeActivator(serviceProvider);

        var app = new CliApplicationBuilder()
            .AddCommand<SettingsPathCommand>()
            .UseConsole(console)
            .UseTypeActivator(typeActivator)
            .Build();

        var exitCode = await app.RunAsync(["settings"]);

        await Assert.That(exitCode).IsEqualTo(0);

        var output = console.ReadOutputString();
        await Assert.That(output).Contains(SettingsManager.DefaultSettingsPath);
    }
}
