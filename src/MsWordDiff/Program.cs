public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        Logging.Init();

        var builder = CreateBuilder();
        return await builder
            .Build()
            .RunAsync(args);
    }

    public static CliApplicationBuilder CreateBuilder(string? settingsPath = null)
    {
        var services = new ServiceCollection();
        services.AddSingleton(new SettingsManager(settingsPath ?? SettingsManager.DefaultSettingsPath));

        var serviceProvider = services.BuildServiceProvider();
        var typeActivator = new DependencyInjectionTypeActivator(serviceProvider);

        var builder = new CliApplicationBuilder();
        builder.AddCommandsFromThisAssembly();
        builder.SetExecutableName("diffword");
        builder.UseTypeActivator(typeActivator);
        return builder;
    }
}
