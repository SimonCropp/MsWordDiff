public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        Logging.Init();

        var settingsManager = new SettingsManager(SettingsManager.DefaultSettingsPath);
        await settingsManager.Setup();
        var builder = CreateBuilder(settingsManager);
        return await builder
            .Build()
            .RunAsync(args);
    }

    public static CommandLineApplicationBuilder CreateBuilder(string? settingsPath = null)
    {
        var settingsManager = new SettingsManager(settingsPath ?? SettingsManager.DefaultSettingsPath);
        return CreateBuilder(settingsManager);
    }

    static CommandLineApplicationBuilder CreateBuilder(SettingsManager settingsManager)
    {
        var builder = new CommandLineApplicationBuilder();
        builder.AddCommandsFromThisAssembly();
        builder.SetExecutableName("diffword");
        builder.UseTypeInstantiator(commands =>
        {
            var services = new ServiceCollection();
            services.AddSingleton(settingsManager);
            foreach (var command in commands)
            {
                services.AddSingleton(command.Type);
            }
            return services.BuildServiceProvider();
        });
        return builder;
    }
}