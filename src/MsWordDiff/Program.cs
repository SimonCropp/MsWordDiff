public static class Program
{
    static List<Type> commands = typeof(Program).Assembly.GetTypes()
        .Where(_ => _.IsAssignableTo(typeof(ICommand)) && !_.IsAbstract).ToList();

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

        var builder = new CliApplicationBuilder();
        foreach (var type in commands)
        {
            builder.AddCommand(type);
            services.AddSingleton(type);
        }

        var serviceProvider = services.BuildServiceProvider();
        var typeActivator = new DependencyInjectionTypeActivator(serviceProvider);
        builder.SetExecutableName("diffword");
        builder.UseTypeActivator(typeActivator);
        return builder;
    }
}