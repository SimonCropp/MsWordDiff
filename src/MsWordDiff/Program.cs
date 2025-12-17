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

    public static CliApplicationBuilder CreateBuilder()
    {
        var builder = new CliApplicationBuilder();
        builder.AddCommandsFromThisAssembly();
        builder.SetExecutableName("diffword");
        return builder;
    }
}
