public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        Logging.Init();

        return await new CliApplicationBuilder()
            .AddCommandsFromThisAssembly()
            .Build()
            .RunAsync(args);
    }
}
