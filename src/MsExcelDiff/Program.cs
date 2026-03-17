public static class Program
{
    static List<Type> commands = typeof(Program)
        .Assembly
        .GetTypes()
        .Where(_ => _.IsAssignableTo(typeof(ICommand)) && !_.IsAbstract)
        .ToList();

    public static async Task<int> Main(string[] args)
    {
        Logging.Init();

        await using var services = ServiceProvider(null);
        var builder = CreateBuilder(services);
        return await builder
            .Build()
            .RunAsync(args);
    }

    public static CliApplicationBuilder CreateBuilder(string? settingsPath = null)
    {
        var services = ServiceProvider(settingsPath);
        return CreateBuilder(services);
    }

    static ServiceProvider ServiceProvider(string? settingsPath)
    {
        var settingsManager = new SettingsManager(settingsPath ?? SettingsManager.DefaultSettingsPath);
        var services = new ServiceCollection();
        services.AddSingleton(settingsManager);

        foreach (var type in commands)
        {
            services.AddSingleton(type);
        }

        return services.BuildServiceProvider();
    }

    static CliApplicationBuilder CreateBuilder(ServiceProvider serviceProvider)
    {
        var builder = new CliApplicationBuilder();
        foreach (var type in commands)
        {
            builder.AddCommand(type);
        }
        var activator = new DependencyInjectionTypeActivator(serviceProvider);
        builder.SetExecutableName("diffexcel");
        builder.UseTypeActivator(activator);
        return builder;
    }
}
