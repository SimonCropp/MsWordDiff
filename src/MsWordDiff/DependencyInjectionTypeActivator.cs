public class DependencyInjectionTypeActivator(IServiceProvider serviceProvider) : ITypeActivator
{
    public object CreateInstance(Type type)
    {
        var instance = serviceProvider.GetService(type);
        if (instance != null)
        {
            return instance;
        }

        // Fallback to manual construction with dependency resolution
        var constructors = type.GetConstructors();
        if (constructors.Length == 0)
        {
            throw new InvalidOperationException($"Type {type.Name} has no public constructors");
        }

        var constructor = constructors[0];
        var parameters = constructor.GetParameters();
        var args = new object[parameters.Length];

        for (var i = 0; i < parameters.Length; i++)
        {
            args[i] = serviceProvider.GetService(parameters[i].ParameterType)
                      ?? throw new InvalidOperationException(
                          $"Unable to resolve dependency {parameters[i].ParameterType.Name} for {type.Name}");
        }

        return constructor.Invoke(args);
    }
}
