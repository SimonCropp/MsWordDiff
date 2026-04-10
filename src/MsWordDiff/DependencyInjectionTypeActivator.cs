public class DependencyInjectionTypeActivator(IServiceProvider provider) :
    ITypeInstantiator
{
    public object CreateInstance(Type type) => provider.GetRequiredService(type);
}
