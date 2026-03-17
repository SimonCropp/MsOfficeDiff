public class DependencyInjectionTypeActivator(IServiceProvider provider) :
    ITypeActivator
{
    public object CreateInstance(Type type) => provider.GetRequiredService(type);
}
