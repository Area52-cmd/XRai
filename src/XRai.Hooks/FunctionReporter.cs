using System.Reflection;
using ExcelDna.Integration;

namespace XRai.Hooks;

public static class FunctionReporter
{
    public static object[] GetRegisteredFunctions()
    {
        var results = new List<object>();

        // Search loaded assemblies for ExcelFunction attributes
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
        {
            try
            {
                foreach (var type in asm.GetTypes())
                {
                    foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Static))
                    {
                        var attr = method.GetCustomAttribute<ExcelFunctionAttribute>();
                        if (attr == null) continue;

                        var parameters = method.GetParameters().Select(p => new
                        {
                            name = p.Name,
                            type = p.ParameterType.Name,
                        }).ToArray();

                        results.Add(new
                        {
                            name = attr.Name ?? method.Name,
                            description = attr.Description,
                            parameters,
                        });
                    }
                }
            }
            catch
            {
                // Skip assemblies that can't be reflected
            }
        }

        return results.ToArray();
    }
}
