using System.ComponentModel;
using System.Reflection;

namespace XRai.Hooks;

public class ModelAdapter
{
    private readonly INotifyPropertyChanged _model;
    private readonly Dictionary<string, PropertyInfo> _properties;

    public ModelAdapter(INotifyPropertyChanged model)
    {
        _model = model;
        _properties = model.GetType()
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead)
            .ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);
    }

    public Dictionary<string, object?> GetAll()
    {
        var result = new Dictionary<string, object?>();
        foreach (var kvp in _properties)
        {
            try
            {
                result[kvp.Key] = kvp.Value.GetValue(_model);
            }
            catch
            {
                result[kvp.Key] = null;
            }
        }
        return result;
    }

    public object? GetProperty(string name)
    {
        if (!_properties.TryGetValue(name, out var prop))
            throw new ArgumentException($"Property not found: {name}");
        return prop.GetValue(_model);
    }

    public void SetProperty(string name, object? value)
    {
        if (!_properties.TryGetValue(name, out var prop))
            throw new ArgumentException($"Property not found: {name}");
        if (!prop.CanWrite)
            throw new InvalidOperationException($"Property '{name}' is read-only");

        var converted = Convert.ChangeType(value, prop.PropertyType);
        prop.SetValue(_model, converted);
    }
}
