using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;

namespace XRai.Hooks;

public class ModelAdapter : IDisposable
{
    /// <summary>
    /// In-process hook fired whenever any exposed ViewModel property changes.
    /// Args: (modelName, propertyName, oldValue, newValue). XRai.Studio
    /// subscribes here to get the live ViewModel state stream. If nothing
    /// subscribes, this is a zero-cost no-op.
    ///
    /// Subscribers run on whatever thread raised PropertyChanged (typically
    /// the WPF UI dispatcher). Keep them fast; do not call back into the UI
    /// via Dispatcher.Invoke or you risk deadlocks.
    /// </summary>
    public static event Action<string, string, object?, object?>? OnModelChanged;

    private readonly INotifyPropertyChanged _model;
    private readonly Dictionary<string, PropertyInfo> _properties;
    private readonly Dictionary<string, object?> _lastValues = new(StringComparer.OrdinalIgnoreCase);
    private string _modelName = "default";
    private bool _disposed;

    public ModelAdapter(INotifyPropertyChanged model)
    {
        _model = model;
        _properties = model.GetType()
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead)
            .ToDictionary(p => p.Name, StringComparer.OrdinalIgnoreCase);

        // Seed last-known values so the first change event has an accurate 'old'.
        foreach (var kvp in _properties)
        {
            try { _lastValues[kvp.Key] = kvp.Value.GetValue(_model); }
            catch { _lastValues[kvp.Key] = null; }
        }

        // Subscribe to PropertyChanged so the Studio gets a live stream.
        // Wrapped in try/catch: some models implement INotifyPropertyChanged
        // but throw when the event is subscribed (rare but possible).
        try { _model.PropertyChanged += OnPropertyChanged; }
        catch (Exception ex) { Debug.WriteLine($"ModelAdapter subscribe failed: {ex.Message}"); }
    }

    /// <summary>
    /// Called by ModelRegistry.Register(model, name) so the live change event
    /// carries the correct model name instead of the generic "default".
    /// </summary>
    public void SetName(string name) => _modelName = name;

    private void OnPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (_disposed) return;
        if (string.IsNullOrEmpty(e.PropertyName)) return;
        if (!_properties.TryGetValue(e.PropertyName, out var prop)) return;

        object? newVal;
        try { newVal = prop.GetValue(_model); }
        catch { newVal = null; }

        _lastValues.TryGetValue(e.PropertyName, out var oldVal);
        _lastValues[e.PropertyName] = newVal;

        // Fire the static event outside the try so subscriber exceptions
        // surface in Debug but don't break the adapter.
        try { OnModelChanged?.Invoke(_modelName, e.PropertyName, oldVal, newVal); }
        catch (Exception ex) { Debug.WriteLine($"OnModelChanged subscriber threw: {ex.Message}"); }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try { _model.PropertyChanged -= OnPropertyChanged; } catch { }
    }

    public Dictionary<string, object?> GetAll()
    {
        var result = new Dictionary<string, object?>();
        foreach (var kvp in _properties)
        {
            try
            {
                var raw = kvp.Value.GetValue(_model);
                result[kvp.Key] = SafeValue(raw);
            }
            catch
            {
                result[kvp.Key] = null;
            }
        }
        return result;
    }

    /// <summary>
    /// Convert a property value to a JSON-safe representation. WPF types
    /// like SolidColorBrush, Transform, Geometry have cyclical references
    /// (Transform.Inverse → Transform) that blow up JsonSerializer with
    /// "object cycle detected at depth 64". We convert known-problematic
    /// types to safe string representations and fall back to ToString()
    /// for anything that isn't a primitive, string, or collection.
    /// </summary>
    private static object? SafeValue(object? raw)
    {
        if (raw == null) return null;

        var type = raw.GetType();

        // Primitives, strings, enums — safe as-is
        if (type.IsPrimitive || type == typeof(string) || type == typeof(decimal)
            || type == typeof(DateTime) || type == typeof(DateTimeOffset)
            || type == typeof(TimeSpan) || type == typeof(Guid)
            || type.IsEnum)
        {
            return raw;
        }

        // Collections of primitives/strings — safe as-is
        if (raw is System.Collections.IEnumerable enumerable && type != typeof(string))
        {
            try
            {
                var list = new List<object?>();
                foreach (var item in enumerable)
                {
                    list.Add(SafeValue(item));
                    if (list.Count > 200) { list.Add("... (truncated)"); break; }
                }
                return list;
            }
            catch
            {
                return raw.ToString();
            }
        }

        // Known WPF types that cause cycles — convert to string
        var ns = type.Namespace ?? "";
        if (ns.StartsWith("System.Windows.Media", StringComparison.Ordinal)
            || ns.StartsWith("System.Windows.Input", StringComparison.Ordinal)
            || ns.StartsWith("System.Windows.Threading", StringComparison.Ordinal))
        {
            return raw.ToString();
        }

        // Nullable<T> — unwrap
        var underlying = Nullable.GetUnderlyingType(type);
        if (underlying != null) return SafeValue(raw);

        // Anything else — try ToString to avoid cycles
        try { return raw.ToString(); }
        catch { return $"<{type.Name}>"; }
    }

    public object? GetProperty(string name)
    {
        if (!_properties.TryGetValue(name, out var prop))
            throw new ArgumentException($"Property not found: {name}");
        return SafeValue(prop.GetValue(_model));
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
