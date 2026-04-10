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
