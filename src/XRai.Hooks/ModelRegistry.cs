// Leak-audited: 2026-04-10 — Register replaces existing entries by key, so
// repeat ExposeModel calls do not accumulate. Default tracked explicitly so
// the most-recently-set model wins regardless of dictionary insertion order.

using System.ComponentModel;

namespace XRai.Hooks;

public class ModelRegistry
{
    private readonly Dictionary<string, ModelAdapter> _models = new(StringComparer.OrdinalIgnoreCase);
    private ModelAdapter? _default;

    public void Register(INotifyPropertyChanged model, string name)
    {
        // Dispose the previous adapter (if any) so its PropertyChanged
        // subscription is released. Prevents a slow leak on repeat ExposeModel.
        if (_models.TryGetValue(name, out var existing))
        {
            try { existing.Dispose(); } catch { }
        }
        var adapter = new ModelAdapter(model);
        adapter.SetName(name);
        _models[name] = adapter;
    }

    /// <summary>
    /// Mark this model as the "default" returned by {"cmd":"model"} when no
    /// name is supplied. Pilot.ExposeModel always sets this so callers don't
    /// have to guess which model HandleModel will return.
    /// </summary>
    public void SetDefault(INotifyPropertyChanged model)
    {
        // Dispose the previous default adapter if it was a distinct instance
        // (avoid double-dispose if it's also in _models).
        if (_default != null && !_models.ContainsValue(_default))
        {
            try { _default.Dispose(); } catch { }
        }
        var adapter = new ModelAdapter(model);
        adapter.SetName("default");
        _default = adapter;
    }

    public bool TryGet(string name, out ModelAdapter adapter)
    {
        return _models.TryGetValue(name, out adapter!);
    }

    public IEnumerable<KeyValuePair<string, ModelAdapter>> All => _models;

    /// <summary>
    /// Returns the explicitly-set default if any, otherwise the first
    /// registered model (legacy behavior). Never throws.
    /// </summary>
    public ModelAdapter? Default => _default ?? _models.Values.FirstOrDefault();
}
