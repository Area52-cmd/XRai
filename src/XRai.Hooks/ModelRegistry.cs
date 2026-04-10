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
        _models[name] = new ModelAdapter(model);
    }

    /// <summary>
    /// Mark this model as the "default" returned by {"cmd":"model"} when no
    /// name is supplied. Pilot.ExposeModel always sets this so callers don't
    /// have to guess which model HandleModel will return.
    /// </summary>
    public void SetDefault(INotifyPropertyChanged model)
    {
        _default = new ModelAdapter(model);
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
