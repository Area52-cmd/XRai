// Leak-audited: 2026-04-10 — Register replaces existing entries by key so
// repeated Pilot.Expose calls do not accumulate adapter instances. As of the
// Studio instrumentation, ControlAdapter subscribes to DependencyPropertyDescriptor
// value-changed callbacks, so the registry now proactively disposes the old
// adapter on re-registration to release those subscriptions.

using System.Windows;

namespace XRai.Hooks;

public class ControlRegistry
{
    private readonly Dictionary<string, IControlAdapter> _controls = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The root FrameworkElement passed to Pilot.Expose — used by pane_screenshot
    /// to capture the entire WPF pane via RenderTargetBitmap.
    /// </summary>
    public FrameworkElement? RootElement { get; set; }

    public int Count => _controls.Count;

    public void Register(string name, IControlAdapter adapter)
    {
        if (_controls.TryGetValue(name, out var existing) && existing is IDisposable d)
        {
            try { d.Dispose(); } catch { }
        }
        _controls[name] = adapter;
    }

    public bool TryGet(string name, out IControlAdapter adapter)
    {
        if (_controls.TryGetValue(name, out adapter!)) return true;

        // Pathed lookup: "ItemsControl[0].ChildName", "Grid[key=Id:7].Btn", etc.
        // Only triggered when the flat dict misses AND the name contains '['.
        if (!string.IsNullOrEmpty(name) && name.IndexOf('[') >= 0)
        {
            if (ControlPathResolver.TryResolve(this, name, out var resolved) && resolved != null)
            {
                adapter = resolved;
                return true;
            }
        }

        adapter = null!;
        return false;
    }

    public IEnumerable<KeyValuePair<string, IControlAdapter>> All => _controls;

    /// <summary>
    /// Dispose every registered adapter — releases all DependencyPropertyDescriptor
    /// subscriptions so the old visual tree can be GC'd cleanly. Called before
    /// a fresh Pilot.Expose walks a newly-built pane.
    /// </summary>
    public void Clear()
    {
        foreach (var adapter in _controls.Values)
        {
            if (adapter is IDisposable d)
            {
                try { d.Dispose(); } catch { }
            }
        }
        _controls.Clear();
    }
}
