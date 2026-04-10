// Leak-audited: 2026-04-10 — Register replaces existing entries by key so
// repeated Pilot.Expose calls do not accumulate adapter instances. The
// registry holds a strong reference to each FrameworkElement adapter;
// rerunning Expose with a freshly-built pane replaces the old adapters but
// does NOT proactively unsubscribe events from the old elements. Since
// adapters only hold the FrameworkElement and don't subscribe to its events,
// no event-handler leak is possible — the old adapters become unreachable
// once Register replaces them and the GC can collect them.

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
        _controls[name] = adapter;
    }

    public bool TryGet(string name, out IControlAdapter adapter)
    {
        return _controls.TryGetValue(name, out adapter!);
    }

    public IEnumerable<KeyValuePair<string, IControlAdapter>> All => _controls;
}
