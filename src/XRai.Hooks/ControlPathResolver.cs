using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace XRai.Hooks;

/// <summary>
/// Resolves pathed control lookups like:
///   "TunersListView[0].DiagToggle"
///   "TunersListView[key=Symbol:AAPL].PriceText"
///   "OuterList[0].InnerList[2].ActionButton"
///
/// Works against ItemsControl (ListView, ListBox, DataGrid, TreeView, etc.). Lazily
/// resolves indices at command time — never eagerly enumerates virtualized rows.
/// Handles VirtualizingStackPanel by calling ScrollIntoView and waiting for the
/// generator to materialize the container before descending into the template.
///
/// If the path contains no '[' the resolver returns false immediately and the
/// registry falls back to its normal flat-name dictionary lookup.
/// </summary>
internal static class ControlPathResolver
{
    // Segment: Name or Name[index] or Name[key=Prop:Value]
    private static readonly Regex SegmentRx = new(
        @"^(?<name>[A-Za-z_][A-Za-z0-9_]*)(?:\[(?<sel>[^\]]+)\])?$",
        RegexOptions.Compiled);

    public static bool TryResolve(
        ControlRegistry registry,
        string path,
        out IControlAdapter? adapter)
    {
        adapter = null;
        if (string.IsNullOrWhiteSpace(path)) return false;
        if (path.IndexOf('[') < 0) return false; // flat name — let caller handle

        var segments = path.Split('.');
        if (segments.Length == 0) return false;

        // First segment must resolve against the registry (the rooted ItemsControl
        // or named container the user exposed via x:Name).
        var first = SegmentRx.Match(segments[0]);
        if (!first.Success) return false;

        if (!registry.TryGet(first.Groups["name"].Value, out var rootAdapter))
            return false;

        if (rootAdapter is not ControlAdapter ca) return false;
        FrameworkElement? current = ca.Element;

        // Apply the first segment's selector (if any), then descend.
        if (first.Groups["sel"].Success)
        {
            current = ResolveItemContainer(current, first.Groups["sel"].Value);
            if (current == null) return false;
        }

        for (int i = 1; i < segments.Length; i++)
        {
            var m = SegmentRx.Match(segments[i]);
            if (!m.Success) return false;

            var childName = m.Groups["name"].Value;

            // Find the named child inside current's visual + logical subtree.
            var child = FindNamedDescendant(current!, childName);
            if (child == null) return false;

            current = child;

            if (m.Groups["sel"].Success)
            {
                current = ResolveItemContainer(current, m.Groups["sel"].Value);
                if (current == null) return false;
            }
        }

        if (current == null) return false;
        adapter = new ControlAdapter(current);
        return true;
    }

    /// <summary>
    /// Resolve a selector like "0" or "key=Symbol:AAPL" against an ItemsControl.
    /// Returns the materialized container (ListBoxItem, ListViewItem, DataGridRow, etc.).
    /// </summary>
    private static FrameworkElement? ResolveItemContainer(FrameworkElement? host, string selector)
    {
        if (host is not ItemsControl ic) return null;
        if (ic.Items == null || ic.Items.Count == 0) return null;

        int? index = null;

        if (int.TryParse(selector, out var idx))
        {
            index = idx;
        }
        else if (selector.StartsWith("key=", StringComparison.OrdinalIgnoreCase))
        {
            var rest = selector.Substring(4);
            var colon = rest.IndexOf(':');
            if (colon < 0) return null;
            var propName = rest.Substring(0, colon);
            var wanted = rest.Substring(colon + 1);

            for (int i = 0; i < ic.Items.Count; i++)
            {
                var item = ic.Items[i];
                if (item == null) continue;
                var pi = item.GetType().GetProperty(propName);
                if (pi == null) continue;
                var val = pi.GetValue(item)?.ToString();
                if (string.Equals(val, wanted, StringComparison.OrdinalIgnoreCase))
                {
                    index = i;
                    break;
                }
            }
        }

        if (index == null || index < 0 || index >= ic.Items.Count) return null;

        // Force the container to materialize — matters for VirtualizingStackPanel.
        var container = ic.ItemContainerGenerator.ContainerFromIndex(index.Value) as FrameworkElement;
        if (container == null)
        {
            try
            {
                // ScrollIntoView forces the virtualizing panel to realize this row.
                if (ic is ListBox lb) lb.ScrollIntoView(ic.Items[index.Value]);
                else if (ic is DataGrid dg) dg.ScrollIntoView(ic.Items[index.Value]);
                ic.UpdateLayout();
            }
            catch { }

            // Container generation can be asynchronous inside a virtualizing panel.
            // Pump the dispatcher briefly (up to ~500ms) while waiting for it.
            var deadline = DateTime.UtcNow.AddMilliseconds(500);
            while (container == null && DateTime.UtcNow < deadline)
            {
                try
                {
                    container = ic.ItemContainerGenerator.ContainerFromIndex(index.Value) as FrameworkElement;
                    if (container != null) break;

                    // Drain pending UI work at Background priority so the generator ticks.
                    var frame = new System.Windows.Threading.DispatcherFrame();
                    System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        new Action(() => frame.Continue = false));
                    System.Windows.Threading.Dispatcher.PushFrame(frame);
                }
                catch { break; }
            }
        }

        // Ensure the container's template is realized so FindNamedDescendant can
        // see the named children inside it.
        try { container?.ApplyTemplate(); } catch { }
        try { container?.UpdateLayout(); } catch { }

        return container;
    }

    /// <summary>
    /// Breadth-first search for a child FrameworkElement with x:Name == name,
    /// searching both visual and logical trees. Stops at nested ItemsControls
    /// so that "Outer[0].Inner[2].Btn" resolves correctly rather than greedily
    /// descending into Outer's own templates.
    /// </summary>
    private static FrameworkElement? FindNamedDescendant(FrameworkElement root, string name)
    {
        try { root.ApplyTemplate(); } catch { }

        var queue = new Queue<DependencyObject>();
        queue.Enqueue(root);

        while (queue.Count > 0)
        {
            var node = queue.Dequeue();

            int vcount = VisualTreeHelper.GetChildrenCount(node);
            for (int i = 0; i < vcount; i++)
            {
                var child = VisualTreeHelper.GetChild(node, i);
                if (child is FrameworkElement fe && fe.Name == name)
                    return fe;
                queue.Enqueue(child);
            }

            foreach (var lchild in LogicalTreeHelper.GetChildren(node))
            {
                if (lchild is FrameworkElement fe2 && fe2.Name == name)
                    return fe2;
                if (lchild is DependencyObject dobj)
                    queue.Enqueue(dobj);
            }
        }

        return null;
    }
}
