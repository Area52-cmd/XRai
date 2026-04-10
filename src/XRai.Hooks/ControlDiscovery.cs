using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

// Disambiguate WPF vs WinForms types (UseWindowsForms=true in csproj)
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using TextBoxBase = System.Windows.Controls.Primitives.TextBoxBase;
using Selector = System.Windows.Controls.Primitives.Selector;
using ComboBox = System.Windows.Controls.ComboBox;
using Label = System.Windows.Controls.Label;
using ProgressBar = System.Windows.Controls.ProgressBar;
using TreeView = System.Windows.Controls.TreeView;

namespace XRai.Hooks;

/// <summary>
/// Walks a WPF visual tree and registers every element into the ControlRegistry.
/// - Elements WITH x:Name are registered by their declared name.
/// - Elements WITHOUT x:Name that are interactive (Button, TextBox, ComboBox,
///   CheckBox, RadioButton, DataGrid, etc.) are registered with a synthetic name
///   of the form "_unnamed_{Type}_{Index}" so agents can still target them.
/// - Non-interactive unnamed elements (Border, Grid, StackPanel) are skipped to
///   avoid drowning the registry in template chrome.
///
/// This fixes the pre-Round-7 bug where empty-name elements all collided on the
/// same empty key in the registry dictionary and silently dropped every unnamed
/// interactive control.
/// </summary>
public static class ControlDiscovery
{
    private const int MaxControls = 10_000;

    public static void Walk(FrameworkElement root, ControlRegistry registry)
    {
        // Force the visual tree to be fully built before walking. Without this,
        // controls nested 6+ levels deep inside ScrollViewer/ItemsControl/TabControl
        // may not have been measured/arranged yet, so VisualTreeHelper.GetChildrenCount
        // returns 0 and the walker stops short.
        try { root.ApplyTemplate(); } catch { }
        try { root.UpdateLayout(); } catch { }

        var counters = new Dictionary<string, int>();
        var visited = new HashSet<nint>(); // prevent double-registration across trees
        int registered = 0;

        // Primary: visual tree — the authoritative tree for rendered elements.
        WalkVisualTree(root, registry, counters, visited, ref registered);

        // Secondary: logical tree — catches controls inside DataTemplates,
        // ContentPresenters, and deferred-instantiation containers that haven't
        // been realized into the visual tree yet. Safe because `visited` prevents
        // double-registration of elements found in both trees.
        WalkLogicalTree(root, registry, counters, visited, ref registered);
    }

    private static void WalkVisualTree(
        DependencyObject parent, ControlRegistry registry,
        Dictionary<string, int> counters, HashSet<nint> visited, ref int registered)
    {
        if (registered >= MaxControls) return;

        int childCount = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < childCount; i++)
        {
            if (registered >= MaxControls) return;

            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is FrameworkElement fe && visited.Add(GetId(fe)))
            {
                TryRegister(fe, registry, counters);
                registered++;
            }

            WalkVisualTree(child, registry, counters, visited, ref registered);
        }
    }

    private static void WalkLogicalTree(
        DependencyObject parent, ControlRegistry registry,
        Dictionary<string, int> counters, HashSet<nint> visited, ref int registered)
    {
        if (registered >= MaxControls) return;

        foreach (var child in LogicalTreeHelper.GetChildren(parent))
        {
            if (registered >= MaxControls) return;

            if (child is not DependencyObject dobj) continue;

            if (child is FrameworkElement fe && visited.Add(GetId(fe)))
            {
                TryRegister(fe, registry, counters);
                registered++;
            }

            WalkLogicalTree(dobj, registry, counters, visited, ref registered);
        }
    }

    /// <summary>
    /// Identity hash for dedup across visual + logical tree walks.
    /// Uses RuntimeHelpers to get a stable identity for the same object
    /// even if GetHashCode is overridden.
    /// </summary>
    private static nint GetId(FrameworkElement fe) =>
        (nint)System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(fe);

    private static void TryRegister(FrameworkElement fe, ControlRegistry registry, Dictionary<string, int> counters)
    {
        // Named: register by its declared name (canonical path)
        if (!string.IsNullOrEmpty(fe.Name))
        {
            registry.Register(fe.Name, new ControlAdapter(fe));
            return;
        }

        // Unnamed: only register if the element is interactive enough to be
        // meaningful as a click/read target. Skip pure layout/chrome elements.
        if (!IsInteractive(fe)) return;

        var typeName = fe.GetType().Name;
        if (!counters.TryGetValue(typeName, out var count)) count = 0;
        counters[typeName] = count + 1;

        var syntheticName = $"_unnamed_{typeName}_{count}";

        // Attach an extracted label if the element has one (Button content,
        // TextBlock text, etc.) so agents can still correlate it visually.
        var label = ExtractLabel(fe);
        if (!string.IsNullOrWhiteSpace(label))
        {
            // Include a short label suffix — sanitized to keep it registry-safe
            var cleanLabel = new string(label.Take(40).Where(c => char.IsLetterOrDigit(c) || c == ' ').ToArray()).Trim().Replace(' ', '_');
            if (!string.IsNullOrWhiteSpace(cleanLabel))
                syntheticName = $"_unnamed_{typeName}_{cleanLabel}_{count}";
        }

        registry.Register(syntheticName, new ControlAdapter(fe));
    }

    private static bool IsInteractive(FrameworkElement fe)
    {
        return fe is ButtonBase         // Button, ToggleButton, RadioButton, CheckBox, etc.
            || fe is TextBoxBase        // TextBox, RichTextBox
            || fe is Selector           // ComboBox, ListBox, ListView, TabControl, DataGrid
            || fe is Slider
            || fe is ProgressBar
            || fe is DatePicker
            || fe is Calendar
            || fe is Expander
            || fe is MenuItem
            || fe is PasswordBox
            || fe is TreeView
            || fe is TreeViewItem;
    }

    private static string? ExtractLabel(FrameworkElement fe)
    {
        try
        {
            // Buttons and similar usually put their label in Content
            if (fe is ContentControl cc && cc.Content is string s)
                return s;
            if (fe is ContentControl cc2 && cc2.Content is TextBlock tb)
                return tb.Text;

            // Labels
            if (fe is Label lbl && lbl.Content is string ls)
                return ls;

            // TextBlocks (rarely registered but handled)
            if (fe is TextBlock tb2)
                return tb2.Text;

            // ComboBox placeholder
            if (fe is ComboBox cb && cb.Items.Count > 0)
                return $"ComboBox_{cb.Items.Count}items";
        }
        catch { }
        return null;
    }
}
