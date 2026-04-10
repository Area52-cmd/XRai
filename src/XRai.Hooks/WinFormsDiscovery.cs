using System.Windows.Forms;

namespace XRai.Hooks;

/// <summary>
/// Walks a WinForms control tree and registers every control into the ControlRegistry.
/// Mirrors the WPF ControlDiscovery logic: named controls are registered by name,
/// unnamed interactive controls get synthetic names.
/// </summary>
public static class WinFormsDiscovery
{
    public static void Walk(Control root, ControlRegistry registry)
    {
        var counters = new Dictionary<string, int>();
        WalkControlTree(root, registry, counters);
    }

    private static void WalkControlTree(Control parent, ControlRegistry registry, Dictionary<string, int> counters)
    {
        foreach (Control child in parent.Controls)
        {
            if (!string.IsNullOrEmpty(child.Name))
            {
                registry.Register(child.Name, new WinFormsAdapter(child));
            }
            else if (IsInteractive(child))
            {
                // Synthetic naming — same pattern as WPF ControlDiscovery
                var typeName = child.GetType().Name;
                if (!counters.TryGetValue(typeName, out var count)) count = 0;
                counters[typeName] = count + 1;
                var syntheticName = $"_unnamed_{typeName}_{count}";

                var label = child.Text;
                if (!string.IsNullOrWhiteSpace(label))
                {
                    var cleanLabel = new string(label.Take(40)
                        .Where(c => char.IsLetterOrDigit(c) || c == ' ')
                        .ToArray()).Trim().Replace(' ', '_');
                    if (!string.IsNullOrWhiteSpace(cleanLabel))
                        syntheticName = $"_unnamed_{typeName}_{cleanLabel}_{count}";
                }

                registry.Register(syntheticName, new WinFormsAdapter(child));
            }

            WalkControlTree(child, registry, counters);
        }
    }

    private static bool IsInteractive(Control c)
    {
        return c is Button
            || c is TextBox
            || c is ComboBox
            || c is ListBox
            || c is CheckBox
            || c is RadioButton
            || c is NumericUpDown
            || c is TrackBar
            || c is DateTimePicker
            || c is DataGridView
            || c is TabControl
            || c is TreeView
            || c is RichTextBox
            || c is MaskedTextBox
            || c is ProgressBar
            || c is Label;
    }
}
