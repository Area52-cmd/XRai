using System.Runtime.InteropServices;
using System.Windows.Forms;
using WinButton = System.Windows.Forms.Button;
using WinTextBox = System.Windows.Forms.TextBox;
using WinComboBox = System.Windows.Forms.ComboBox;
using WinListBox = System.Windows.Forms.ListBox;
using WinCheckBox = System.Windows.Forms.CheckBox;
using WinRadioButton = System.Windows.Forms.RadioButton;
using WinLabel = System.Windows.Forms.Label;
using WinProgressBar = System.Windows.Forms.ProgressBar;
using WinTabControl = System.Windows.Forms.TabControl;
using WinTreeView = System.Windows.Forms.TreeView;

namespace XRai.Hooks;

/// <summary>
/// Wraps a System.Windows.Forms.Control and implements IControlAdapter
/// so PipeServer can interact with WinForms controls identically to WPF ones.
/// </summary>
public class WinFormsAdapter : IControlAdapter
{
    private readonly Control _control;

    public WinFormsAdapter(Control control)
    {
        _control = control ?? throw new ArgumentNullException(nameof(control));
        Type = control switch
        {
            DataGridView => "DataGrid",
            WinTabControl => "TabControl",
            WinTreeView => "TreeView",
            WinRadioButton => "RadioButton",
            DateTimePicker => "DatePicker",
            TrackBar => "Slider",
            WinProgressBar => "ProgressBar",
            RichTextBox => "RichTextBox",
            MaskedTextBox => "MaskedTextBox",
            WinTextBox => "TextBox",
            WinButton => "Button",
            WinLabel => "Label",
            WinComboBox => "ComboBox",
            WinCheckBox => "CheckBox",
            WinListBox => "ListBox",
            NumericUpDown => "NumericUpDown",
            _ => control.GetType().Name,
        };
    }

    public string Name => _control.Name;
    public string Type { get; }
    public bool IsEnabled => _control.Enabled;
    public bool IsVisible => _control.Visible;
    public bool HasCommand => false; // WinForms doesn't have ICommand

    /// <summary>
    /// WPF element access — not available for WinForms controls.
    /// </summary>
    public System.Windows.FrameworkElement? Element => null;

    public string? GetValue()
    {
        return _control switch
        {
            WinTextBox tb => tb.Text,
            RichTextBox rtb => rtb.Text,
            MaskedTextBox mtb => mtb.Text,
            WinLabel lbl => lbl.Text,
            WinComboBox cb => cb.SelectedItem?.ToString() ?? cb.Text,
            WinListBox lb => lb.SelectedItem?.ToString(),
            WinRadioButton rb => rb.Checked.ToString(),
            WinCheckBox chk => chk.Checked.ToString(),
            NumericUpDown nud => nud.Value.ToString(),
            TrackBar tb => tb.Value.ToString(),
            WinProgressBar pb => pb.Value.ToString(),
            DateTimePicker dtp => dtp.Value.ToString("yyyy-MM-dd"),
            WinTabControl tc => tc.SelectedIndex.ToString(),
            DataGridView dgv => $"DataGrid[{dgv.RowCount} rows x {dgv.ColumnCount} cols]",
            WinTreeView tv => $"TreeView[{tv.Nodes.Count} root nodes]",
            _ => _control.Text,
        };
    }

    public void SetValue(string value)
    {
        switch (_control)
        {
            case WinTextBox tb:
                tb.Text = value;
                break;
            case RichTextBox rtb:
                rtb.Text = value;
                break;
            case MaskedTextBox mtb:
                mtb.Text = value;
                break;
            case WinComboBox cb:
                // Try to select by item text first
                for (int i = 0; i < cb.Items.Count; i++)
                {
                    if (cb.Items[i]?.ToString() == value)
                    {
                        cb.SelectedIndex = i;
                        return;
                    }
                }
                cb.Text = value;
                break;
            case WinListBox lb:
                for (int i = 0; i < lb.Items.Count; i++)
                {
                    if (lb.Items[i]?.ToString() == value)
                    {
                        lb.SelectedIndex = i;
                        return;
                    }
                }
                break;
            case WinRadioButton rb:
                rb.Checked = bool.TryParse(value, out var rbVal) && rbVal;
                break;
            case WinCheckBox chk:
                chk.Checked = bool.TryParse(value, out var chkVal) && chkVal;
                break;
            case NumericUpDown nud:
                if (decimal.TryParse(value, out var dVal))
                    nud.Value = Math.Max(nud.Minimum, Math.Min(nud.Maximum, dVal));
                break;
            case TrackBar tb:
                if (int.TryParse(value, out var iVal))
                    tb.Value = Math.Max(tb.Minimum, Math.Min(tb.Maximum, iVal));
                break;
            case DateTimePicker dtp:
                if (DateTime.TryParse(value, out var dt))
                    dtp.Value = dt;
                break;
            case WinTabControl tc:
                if (int.TryParse(value, out var idx) && idx >= 0 && idx < tc.TabCount)
                    tc.SelectedIndex = idx;
                else
                    SelectTabByName(tc, value);
                break;
            default:
                throw new InvalidOperationException($"Cannot set value on {Type} control");
        }
    }

    public ControlAdapter.ClickResult Click()
    {
        var result = new ControlAdapter.ClickResult();

        if (_control is WinButton btn)
        {
            result.ResolvedToButtonBase = true;
            result.ResolvedTargetType = btn.GetType().Name;
            result.HasCommand = false;
            try
            {
                btn.PerformClick();
                result.Method = "Button.PerformClick";
            }
            catch (Exception ex)
            {
                result.ErrorHint = $"PerformClick threw: {ex.Message}";
                result.Method = "Button.PerformClick (handler threw)";
            }
            return result;
        }

        // Non-button: focus + post ENTER to the control's window only.
        // Previously used global System.Windows.Forms.SendKeys.Send("{ENTER}"),
        // which calls SendInput under the hood and caused stuck-modifier-key
        // bugs when the target lost focus (e.g. Excel killed during rebuild)
        // between the keydown and keyup. PostMessage targets the control's
        // HWND directly so no global keyboard state is touched.
        result.ResolvedToButtonBase = false;
        result.ResolvedTargetType = _control.GetType().Name;
        try
        {
            _control.Focus();
            PostKey(_control.Handle, VK_RETURN);
            result.Method = "Focus+PostMessage(WM_KEYDOWN/WM_KEYUP VK_RETURN)";
        }
        catch (Exception ex)
        {
            result.ErrorHint = $"Click fallback threw: {ex.Message}";
            result.Method = "Focus+PostMessage(ENTER) (threw)";
        }
        return result;
    }

    // ── Safe keystroke delivery (no global SendInput) ─────────────────
    // All keystroke paths go through PostMessage on the control's HWND so
    // we never touch the OS-level keyboard state. Prevents the stuck Ctrl
    // / Shift / Alt bug the global SendKeys.Send path used to cause when
    // focus shifted mid-keystroke (common when Excel is killed during
    // rebuild). Scoped window messages never leave modifiers dangling.
    private const uint WM_KEYDOWN = 0x0100;
    private const uint WM_KEYUP = 0x0101;
    private const uint WM_CHAR = 0x0102;
    private const int VK_RETURN = 0x0D;
    private const int VK_TAB = 0x09;
    private const int VK_ESCAPE = 0x1B;
    private const int VK_BACK = 0x08;
    private const int VK_DELETE = 0x2E;
    private const int VK_SPACE = 0x20;
    private const int VK_LEFT = 0x25;
    private const int VK_UP = 0x26;
    private const int VK_RIGHT = 0x27;
    private const int VK_DOWN = 0x28;
    private const int VK_HOME = 0x24;
    private const int VK_END = 0x23;
    private const int VK_F1 = 0x70;

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    private static void PostKey(IntPtr hwnd, int vk)
    {
        if (hwnd == IntPtr.Zero) return;
        PostMessage(hwnd, WM_KEYDOWN, (IntPtr)vk, IntPtr.Zero);
        PostMessage(hwnd, WM_KEYUP, (IntPtr)vk, IntPtr.Zero);
    }

    private static void PostChar(IntPtr hwnd, char c)
    {
        if (hwnd == IntPtr.Zero) return;
        PostMessage(hwnd, WM_CHAR, (IntPtr)c, IntPtr.Zero);
    }

    /// <summary>
    /// Translate a common key token (ENTER, TAB, F1, ESCAPE, etc.) into a
    /// Win32 virtual-key code. Returns 0 if not a known named key — caller
    /// then treats it as a literal character.
    /// </summary>
    private static int NamedVk(string token) => token.ToUpperInvariant() switch
    {
        "ENTER" or "RETURN" => VK_RETURN,
        "TAB" => VK_TAB,
        "ESC" or "ESCAPE" => VK_ESCAPE,
        "BACK" or "BACKSPACE" => VK_BACK,
        "DEL" or "DELETE" => VK_DELETE,
        "SPACE" => VK_SPACE,
        "LEFT" => VK_LEFT,
        "UP" => VK_UP,
        "RIGHT" => VK_RIGHT,
        "DOWN" => VK_DOWN,
        "HOME" => VK_HOME,
        "END" => VK_END,
        "F1" => VK_F1,
        _ => 0,
    };

    public void Toggle()
    {
        switch (_control)
        {
            case WinCheckBox chk:
                chk.Checked = !chk.Checked;
                break;
            case WinRadioButton rb:
                rb.Checked = !rb.Checked;
                break;
            default:
                throw new InvalidOperationException($"Cannot toggle {Type} control");
        }
    }

    public void Focus()
    {
        _control.Focus();
    }

    public void DoubleClick()
    {
        // WinForms doesn't expose a programmatic double-click API.
        // Focus and send a synthetic message via reflection on the OnDoubleClick method.
        _control.Focus();
        var method = typeof(Control).GetMethod("OnDoubleClick",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        method?.Invoke(_control, new object[] { EventArgs.Empty });
    }

    public void RightClick()
    {
        _control.Focus();
        var method = typeof(Control).GetMethod("OnMouseClick",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        method?.Invoke(_control, new object[]
        {
            new MouseEventArgs(MouseButtons.Right, 1, _control.Width / 2, _control.Height / 2, 0)
        });
    }

    public void Hover()
    {
        var method = typeof(Control).GetMethod("OnMouseEnter",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        method?.Invoke(_control, new object[] { EventArgs.Empty });
    }

    public string Expand(bool open = true)
    {
        switch (_control)
        {
            case WinComboBox cb:
                cb.DroppedDown = open;
                return open ? "ComboBox dropdown opened" : "ComboBox dropdown closed";
            case WinTreeView tv when tv.SelectedNode != null:
                if (open) tv.SelectedNode.Expand(); else tv.SelectedNode.Collapse();
                return open ? "TreeNode expanded" : "TreeNode collapsed";
            default:
                throw new InvalidOperationException(
                    $"pane.expand not supported on {_control.GetType().Name}. " +
                    "Supported: ComboBox, TreeView (with selected node).");
        }
    }

    public ControlAdapter.SendKeysResult SendKeys(string keys)
    {
        var result = new ControlAdapter.SendKeysResult { InputKeys = keys };

        try { _control.Focus(); }
        catch (Exception ex) { result.FocusWarning = ex.Message; }

        // Replaced the old System.Windows.Forms.SendKeys.Send(keys) path —
        // that used global SendInput and would leave Ctrl/Shift/Alt stuck
        // on the OS keyboard state if focus shifted mid-keystroke (exactly
        // what happens when a rebuild kills Excel with a keystroke still
        // in flight). We now post window-scoped messages: WM_KEYDOWN/
        // WM_KEYUP for named keys (ENTER, TAB, F1, arrows, ...) and WM_CHAR
        // for literal text characters. Scoped to the control's HWND so no
        // global keyboard state is ever modified.
        var hwnd = _control.Handle;
        if (hwnd == IntPtr.Zero)
        {
            result.FailedKeys.Add($"{keys}: control has no window handle");
            result.ErrorHint = "Control is not yet realized (Handle == 0).";
            return result;
        }

        try
        {
            // Plus-separated tokens: "ENTER", "F1", "Hello" (last is literal).
            foreach (var raw in keys.Split('+'))
            {
                var token = raw.Trim();
                if (token.Length == 0) continue;

                var vk = NamedVk(token);
                if (vk != 0)
                {
                    PostKey(hwnd, vk);
                    result.DeliveredKeys.Add(token);
                }
                else
                {
                    // Literal text: emit one WM_CHAR per character.
                    foreach (var ch in token) PostChar(hwnd, ch);
                    result.DeliveredKeys.Add(token);
                }
            }
        }
        catch (Exception ex)
        {
            result.FailedKeys.Add($"{keys}: {ex.Message}");
            result.ErrorHint = ex.Message;
        }

        return result;
    }

    public (string[] Items, int SelectedIndex) ReadListItems()
    {
        switch (_control)
        {
            case WinComboBox cb:
            {
                var items = new string[cb.Items.Count];
                for (int i = 0; i < cb.Items.Count; i++)
                    items[i] = cb.Items[i]?.ToString() ?? "";
                return (items, cb.SelectedIndex);
            }
            case WinListBox lb:
            {
                var items = new string[lb.Items.Count];
                for (int i = 0; i < lb.Items.Count; i++)
                    items[i] = lb.Items[i]?.ToString() ?? "";
                return (items, lb.SelectedIndex);
            }
            default:
                throw new InvalidOperationException(
                    $"pane.list.read requires a ComboBox or ListBox. Got: {_control.GetType().Name}");
        }
    }

    public string SelectListItem(int? index, string? text)
    {
        switch (_control)
        {
            case WinComboBox cb:
                return SelectInCollection(cb.Items, i => cb.SelectedIndex = i, cb.Items.Count, index, text);
            case WinListBox lb:
                return SelectInCollection(lb.Items, i => lb.SelectedIndex = i, lb.Items.Count, index, text);
            default:
                throw new InvalidOperationException(
                    $"pane.list.select requires a ComboBox or ListBox. Got: {_control.GetType().Name}");
        }
    }

    public object? GetDetailedInfo()
    {
        return _control switch
        {
            DataGridView dgv => new
            {
                rows = dgv.RowCount,
                columns = dgv.ColumnCount,
                selected_index = dgv.CurrentRow?.Index ?? -1,
                column_headers = dgv.Columns.Cast<DataGridViewColumn>()
                    .Select(c => c.HeaderText).ToArray(),
            },
            WinTabControl tc => new
            {
                selected_index = tc.SelectedIndex,
                tab_count = tc.TabCount,
                tabs = tc.TabPages.Cast<TabPage>()
                    .Select(tp => tp.Text).ToArray(),
            },
            WinListBox lb => new
            {
                item_count = lb.Items.Count,
                selected_index = lb.SelectedIndex,
                items = lb.Items.Cast<object>().Take(50)
                    .Select(i => i.ToString()).ToArray(),
            },
            WinComboBox cb => new
            {
                selected_index = cb.SelectedIndex,
                item_count = cb.Items.Count,
                items = cb.Items.Cast<object>().Take(50)
                    .Select(i => i.ToString()).ToArray(),
            },
            WinTreeView tv => new
            {
                item_count = tv.Nodes.Count,
            },
            TrackBar tb => new
            {
                value = tb.Value,
                minimum = tb.Minimum,
                maximum = tb.Maximum,
            },
            WinProgressBar pb => new
            {
                value = pb.Value,
                minimum = pb.Minimum,
                maximum = pb.Maximum,
                percentage = pb.Maximum > 0 ? ((double)pb.Value / pb.Maximum * 100) : 0,
            },
            NumericUpDown nud => new
            {
                value = nud.Value,
                minimum = nud.Minimum,
                maximum = nud.Maximum,
            },
            _ => null,
        };
    }

    // DataGrid operations

    public string? GetDataGridCell(int row, int col)
    {
        if (_control is not DataGridView dgv) return null;
        if (row < 0 || row >= dgv.RowCount || col < 0 || col >= dgv.ColumnCount) return null;
        return dgv.Rows[row].Cells[col].Value?.ToString();
    }

    public object?[][] GetDataGridAllData()
    {
        if (_control is not DataGridView dgv) return Array.Empty<object?[]>();
        var result = new List<object?[]>();
        for (int r = 0; r < dgv.RowCount; r++)
        {
            var row = new object?[dgv.ColumnCount];
            for (int c = 0; c < dgv.ColumnCount; c++)
                row[c] = dgv.Rows[r].Cells[c].Value?.ToString();
            result.Add(row);
        }
        return result.ToArray();
    }

    public void SelectDataGridRow(int index)
    {
        if (_control is DataGridView dgv && index >= 0 && index < dgv.RowCount)
        {
            dgv.ClearSelection();
            dgv.Rows[index].Selected = true;
            dgv.CurrentCell = dgv.Rows[index].Cells[0];
        }
    }

    // Tree operations

    public void ExpandTreeNode(string path, bool open = true)
    {
        if (_control is not WinTreeView tv) return;
        var parts = path.Split('/');
        var nodes = tv.Nodes;
        TreeNode? target = null;
        foreach (var part in parts)
        {
            TreeNode? found = null;
            foreach (TreeNode node in nodes)
            {
                if (node.Text == part)
                {
                    found = node;
                    break;
                }
            }
            if (found == null) return;
            // Intermediate nodes always expand so we can descend; the final
            // target's state is set by `open` below.
            found.Expand();
            nodes = found.Nodes;
            target = found;
        }
        if (target != null)
        {
            if (open) target.Expand(); else target.Collapse();
        }
    }

    // Scroll

    public void ScrollTo(double offset)
    {
        if (_control is DataGridView dgv)
        {
            var row = Math.Max(0, Math.Min((int)offset, dgv.RowCount - 1));
            if (row < dgv.RowCount)
                dgv.FirstDisplayedScrollingRowIndex = row;
        }
        else if (_control is WinListBox lb)
        {
            lb.TopIndex = Math.Max(0, Math.Min((int)offset, lb.Items.Count - 1));
        }
        // For other controls, scroll is not applicable
    }

    // Drag-and-drop (WPF-only; not supported for WinForms)

    public void DragTo(System.Windows.FrameworkElement target)
    {
        throw new NotSupportedException("DragTo is not supported for WinForms controls.");
    }

    // Context menu

    public void OpenContextMenu()
    {
        if (_control.ContextMenuStrip != null)
        {
            _control.ContextMenuStrip.Show(_control, new System.Drawing.Point(_control.Width / 2, _control.Height / 2));
        }
        else
        {
            throw new InvalidOperationException($"Control '{_control.Name}' has no ContextMenuStrip.");
        }
    }

    public string[] GetContextMenuItems()
    {
        if (_control.ContextMenuStrip != null)
        {
            return _control.ContextMenuStrip.Items
                .Cast<ToolStripItem>()
                .Select(i => i.Text ?? "")
                .ToArray();
        }
        return Array.Empty<string>();
    }

    public void ClickContextMenuItem(string name)
    {
        if (_control.ContextMenuStrip == null)
            throw new InvalidOperationException($"Control '{_control.Name}' has no ContextMenuStrip.");

        foreach (ToolStripItem item in _control.ContextMenuStrip.Items)
        {
            if (string.Equals(item.Text, name, StringComparison.OrdinalIgnoreCase))
            {
                item.PerformClick();
                return;
            }
        }
        throw new ArgumentException($"Context menu item not found: {name}");
    }

    // Private helpers

    private static void SelectTabByName(WinTabControl tc, string name)
    {
        for (int i = 0; i < tc.TabCount; i++)
        {
            if (tc.TabPages[i].Text == name)
            {
                tc.SelectedIndex = i;
                return;
            }
        }
        throw new ArgumentException($"Tab not found: {name}");
    }

    private static string SelectInCollection(
        dynamic items, Action<int> setIndex, int count, int? index, string? text)
    {
        if (index.HasValue)
        {
            if (index.Value < 0 || index.Value >= count)
                throw new ArgumentOutOfRangeException($"index {index.Value} out of range (0..{count - 1})");
            setIndex(index.Value);
            return items[index.Value]?.ToString() ?? "";
        }

        if (text == null)
            throw new ArgumentException("pane.list.select requires 'index' (int) or 'text' (string)");

        for (int i = 0; i < count; i++)
        {
            var itemStr = items[i]?.ToString() ?? "";
            if (itemStr.Contains(text, StringComparison.OrdinalIgnoreCase))
            {
                setIndex(i);
                return itemStr;
            }
        }

        throw new ArgumentException(
            $"No item matching '{text}' found in {count}-item list. " +
            $"Use index (0..{count - 1}) for direct selection.");
    }
}
