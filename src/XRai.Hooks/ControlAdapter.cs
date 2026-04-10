using System.Reflection;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

// Disambiguate WPF vs WinForms types (UseWindowsForms=true in csproj)
using Button = System.Windows.Controls.Button;
using CheckBox = System.Windows.Controls.CheckBox;
using ComboBox = System.Windows.Controls.ComboBox;
using Control = System.Windows.Controls.Control;
using DataGrid = System.Windows.Controls.DataGrid;
using DataObject = System.Windows.DataObject;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using KeyEventHandler = System.Windows.Input.KeyEventHandler;
using Label = System.Windows.Controls.Label;
using ListBox = System.Windows.Controls.ListBox;
using ListView = System.Windows.Controls.ListView;
using Menu = System.Windows.Controls.Menu;
using MenuItem = System.Windows.Controls.MenuItem;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using PasswordBox = System.Windows.Controls.PasswordBox;
using ProgressBar = System.Windows.Controls.ProgressBar;
using RadioButton = System.Windows.Controls.RadioButton;
using RichTextBox = System.Windows.Controls.RichTextBox;
using ScrollViewer = System.Windows.Controls.ScrollViewer;
using Slider = System.Windows.Controls.Slider;
using TabControl = System.Windows.Controls.TabControl;
using TabItem = System.Windows.Controls.TabItem;
using TextBox = System.Windows.Controls.TextBox;
using TreeView = System.Windows.Controls.TreeView;
using TreeViewItem = System.Windows.Controls.TreeViewItem;
using ButtonBase = System.Windows.Controls.Primitives.ButtonBase;
using ToggleButton = System.Windows.Controls.Primitives.ToggleButton;
using Selector = System.Windows.Controls.Primitives.Selector;
using Expander = System.Windows.Controls.Expander;
using DatePicker = System.Windows.Controls.DatePicker;

namespace XRai.Hooks;

public class ControlAdapter : IControlAdapter
{
    private readonly FrameworkElement _element;

    public string Name => _element.Name;
    public string Type { get; }
    public bool IsEnabled => _element.IsEnabled;
    public bool IsVisible => _element.Visibility == Visibility.Visible;
    public FrameworkElement Element => _element;

    /// <summary>
    /// True if this control (or its nearest ButtonBase ancestor) has a non-null Command binding.
    /// Used by pane list output so agents can decide between pane.click vs pane.focus+pane.key.
    /// </summary>
    public bool HasCommand
    {
        get
        {
            try
            {
                var target = FindClickTarget(_element);
                if (target is ButtonBase btn && btn.Command != null) return true;
            }
            catch { }
            return false;
        }
    }

    public ControlAdapter(FrameworkElement element)
    {
        _element = element;
        Type = element switch
        {
            DataGrid => "DataGrid",
            TabControl => "TabControl",
            TreeView => "TreeView",
            ListView => "ListView",
            RadioButton => "RadioButton",
            DatePicker => "DatePicker",
            Slider => "Slider",
            ProgressBar => "ProgressBar",
            RichTextBox => "RichTextBox",
            Expander => "Expander",
            Menu => "Menu",
            TextBox => "TextBox",
            Button => "Button",
            Label => "Label",
            ComboBox => "ComboBox",
            CheckBox => "CheckBox",
            ToggleButton => "ToggleButton",
            ListBox => "ListBox",
            PasswordBox => "PasswordBox",
            ScrollViewer => "ScrollViewer",
            _ => element.GetType().Name,
        };
    }

    public string? GetValue()
    {
        return _element switch
        {
            TextBox tb => tb.Text,
            Label lbl => lbl.Content?.ToString(),
            ComboBox cb => cb.SelectedItem?.ToString() ?? cb.Text,
            RadioButton rb => rb.IsChecked?.ToString(),
            CheckBox chk => chk.IsChecked?.ToString(),
            ToggleButton tb => tb.IsChecked?.ToString(),
            Slider sl => sl.Value.ToString(),
            ProgressBar pb => pb.Value.ToString(),
            DatePicker dp => dp.SelectedDate?.ToString("yyyy-MM-dd"),
            TabControl tc => tc.SelectedIndex.ToString(),
            Expander ex => ex.IsExpanded.ToString(),
            RichTextBox rtb => new System.Windows.Documents.TextRange(rtb.Document.ContentStart, rtb.Document.ContentEnd).Text.TrimEnd(),
            ListView lv => lv.SelectedItem?.ToString(),
            ListBox lb => lb.SelectedItem?.ToString(),
            DataGrid dg => GetDataGridValue(dg),
            _ => _element.ToString(),
        };
    }

    public void SetValue(string value)
    {
        switch (_element)
        {
            case TextBox tb:
                tb.Text = value;
                break;
            case ComboBox cb:
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
            case RadioButton rb:
                rb.IsChecked = bool.TryParse(value, out var rb2) && rb2;
                break;
            case CheckBox chk:
                chk.IsChecked = bool.TryParse(value, out var b) ? b : null;
                break;
            case Slider sl:
                if (double.TryParse(value, out var sv)) sl.Value = sv;
                break;
            case DatePicker dp:
                if (DateTime.TryParse(value, out var dt)) dp.SelectedDate = dt;
                break;
            case TabControl tc:
                if (int.TryParse(value, out var idx) && idx >= 0 && idx < tc.Items.Count)
                    tc.SelectedIndex = idx;
                else
                    SelectTabByName(tc, value);
                break;
            case Expander ex:
                ex.IsExpanded = bool.TryParse(value, out var exp) && exp;
                break;
            case ListView lv:
                SelectItemByText(lv, value);
                break;
            case ListBox lb:
                SelectItemByText(lb, value);
                break;
            case PasswordBox pb:
                pb.Password = value;
                break;
            case RichTextBox rtb:
                rtb.Document.Blocks.Clear();
                rtb.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new System.Windows.Documents.Run(value)));
                break;
            default:
                throw new InvalidOperationException($"Cannot set value on {Type} control");
        }
    }

    /// <summary>
    /// Click the control. Invokes ButtonBase.OnClick exactly once via reflection
    /// for ButtonBase targets (no pre-click Focus/dispatcher pump, no post-click
    /// observation, no retry, no fallback). Returns a ClickResult reporting the
    /// synchronous CanExecute snapshot and the dispatch method — callers verify
    /// actual effect via their own state checks (model, win32.dialog.list, etc.)
    /// rather than trusting any post-click observation we might do here.
    /// </summary>
    public ClickResult Click()
    {
        var result = new ClickResult();
        var target = FindClickTarget(_element);

        if (target is ButtonBase btn)
        {
            result.ResolvedToButtonBase = true;
            result.ResolvedTargetType = btn.GetType().Name;
            result.HasCommand = btn.Command != null;

            // Snapshot CanExecute BEFORE invoking. This is the ONLY thing we can
            // report synchronously about whether the Command ran. Any post-click
            // observation (routed-event observer, SilentNoOp detection, dispatcher
            // flush + state check) is fundamentally unable to distinguish:
            //   - Command ran and returned instantly
            //   - Command scheduled async work and returned (ShowDialog via
            //     Dispatcher.BeginInvoke, Task.Run, await continuations, etc.)
            //   - Command opened a modal via ShowDialog() and is still blocked
            //     in a nested dispatcher frame
            // Any fallback or retry based on post-click "did it actually work"
            // guesses is a double-fire hazard for modal-opening Commands. So
            // we do NONE of that. CanExecute is reported; verification is the
            // caller's responsibility (read `model`, `win32.dialog.list`, etc).
            result.CommandCanExecute = btn.Command?.CanExecute(btn.CommandParameter) ?? false;

            // Invoke ButtonBase.OnClick via reflection. This is exactly what
            // ButtonAutomationPeer calls internally (via AutomationButtonBaseClick)
            // — but SYNCHRONOUSLY, bypassing the peer's Dispatcher.BeginInvoke.
            // Reflection on the ButtonBase MethodInfo honors virtual dispatch
            // automatically, so subclass overrides (ToggleButton.OnToggle,
            // CheckBox/RadioButton state, RepeatButton timing) run correctly
            // without walking the type hierarchy.
            //
            // OnClick is called EXACTLY ONCE. No pre-click Focus() or dispatcher
            // pump (both can run unrelated queued work and alter state).
            // No post-click observer. No fallback Command.Execute. No retry.
            var onClickMethod = typeof(ButtonBase).GetMethod(
                "OnClick",
                BindingFlags.Instance | BindingFlags.NonPublic);

            if (onClickMethod == null)
            {
                // Should be impossible for any supported WPF — ButtonBase.OnClick
                // has existed since .NET 3.0. Surface the failure instead of
                // attempting an alternative code path that might double-fire.
                result.ErrorHint = "ButtonBase.OnClick reflection lookup failed — WPF runtime mismatch.";
                result.Method = "error: OnClick not found";
                return result;
            }

            try
            {
                onClickMethod.Invoke(btn, null);
                result.Method = "ButtonBase.OnClick";
                // Report CommandExecuted from the CanExecute snapshot — the only
                // synchronous signal available. The caller can verify real effect
                // via `model`, `win32.dialog.list`, or whatever state check is
                // appropriate for their Command.
                result.CommandExecuted = result.CommandCanExecute;
            }
            catch (TargetInvocationException tie)
            {
                // OnClick's handler threw (Command body exception, etc.).
                // The click DID happen — OnClick was called, ClickEvent was
                // raised, ExecuteCommandSource was reached. We just didn't
                // make it through the handler. Report as a warning, not a
                // silent no-op, so the caller doesn't retry and double-fire.
                result.ErrorHint = $"OnClick handler threw: {tie.InnerException?.Message ?? tie.Message}";
                result.Method = "ButtonBase.OnClick (handler threw)";
                result.CommandExecuted = result.CommandCanExecute;
            }

            return result;
        }

        // Not a ButtonBase and no ancestor found
        result.ResolvedToButtonBase = false;
        result.ResolvedTargetType = _element.GetType().Name;

        var elementPeer = UIElementAutomationPeer.CreatePeerForElement(_element);
        var elementInvoke = elementPeer?.GetPattern(PatternInterface.Invoke) as IInvokeProvider;
        if (elementInvoke != null)
        {
            elementInvoke.Invoke();
            result.Method = "IInvokeProvider.Invoke (non-button)";
            return result;
        }

        // Last resort: raw mouse events on the element
        _element.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
        {
            RoutedEvent = UIElement.PreviewMouseLeftButtonDownEvent
        });
        _element.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
        {
            RoutedEvent = UIElement.MouseLeftButtonUpEvent
        });
        result.Method = "MouseEvents (fallback)";
        result.ErrorHint = "No ButtonBase ancestor found and no InvokePattern available. " +
            "Raw mouse events were synthesized but may not trigger Command bindings.";
        return result;
    }

    public class ClickResult
    {
        public string? Method { get; set; }
        public bool ResolvedToButtonBase { get; set; }
        public string? ResolvedTargetType { get; set; }
        public bool HasCommand { get; set; }
        public bool CommandCanExecute { get; set; }
        public bool CommandExecuted { get; set; }
        public string? ErrorHint { get; set; }
    }

    /// <summary>
    /// Walks UP the visual tree from a named element looking for the nearest
    /// ButtonBase ancestor. This handles the common ControlTemplate pattern where
    /// developers put x:Name on a visual child (Border, Grid, ContentPresenter)
    /// inside a button's template. Without this walk, pane.click would hit the
    /// template child's routed event handler but never reach the outer Button's
    /// OnClick — and therefore never execute Command bindings.
    ///
    /// Returns the element itself if it IS a ButtonBase, otherwise the nearest
    /// ButtonBase ancestor, or the original element if no ancestor qualifies.
    /// </summary>
    private static DependencyObject FindClickTarget(FrameworkElement start)
    {
        if (start is ButtonBase) return start;

        DependencyObject? current = start;
        int maxDepth = 20; // sanity limit — templates rarely nest this deep
        while (current != null && maxDepth-- > 0)
        {
            if (current is ButtonBase) return current;
            current = VisualTreeHelper.GetParent(current);
        }
        return start;
    }

    public void DoubleClick()
    {
        _element.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
        {
            RoutedEvent = Control.MouseDoubleClickEvent
        });
    }

    public void RightClick()
    {
        _element.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Right)
        {
            RoutedEvent = UIElement.MouseRightButtonDownEvent
        });
        _element.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Right)
        {
            RoutedEvent = UIElement.MouseRightButtonUpEvent
        });
    }

    public void Hover()
    {
        _element.RaiseEvent(new MouseEventArgs(Mouse.PrimaryDevice, 0)
        {
            RoutedEvent = UIElement.MouseEnterEvent
        });
    }

    public void Focus()
    {
        _element.Focus();
        Keyboard.Focus(_element);
    }

    /// <summary>
    /// Sends key events to the focused element. Returns a SendKeysResult indicating
    /// which keys were delivered, whether the element had a presentation source,
    /// and any post-delivery exception (captured but not rethrown — a keystroke
    /// that triggers a modal dialog may invalidate the presentation source AFTER
    /// the key has already been processed, and that's a success, not a failure).
    /// </summary>
    public SendKeysResult SendKeys(string keys)
    {
        var result = new SendKeysResult { InputKeys = keys };

        try { Focus(); }
        catch (Exception ex) { result.FocusWarning = ex.Message; }

        var source = PresentationSource.FromVisual(_element);
        if (source == null)
        {
            result.NoPresentationSource = true;
            result.ErrorHint = "Element has no PresentationSource — it may not be in a visible window. " +
                "Try {\"cmd\":\"pane.focus\",\"control\":\"...\"} first, or ensure the containing pane is visible.";
            return result;
        }

        foreach (var keyStr in keys.Split('+'))
        {
            var trimmed = keyStr.Trim();
            if (!Enum.TryParse<Key>(trimmed, true, out var key))
            {
                result.UnknownKeys.Add(trimmed);
                continue;
            }

            // Register a peek handler that fires on any KeyDown — including after
            // the event has been marked Handled. If ANY handler (including the
            // Button's OnKeyDown which fires Click, which opens the modal, which
            // invalidates the presentation source) sees the event, the keystroke
            // was delivered. That's the ground truth for "did the key actually
            // reach its target", independent of any exception during routing.
            bool keyDownSawHandler = false;
            KeyEventHandler peek = (_, _) => keyDownSawHandler = true;
            _element.AddHandler(UIElement.KeyDownEvent, peek, handledEventsToo: true);

            Exception? routingException = null;
            try
            {
                _element.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice, source, 0, key)
                {
                    RoutedEvent = UIElement.KeyDownEvent
                });
            }
            catch (Exception ex)
            {
                // RaiseEvent threw — but a handler may have still run before the throw.
                // The peek will tell us whether the event was dispatched.
                routingException = ex;
            }
            finally
            {
                try { _element.RemoveHandler(UIElement.KeyDownEvent, peek); } catch { }
            }

            if (keyDownSawHandler)
            {
                // Event was delivered to at least one handler. Command fires in the
                // Button's OnKeyDown → OnClick pathway. This is success.
                result.DeliveredKeys.Add(trimmed);
                if (routingException != null)
                {
                    result.PostDeliveryWarnings.Add(
                        $"KeyDown for '{trimmed}' dispatched to handlers but routing threw post-handler " +
                        $"(common when the key opens a modal dialog): {routingException.Message}");
                }
            }
            else
            {
                // No handler ever saw the event. This is a real failure.
                if (routingException != null)
                    result.FailedKeys.Add($"{trimmed}: {routingException.Message}");
                else
                    result.FailedKeys.Add($"{trimmed}: event raised but no handlers processed it");
                continue;
            }

            // KeyUp is best-effort — many handlers only care about KeyDown, and by
            // the time we try to send KeyUp, a modal may have already invalidated
            // the presentation source. Any exception here is a warning, not a failure.
            try
            {
                var currentSource = PresentationSource.FromVisual(_element);
                if (currentSource != null)
                {
                    _element.RaiseEvent(new KeyEventArgs(Keyboard.PrimaryDevice, currentSource, 0, key)
                    {
                        RoutedEvent = UIElement.KeyUpEvent
                    });
                }
            }
            catch (Exception upEx)
            {
                result.PostDeliveryWarnings.Add($"KeyUp for '{trimmed}' failed but KeyDown was delivered: {upEx.Message}");
            }
        }

        return result;
    }

    public class SendKeysResult
    {
        public string InputKeys { get; set; } = "";
        public List<string> DeliveredKeys { get; } = new();
        public List<string> FailedKeys { get; } = new();
        public List<string> UnknownKeys { get; } = new();
        public List<string> PostDeliveryWarnings { get; } = new();
        public bool NoPresentationSource { get; set; }
        public string? FocusWarning { get; set; }
        public string? ErrorHint { get; set; }

        public bool AnyDelivered => DeliveredKeys.Count > 0;
        public bool AllDelivered => FailedKeys.Count == 0 && UnknownKeys.Count == 0 && DeliveredKeys.Count > 0;
    }

    public void Toggle()
    {
        switch (_element)
        {
            case ToggleButton tb:
                tb.IsChecked = !tb.IsChecked;
                break;
            case Expander ex:
                ex.IsExpanded = !ex.IsExpanded;
                break;
            default:
                throw new InvalidOperationException($"Cannot toggle {Type} control");
        }
    }

    /// <summary>
    /// Programmatically open/expand a ComboBox dropdown, Expander, or TreeViewItem.
    /// Allows agents to screenshot dropdown styling, inspect dropdown items, or
    /// verify visual state without simulating mouse clicks.
    /// </summary>
    public string Expand(bool open = true)
    {
        switch (_element)
        {
            case ComboBox cb:
                cb.IsDropDownOpen = open;
                return open ? "ComboBox dropdown opened" : "ComboBox dropdown closed";
            case Expander ex:
                ex.IsExpanded = open;
                return open ? "Expander expanded" : "Expander collapsed";
            case TreeViewItem tvi:
                tvi.IsExpanded = open;
                return open ? "TreeViewItem expanded" : "TreeViewItem collapsed";
            case MenuItem mi when mi.HasItems:
                mi.IsSubmenuOpen = open;
                return open ? "MenuItem submenu opened" : "MenuItem submenu closed";
            default:
                throw new InvalidOperationException(
                    $"pane.expand not supported on {_element.GetType().Name}. " +
                    "Supported: ComboBox, Expander, TreeViewItem, MenuItem.");
        }
    }

    public void ScrollTo(double offset)
    {
        if (_element is ScrollViewer sv)
        {
            sv.ScrollToVerticalOffset(offset);
        }
        else
        {
            // Try to find a ScrollViewer inside the control
            var child = FindChild<ScrollViewer>(_element);
            child?.ScrollToVerticalOffset(offset);
        }
    }

    public object? GetDetailedInfo()
    {
        return _element switch
        {
            DataGrid dg => new
            {
                rows = dg.Items.Count,
                columns = dg.Columns.Count,
                selected_index = dg.SelectedIndex,
                column_headers = dg.Columns.Select(c => c.Header?.ToString()).ToArray(),
            },
            TabControl tc => new
            {
                selected_index = tc.SelectedIndex,
                tab_count = tc.Items.Count,
                tabs = tc.Items.Cast<object>().Select((item, i) =>
                {
                    if (item is TabItem ti) return ti.Header?.ToString() ?? $"Tab {i}";
                    return item.ToString();
                }).ToArray(),
            },
            ListView lv => new
            {
                item_count = lv.Items.Count,
                selected_index = lv.SelectedIndex,
            },
            ListBox lb => new
            {
                item_count = lb.Items.Count,
                selected_index = lb.SelectedIndex,
                items = lb.Items.Cast<object>().Take(50).Select(i => i.ToString()).ToArray(),
            },
            TreeView tv => new
            {
                item_count = tv.Items.Count,
            },
            Slider sl => new
            {
                value = sl.Value,
                minimum = sl.Minimum,
                maximum = sl.Maximum,
            },
            ProgressBar pb => new
            {
                value = pb.Value,
                minimum = pb.Minimum,
                maximum = pb.Maximum,
                percentage = pb.Maximum > 0 ? (pb.Value / pb.Maximum * 100) : 0,
            },
            ComboBox cb => new
            {
                selected_index = cb.SelectedIndex,
                item_count = cb.Items.Count,
                items = cb.Items.Cast<object>().Take(50).Select(i => i.ToString()).ToArray(),
            },
            _ => null,
        };
    }

    // DataGrid helpers
    public string? GetDataGridCell(int row, int col)
    {
        if (_element is not DataGrid dg) return null;
        if (row < 0 || row >= dg.Items.Count || col < 0 || col >= dg.Columns.Count) return null;

        var item = dg.Items[row];
        var column = dg.Columns[col];
        var binding = (column as DataGridBoundColumn)?.Binding as System.Windows.Data.Binding;
        if (binding?.Path?.Path != null && item != null)
        {
            var prop = item.GetType().GetProperty(binding.Path.Path);
            return prop?.GetValue(item)?.ToString();
        }
        return item?.ToString();
    }

    public object?[][] GetDataGridAllData()
    {
        if (_element is not DataGrid dg) return Array.Empty<object?[]>();
        var result = new List<object?[]>();

        for (int r = 0; r < dg.Items.Count; r++)
        {
            var row = new object?[dg.Columns.Count];
            for (int c = 0; c < dg.Columns.Count; c++)
            {
                row[c] = GetDataGridCell(r, c);
            }
            result.Add(row);
        }
        return result.ToArray();
    }

    public void SelectDataGridRow(int index)
    {
        if (_element is DataGrid dg && index >= 0 && index < dg.Items.Count)
        {
            dg.SelectedIndex = index;
            dg.ScrollIntoView(dg.Items[index]);
        }
    }

    // Tree helpers
    public void ExpandTreeNode(string path)
    {
        if (_element is not TreeView tv) return;
        var parts = path.Split('/');
        ItemsControl current = tv;

        foreach (var part in parts)
        {
            for (int i = 0; i < current.Items.Count; i++)
            {
                var container = current.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                if (container != null && (container.Header?.ToString() == part))
                {
                    container.IsExpanded = true;
                    container.UpdateLayout();
                    current = container;
                    break;
                }
            }
        }
    }

    // Private helpers
    private static string? GetDataGridValue(DataGrid dg)
    {
        return $"DataGrid[{dg.Items.Count} rows x {dg.Columns.Count} cols]";
    }

    private static void SelectTabByName(System.Windows.Controls.TabControl tc, string name)
    {
        for (int i = 0; i < tc.Items.Count; i++)
        {
            if (tc.Items[i] is TabItem ti && ti.Header?.ToString() == name)
            {
                tc.SelectedIndex = i;
                return;
            }
        }
        throw new ArgumentException($"Tab not found: {name}");
    }

    private static void SelectItemByText(Selector selector, string text)
    {
        for (int i = 0; i < selector.Items.Count; i++)
        {
            if (selector.Items[i]?.ToString() == text)
            {
                selector.SelectedIndex = i;
                return;
            }
        }
    }

    /// <summary>
    /// Read all items from a ListBox/ListView/ComboBox as display text.
    /// Returns an array of rendered text for each item, plus the selected index.
    /// </summary>
    public (string[] Items, int SelectedIndex) ReadListItems()
    {
        if (_element is not Selector selector)
            throw new InvalidOperationException(
                $"pane.list.read requires a Selector (ListBox, ListView, ComboBox). Got: {_element.GetType().Name}");

        selector.UpdateLayout();
        var items = new string[selector.Items.Count];
        for (int i = 0; i < selector.Items.Count; i++)
        {
            items[i] = GetRenderedText(selector, i)
                       ?? selector.Items[i]?.ToString()
                       ?? "";
        }
        return (items, selector.SelectedIndex);
    }

    /// <summary>
    /// Select an item in a ListBox/ListView/ComboBox by index or text match.
    /// Supports complex-object items where ToString() matches are unreliable:
    ///   - index (int): direct SelectedIndex set
    ///   - text (string): matches against ToString(), then ContentPresenter
    ///     visual text as fallback for DataTemplate-bound items
    /// Returns the selected item's display text for verification.
    /// </summary>
    public string SelectListItem(int? index, string? text)
    {
        if (_element is not Selector selector)
            throw new InvalidOperationException(
                $"pane.list.select requires a Selector (ListBox, ListView, ComboBox). Got: {_element.GetType().Name}");

        if (index.HasValue)
        {
            if (index.Value < 0 || index.Value >= selector.Items.Count)
                throw new ArgumentOutOfRangeException(
                    $"index {index.Value} out of range (0..{selector.Items.Count - 1})");
            selector.SelectedIndex = index.Value;
            return selector.SelectedItem?.ToString() ?? "";
        }

        if (text == null)
            throw new ArgumentException("pane.list.select requires 'index' (int) or 'text' (string)");

        // Ensure containers are generated so we can read rendered text
        selector.UpdateLayout();

        // Tier 1 (primary): match on rendered display text. This is what the user
        // actually sees — whatever DisplayMemberPath or ItemTemplate produces in the
        // generated container's TextBlock. Handles DataTemplate-bound complex objects
        // where ToString() returns "Namespace.ViewModel" but the display shows "Mike Tokyo".
        for (int i = 0; i < selector.Items.Count; i++)
        {
            var displayText = GetRenderedText(selector, i);
            if (displayText != null && displayText.Contains(text, StringComparison.OrdinalIgnoreCase))
            {
                selector.SelectedIndex = i;
                return displayText;
            }
        }

        // Tier 2: fallback to ToString() — covers items without generated containers
        // (virtualized and off-screen) and simple string/value-type collections.
        for (int i = 0; i < selector.Items.Count; i++)
        {
            var itemStr = selector.Items[i]?.ToString() ?? "";
            if (itemStr.Contains(text, StringComparison.OrdinalIgnoreCase))
            {
                selector.SelectedIndex = i;
                return itemStr;
            }
        }

        throw new ArgumentException(
            $"No item matching '{text}' found in {selector.Items.Count}-item {selector.GetType().Name}. " +
            $"Use index (0..{selector.Items.Count - 1}) for direct selection.");
    }

    /// <summary>
    /// Extract the rendered display text for a Selector item at the given index.
    /// Walks the generated container's visual tree for TextBlocks — this captures
    /// whatever DisplayMemberPath or ItemTemplate produces, regardless of what
    /// the data object's ToString() returns.
    /// </summary>
    private static string? GetRenderedText(Selector selector, int index)
    {
        try
        {
            if (selector.ItemContainerGenerator.ContainerFromIndex(index) is not FrameworkElement container)
                return null;

            // Collect ALL TextBlocks in the container and concatenate their text.
            // Multi-TextBlock templates (e.g. "{Name} - {Location}") produce the
            // full display string this way.
            var texts = new List<string>();
            CollectTextBlocks(container, texts);
            return texts.Count > 0 ? string.Join(" ", texts) : null;
        }
        catch { return null; }
    }

    private static void CollectTextBlocks(DependencyObject parent, List<string> texts)
    {
        int count = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is System.Windows.Controls.TextBlock tb && !string.IsNullOrWhiteSpace(tb.Text))
                texts.Add(tb.Text);
            CollectTextBlocks(child, texts);
        }
    }

    /// <summary>
    /// Synthesize WPF drag events from this control to a target control.
    /// Raises PreviewMouseLeftButtonDown on source, then DragEnter + Drop on target
    /// with a DataObject carrying the source element. This triggers any
    /// Drop handlers on the target without requiring a modal DoDragDrop call.
    /// </summary>
    public void DragTo(FrameworkElement target)
    {
        // Raise PreviewMouseLeftButtonDown on source to signal drag initiation
        _element.RaiseEvent(new MouseButtonEventArgs(Mouse.PrimaryDevice, 0, MouseButton.Left)
        {
            RoutedEvent = UIElement.PreviewMouseLeftButtonDownEvent
        });

        // Build drag data carrying the source element
        var data = new DataObject(DataFormats.Serializable, _element);

        // DragEventArgs has an internal constructor in WPF. Create via reflection.
        var enterArgs = CreateDragEventArgs(data, System.Windows.DragDrop.DragEnterEvent, target);
        target.RaiseEvent(enterArgs);

        var dropArgs = CreateDragEventArgs(data, System.Windows.DragDrop.DropEvent, target);
        target.RaiseEvent(dropArgs);
    }

    /// <summary>
    /// Creates a DragEventArgs via reflection (internal constructor in WPF).
    /// </summary>
    private static DragEventArgs CreateDragEventArgs(IDataObject data, RoutedEvent routedEvent, FrameworkElement target)
    {
        // DragEventArgs(IDataObject data, DragDropKeyStates dragDropKeyStates, DragDropEffects allowedEffects,
        //               DependencyObject target, Point point)
        var ctor = typeof(DragEventArgs).GetConstructors(BindingFlags.Instance | BindingFlags.NonPublic);
        if (ctor.Length == 0)
            throw new InvalidOperationException("Cannot find DragEventArgs internal constructor.");

        // Find the constructor with IDataObject parameter
        var dragCtor = ctor.FirstOrDefault(c => c.GetParameters().Length >= 2);
        if (dragCtor == null)
            throw new InvalidOperationException("Cannot find suitable DragEventArgs constructor.");

        var parameters = dragCtor.GetParameters();
        var args = new object?[parameters.Length];

        // Fill parameters by type matching
        for (int i = 0; i < parameters.Length; i++)
        {
            var pt = parameters[i].ParameterType;
            if (pt == typeof(IDataObject))
                args[i] = data;
            else if (pt == typeof(DragDropKeyStates))
                args[i] = DragDropKeyStates.LeftMouseButton;
            else if (pt == typeof(DragDropEffects))
                args[i] = DragDropEffects.Move | DragDropEffects.Copy;
            else if (pt == typeof(DependencyObject))
                args[i] = target;
            else if (pt == typeof(System.Windows.Point))
                args[i] = new System.Windows.Point(target.ActualWidth / 2, target.ActualHeight / 2);
            else
                args[i] = parameters[i].HasDefaultValue ? parameters[i].DefaultValue : null;
        }

        var result = (DragEventArgs)dragCtor.Invoke(args)!;
        result.RoutedEvent = routedEvent;
        return result;
    }

    /// <summary>
    /// Open the ContextMenu on this control (if one exists).
    /// </summary>
    public void OpenContextMenu()
    {
        if (_element.ContextMenu == null)
            throw new InvalidOperationException($"Control '{Name}' ({Type}) has no ContextMenu.");

        _element.ContextMenu.PlacementTarget = _element;
        _element.ContextMenu.IsOpen = true;
    }

    /// <summary>
    /// Return the header text of all items in this control's ContextMenu.
    /// </summary>
    public string[] GetContextMenuItems()
    {
        if (_element.ContextMenu == null)
            throw new InvalidOperationException($"Control '{Name}' ({Type}) has no ContextMenu.");

        var items = new List<string>();
        foreach (var item in _element.ContextMenu.Items)
        {
            if (item is MenuItem mi)
                items.Add(mi.Header?.ToString() ?? "");
            else if (item is Separator)
                items.Add("---");
            else
                items.Add(item?.ToString() ?? "");
        }
        return items.ToArray();
    }

    /// <summary>
    /// Click a MenuItem in this control's ContextMenu by matching its Header text
    /// (case-insensitive contains). Opens the menu first if not already open.
    /// </summary>
    public void ClickContextMenuItem(string name)
    {
        if (_element.ContextMenu == null)
            throw new InvalidOperationException($"Control '{Name}' ({Type}) has no ContextMenu.");

        // Ensure the context menu is open so items are generated
        _element.ContextMenu.PlacementTarget = _element;
        _element.ContextMenu.IsOpen = true;

        foreach (var item in _element.ContextMenu.Items)
        {
            if (item is MenuItem mi && mi.Header?.ToString()?.Contains(name, StringComparison.OrdinalIgnoreCase) == true)
            {
                mi.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
                _element.ContextMenu.IsOpen = false;
                return;
            }
        }

        _element.ContextMenu.IsOpen = false;
        throw new ArgumentException(
            $"Context menu item '{name}' not found. Available: {string.Join(", ", GetContextMenuItems())}");
    }

    private static T? FindChild<T>(DependencyObject parent) where T : DependencyObject
    {
        int count = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T t) return t;
            var found = FindChild<T>(child);
            if (found != null) return found;
        }
        return null;
    }
}
