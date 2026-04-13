using System.Diagnostics;
using System.Text.Json.Nodes;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using XRai.Core;

namespace XRai.UI;

/// <summary>
/// FlaUI-backed ribbon + dialog driver. All ribbon button operations are
/// SCOPED to the active ribbon tab (not window-wide) and use EXACT name
/// matching. This prevents accidentally clicking a button on a different
/// tab that shares a label, or triggering an arbitrary window element that
/// happens to contain the search text (e.g. the Excel "Add-ins" dialog).
/// </summary>
public class RibbonDriver
{
    private UIA3Automation? _automation;
    private Application? _app;
    private Window? _window;
    private readonly Win32DialogDriver? _win32Fallback;

    /// <summary>
    /// Default constructor kept for backwards-compat and reflective instantiation.
    /// Without a Win32 fallback, dialog.click / dialog.dismiss can only see
    /// UIA-visible modal children of Excel's main window.
    /// </summary>
    public RibbonDriver() { }

    /// <summary>
    /// Preferred constructor. The injected Win32DialogDriver is used as a
    /// fallback for dialog.click / dialog.dismiss when the UIA ModalWindows
    /// enumeration finds nothing — this happens for top-level Win32 dialogs
    /// that are not parented to XLMAIN (standard #32770 dialogs, NUIDialog
    /// windows, OLE server-busy message boxes). Without this fallback, agents
    /// see a desync: win32.dialog.list reports a live dialog while dialog.click
    /// reports "No dialog is open".
    /// </summary>
    public RibbonDriver(Win32DialogDriver win32Fallback)
    {
        _win32Fallback = win32Fallback;
    }

    public void Register(CommandRouter router)
    {
        router.RegisterNoSta("ribbon", HandleRibbon);
        router.RegisterNoSta("ribbon.tabs", HandleRibbon);                  // alias
        router.RegisterNoSta("ribbon.buttons", HandleRibbonButtons);         // tab-scoped; no tab = all tabs
        router.RegisterNoSta("ribbon.buttons.all", HandleRibbonButtonsAll);  // explicit all-tabs
        router.RegisterNoSta("ribbon.activate", HandleRibbonActivate);
        router.RegisterNoSta("ribbon.tab.activate", HandleRibbonActivate);   // alias (requested by CellVault session)
        router.RegisterNoSta("ribbon.click", HandleRibbonClick);
        router.RegisterNoSta("dialog.read", HandleDialogRead);
        router.RegisterNoSta("dialog.click", HandleDialogClick);
        router.RegisterNoSta("dialog.dismiss", HandleDialogDismiss);
        router.RegisterNoSta("ui.tree", HandleUiTree);
    }

    // ── Attach ───────────────────────────────────────────────────────

    private void EnsureAttached()
    {
        if (_window != null) return;

        var procs = Process.GetProcessesByName("EXCEL");
        if (procs.Length == 0)
            throw new InvalidOperationException("No Excel process found");

        _automation = new UIA3Automation();
        _app = Application.Attach(procs[0]);
        _window = _app.GetMainWindow(_automation, TimeSpan.FromSeconds(5));
    }

    /// <summary>
    /// Clear cached UIA state. Call after Excel restart or when the user
    /// signals the current cache is stale. Next operation will re-walk.
    /// </summary>
    public void InvalidateCache()
    {
        try { _automation?.Dispose(); } catch { }
        _automation = null;
        _app = null;
        _window = null;
    }

    private static string SafeGet(Func<string?> getter)
    {
        try { return getter() ?? ""; } catch { return ""; }
    }

    // ── Tab operations ───────────────────────────────────────────────

    private AutomationElement[] GetAllTabs()
    {
        // On Windows 11 build 26200 + Office 365, the ribbon's UIA tree
        // can be lazily populated. If the first walk returns 0 tabs:
        //   1. Focus the window (BringToFront)
        //   2. Send Alt to force the ribbon into keytip mode (expands it
        //      if minimized) then Escape to dismiss keytips
        //   3. Retry the walk up to 3 times with 300ms delays
        // This covers the "ribbon.tabs returns empty" scenario reported
        // on Win11 26200.
        for (int attempt = 0; attempt < 4; attempt++)
        {
            var tabs = _window!.FindAllDescendants(cf => cf.ByControlType(ControlType.TabItem));
            if (tabs.Length > 0) return tabs;

            if (attempt == 0)
            {
                // First retry: focus the window
                try { _window.SetForeground(); } catch { }
                Thread.Sleep(200);
            }
            else if (attempt == 1)
            {
                // Second retry: press and release Alt to trigger ribbon
                // keytip expansion (forces UIA tree population)
                try
                {
                    FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.ALT);
                    Thread.Sleep(100);
                    FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.ESCAPE);
                    Thread.Sleep(300);
                }
                catch { }
            }
            else if (attempt == 2)
            {
                // Third retry: invalidate the cached UIA tree and
                // reattach to get a fresh automation session
                try
                {
                    _automation?.Dispose();
                    _automation = new UIA3Automation();
                    var procs = Process.GetProcessesByName("EXCEL");
                    if (procs.Length > 0)
                    {
                        _app = Application.Attach(procs[0]);
                        _window = _app.GetMainWindow(_automation, TimeSpan.FromSeconds(3));
                        foreach (var p in procs) try { p.Dispose(); } catch { }
                    }
                }
                catch { }
                Thread.Sleep(300);
            }
            else
            {
                Thread.Sleep(300);
            }
        }

        // All retries exhausted — return whatever we got (may be empty)
        return _window!.FindAllDescendants(cf => cf.ByControlType(ControlType.TabItem));
    }

    private AutomationElement? GetActiveTab()
    {
        var tabs = GetAllTabs();
        foreach (var tab in tabs)
        {
            try
            {
                var sel = tab.Patterns.SelectionItem.PatternOrDefault;
                if (sel != null && sel.IsSelected.Value)
                    return tab;
            }
            catch { }
        }
        return null;
    }

    private AutomationElement? FindTabByName(string name)
    {
        var tabs = GetAllTabs();
        foreach (var tab in tabs)
        {
            if (string.Equals(SafeGet(() => tab.Name), name, StringComparison.OrdinalIgnoreCase))
                return tab;
        }
        return null;
    }

    private bool ActivateTab(AutomationElement tab)
    {
        try
        {
            tab.Patterns.SelectionItem.PatternOrDefault?.Select();
            Thread.Sleep(200); // let the ribbon render
            return true;
        }
        catch { return false; }
    }

    // ── Ribbon enumeration ───────────────────────────────────────────

    private string HandleRibbon(JsonObject args)
    {
        EnsureAttached();
        var tabs = GetAllTabs();

        var result = new JsonArray();
        foreach (var tab in tabs)
        {
            var name = SafeGet(() => tab.Name);
            if (string.IsNullOrWhiteSpace(name)) continue;

            bool selected = false;
            try { selected = tab.Patterns.SelectionItem.PatternOrDefault?.IsSelected.Value ?? false; }
            catch { }

            result.Add(new JsonObject
            {
                ["name"] = name,
                ["automation_id"] = SafeGet(() => tab.AutomationId),
                ["selected"] = selected,
            });
        }

        return Response.Ok(new { tabs = result, count = result.Count });
    }

    /// <summary>
    /// Enumerate buttons. If a 'tab' arg is provided, activate that tab first and
    /// return its scoped buttons. If no tab is provided, walk ALL tabs, collect
    /// every button, and restore the original active tab at the end.
    /// </summary>
    private string HandleRibbonButtons(JsonObject args)
    {
        EnsureAttached();
        var tabFilter = args["tab"]?.GetValue<string>();

        if (tabFilter == null)
            return EnumerateAllTabButtons();

        // Single-tab mode
        var targetTab = FindTabByName(tabFilter);
        if (targetTab == null)
            return Response.Error($"Ribbon tab not found: {tabFilter}");

        var originalTab = GetActiveTab();
        bool switched = !string.Equals(
            SafeGet(() => originalTab?.Name), tabFilter, StringComparison.OrdinalIgnoreCase);

        if (switched) ActivateTab(targetTab);

        var buttons = ExtractButtonsFromActiveTab(tabFilter);

        if (switched && originalTab != null) ActivateTab(originalTab);

        return Response.Ok(new { buttons, count = buttons.Count, tab = tabFilter });
    }

    private string HandleRibbonButtonsAll(JsonObject args)
    {
        EnsureAttached();
        return EnumerateAllTabButtons();
    }

    private string EnumerateAllTabButtons()
    {
        var originalTab = GetActiveTab();
        var originalName = SafeGet(() => originalTab?.Name);

        var allButtons = new JsonArray();
        var tabs = GetAllTabs();
        var seenIds = new HashSet<string>();

        foreach (var tab in tabs)
        {
            var tabName = SafeGet(() => tab.Name);
            if (string.IsNullOrWhiteSpace(tabName)) continue;

            // Skip sheet tabs (at the bottom), only walk ribbon tabs at the top
            var aid = SafeGet(() => tab.AutomationId);
            if (aid == "SheetTab") continue;

            if (!ActivateTab(tab)) continue;

            var buttons = ExtractButtonsFromActiveTab(tabName);
            foreach (var btn in buttons)
            {
                var id = btn["automation_id"]?.GetValue<string>() ?? "";
                var name = btn["name"]?.GetValue<string>() ?? "";
                var key = $"{tabName}|{id}|{name}";
                if (seenIds.Add(key))
                    allButtons.Add(btn.DeepClone());
            }
        }

        // Restore original tab
        if (originalTab != null && !string.IsNullOrEmpty(originalName))
            ActivateTab(originalTab);

        return Response.Ok(new
        {
            buttons = allButtons,
            count = allButtons.Count,
            restored_tab = originalName
        });
    }

    /// <summary>
    /// Extract buttons ONLY from the currently-active ribbon tab's content.
    /// This is the key safety fix: we walk the visual tree looking for the
    /// active tab's content panel and enumerate buttons within it, rather
    /// than searching the entire window (which would match dialogs, etc.).
    /// </summary>
    private JsonArray ExtractButtonsFromActiveTab(string tabName)
    {
        var result = new JsonArray();

        // Strategy: find all Button descendants of the window, but filter out
        // the ones that are NOT descendants of the active ribbon content panel.
        // In practice, the active tab content is a sibling of the tab strip,
        // so we find buttons that are visible and whose parent chain doesn't
        // include a sheet tab or dialog window.
        var buttons = _window!.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));

        foreach (var btn in buttons)
        {
            var btnName = SafeGet(() => btn.Name);
            var btnId = SafeGet(() => btn.AutomationId);

            if (string.IsNullOrWhiteSpace(btnName) && string.IsNullOrWhiteSpace(btnId)) continue;

            // Skip window chrome
            if (btnName is "Minimize" or "Maximize" or "Close" or "Minimise" or "Maximise"
                or "Restore" or "Help" or "AutoSave") continue;

            // Skip buttons that are not visible on screen (collapsed/offscreen = different tab)
            try
            {
                if (btn.IsOffscreen) continue;
            }
            catch { }

            result.Add(new JsonObject
            {
                ["tab"] = tabName,
                ["name"] = btnName,
                ["automation_id"] = btnId,
                ["enabled"] = btn.IsEnabled,
            });
        }

        return result;
    }

    private string HandleRibbonActivate(JsonObject args)
    {
        var tabName = args["tab"]?.GetValue<string>()
            ?? throw new ArgumentException("ribbon.activate requires 'tab'");

        EnsureAttached();
        var target = FindTabByName(tabName);
        if (target == null)
            return Response.Error($"Ribbon tab not found: {tabName}");

        if (ActivateTab(target))
            return Response.Ok(new { tab = tabName, activated = true });

        return Response.Error($"Failed to activate tab '{tabName}'");
    }

    // ── Ribbon click (scoped + exact + candidates on failure) ────────

    private string HandleRibbonClick(JsonObject args)
    {
        EnsureAttached();

        var automationId = args["automation_id"]?.GetValue<string>();
        var button = args["button"]?.GetValue<string>();
        var tabName = args["tab"]?.GetValue<string>();

        if (automationId == null && button == null)
            return Response.Error("ribbon.click requires 'automation_id' or 'button' (or both)");

        // Step 1: If a tab is specified, activate it
        if (tabName != null)
        {
            var target = FindTabByName(tabName);
            if (target == null)
                return Response.Error($"Ribbon tab not found: {tabName}");
            if (!ActivateTab(target))
                return Response.Error($"Failed to activate tab '{tabName}'");
        }

        // Step 2: Find the button
        // Priority: (a) automation_id exact, (b) button name exact on active tab only
        AutomationElement? hit = null;
        var scope = _window!;

        if (automationId != null)
        {
            hit = scope.FindFirstDescendant(cf => cf.ByAutomationId(automationId));
            // Require the hit to be a Button/SplitButton — NEVER a TabItem
            if (hit != null && hit.ControlType == ControlType.TabItem)
                hit = null;
        }

        if (hit == null && button != null)
        {
            // Find all buttons with EXACT name match, filter to visible only
            var all = scope.FindAllDescendants(cf =>
                cf.ByControlType(ControlType.Button).And(cf.ByName(button)));

            AutomationElement? visible = null;
            foreach (var el in all)
            {
                try
                {
                    if (el.IsOffscreen) continue;
                    if (el.ControlType == ControlType.TabItem) continue;
                    visible = el;
                    break;
                }
                catch { }
            }
            hit = visible;

            // If still null, check SplitButton type
            if (hit == null)
            {
                var splits = scope.FindAllDescendants(cf =>
                    cf.ByControlType(ControlType.SplitButton).And(cf.ByName(button)));
                foreach (var el in splits)
                {
                    try
                    {
                        if (el.IsOffscreen) continue;
                        hit = el;
                        break;
                    }
                    catch { }
                }
            }
        }

        if (hit == null)
            return FormatClickFailure(automationId, button, tabName);

        // Step 3: Click using Invoke pattern (preferred) or mouse click fallback
        try
        {
            var invoke = hit.Patterns.Invoke.PatternOrDefault;
            if (invoke != null)
            {
                invoke.Invoke();
                return Response.Ok(new
                {
                    clicked = true,
                    automation_id = SafeGet(() => hit.AutomationId),
                    name = SafeGet(() => hit.Name),
                    tab = tabName,
                    method = "Invoke"
                });
            }

            hit.Click();
            return Response.Ok(new
            {
                clicked = true,
                automation_id = SafeGet(() => hit.AutomationId),
                name = SafeGet(() => hit.Name),
                tab = tabName,
                method = "Click"
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"Failed to click '{SafeGet(() => hit.Name)}': {ex.Message}");
        }
    }

    /// <summary>
    /// When ribbon.click fails, return a helpful error with candidate matches
    /// so the agent can pick the right one without another discovery call.
    /// </summary>
    private string FormatClickFailure(string? automationId, string? button, string? tabName)
    {
        var searched = automationId ?? button ?? "<none>";
        var candidates = new JsonArray();

        if (button != null)
        {
            // Gather all buttons across the window with matching name for diagnosis
            try
            {
                var all = _window!.FindAllDescendants(cf =>
                    cf.ByControlType(ControlType.Button).And(cf.ByName(button)));

                foreach (var el in all.Take(20))
                {
                    candidates.Add(new JsonObject
                    {
                        ["name"] = SafeGet(() => el.Name),
                        ["automation_id"] = SafeGet(() => el.AutomationId),
                        ["offscreen"] = SafeGet(() => el.IsOffscreen.ToString()),
                    });
                }
            }
            catch { }
        }

        var errorObj = new JsonObject
        {
            ["ok"] = false,
            ["error"] = $"Ribbon button not found: {searched}" +
                (tabName != null ? $" (on tab '{tabName}')" : ""),
            ["suggestion"] = "Use {\"cmd\":\"ribbon.buttons\",\"tab\":\"<tab>\"} to list buttons on a specific tab with their automation_ids, or {\"cmd\":\"ribbon.buttons.all\"} to walk every tab.",
        };
        if (candidates.Count > 0)
            errorObj["candidates"] = candidates;

        return errorObj.ToJsonString();
    }

    // ── Dialog (COM/UIA modal) ───────────────────────────────────────

    private string HandleDialogRead(JsonObject args)
    {
        EnsureAttached();
        var windows = _window!.ModalWindows;
        if (windows.Length == 0)
            return Response.Ok(new { dialog = false });

        var dialog = windows[0];
        var buttons = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
        var btnNames = new JsonArray();
        foreach (var btn in buttons)
        {
            var name = SafeGet(() => btn.Name);
            if (!string.IsNullOrWhiteSpace(name)) btnNames.Add(name);
        }

        var texts = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Text));
        var txtContent = new JsonArray();
        foreach (var txt in texts)
        {
            var name = SafeGet(() => txt.Name);
            if (!string.IsNullOrWhiteSpace(name)) txtContent.Add(name);
        }

        return Response.Ok(new
        {
            dialog = true,
            title = SafeGet(() => dialog.Title),
            text = txtContent,
            buttons = btnNames,
        });
    }

    private string HandleDialogClick(JsonObject args)
    {
        var button = args["button"]?.GetValue<string>()
            ?? throw new ArgumentException("dialog.click requires 'button'");

        // Tier 1: UIA — catches modals that are children of Excel's main window.
        try
        {
            EnsureAttached();
            var windows = _window!.ModalWindows;
            if (windows.Length > 0)
            {
                var dialog = windows[0];
                var buttons = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
                foreach (var btn in buttons)
                {
                    if (string.Equals(SafeGet(() => btn.Name), button, StringComparison.OrdinalIgnoreCase))
                    {
                        try { btn.Patterns.Invoke.PatternOrDefault?.Invoke(); }
                        catch { btn.Click(); }
                        return Response.Ok(new { button, clicked = true, source = "uia" });
                    }
                }
                // UIA saw a modal but no matching button — still fall through to Win32
                // in case there's ANOTHER dialog visible only to Win32 enumeration.
            }
        }
        catch { /* UIA path failed — fall through */ }

        // Tier 2: Win32 EnumWindows fallback. Catches top-level #32770 dialogs,
        // NUIDialog windows, and OLE server-busy message boxes that are NOT
        // parented to Excel's main window and therefore invisible to UIA.
        // This closes the dialog.click ↔ win32.dialog.list desync.
        if (_win32Fallback != null)
        {
            var win32Args = new JsonObject { ["button"] = button };
            var win32Result = _win32Fallback.TryClickButton(win32Args);
            if (win32Result != null) return win32Result;
        }

        return Response.Error($"No dialog is open (checked both UIA and Win32 top-level enumeration). " +
            "If you know a dialog is visible, run {\"cmd\":\"win32.dialog.list\"} for full diagnostics.");
    }

    private string HandleDialogDismiss(JsonObject args)
    {
        string? lastTitle = null;

        // Tier 1: UIA.
        try
        {
            EnsureAttached();
            var windows = _window!.ModalWindows;
            if (windows.Length > 0)
            {
                var dialog = windows[0];
                lastTitle = SafeGet(() => dialog.Title);
                var priorities = new[] { "OK", "Yes", "Continue", "Retry", "Close", "Cancel", "No" };

                foreach (var preferred in priorities)
                {
                    var buttons = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
                    foreach (var btn in buttons)
                    {
                        if (string.Equals(SafeGet(() => btn.Name), preferred, StringComparison.OrdinalIgnoreCase))
                        {
                            try
                            {
                                btn.Patterns.Invoke.PatternOrDefault?.Invoke();
                                return Response.Ok(new { dismissed = true, title = lastTitle, clicked = preferred, source = "uia" });
                            }
                            catch { }
                        }
                    }
                }
                // UIA modal present but no known safe button — fall through to Win32.
            }
        }
        catch { /* fall through */ }

        // Tier 2: Win32 fallback — same one-shot sweep excel.autodismiss uses.
        if (_win32Fallback != null)
        {
            var count = _win32Fallback.DismissOnce();
            if (count > 0)
                return Response.Ok(new { dismissed = true, count, source = "win32" });
        }

        return Response.Ok(new
        {
            dismissed = false,
            reason = lastTitle != null
                ? $"UIA dialog '{lastTitle}' had no recognized button, and Win32 fallback found nothing"
                : "no dialog found via UIA or Win32 enumeration",
        });
    }

    // ── UI tree dump ─────────────────────────────────────────────────

    private string HandleUiTree(JsonObject args)
    {
        var depth = args["depth"]?.GetValue<int>() ?? 3;
        EnsureAttached();
        return Response.Ok(new { tree = WalkTree(_window!, depth, 0) });
    }

    private JsonObject WalkTree(AutomationElement element, int maxDepth, int currentDepth)
    {
        var node = new JsonObject
        {
            ["type"] = SafeGet(() => element.ControlType.ToString()),
            ["name"] = SafeGet(() => element.Name),
        };

        var aid = SafeGet(() => element.AutomationId);
        if (!string.IsNullOrEmpty(aid)) node["automation_id"] = aid;

        if (currentDepth < maxDepth)
        {
            try
            {
                var children = element.FindAllChildren();
                if (children.Length > 0)
                {
                    var arr = new JsonArray();
                    foreach (var child in children.Take(50))
                        arr.Add(WalkTree(child, maxDepth, currentDepth + 1));
                    node["children"] = arr;
                }
            }
            catch { }
        }

        return node;
    }
}
