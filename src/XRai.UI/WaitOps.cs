using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Nodes;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;
using XRai.Core;

namespace XRai.UI;

/// <summary>
/// General-purpose intelligent wait commands that block until a condition is met.
/// These are used for test automation and synchronization scenarios.
///
/// Commands:
///   wait.element   — wait for a UIA element to appear (by name or automation_id)
///   wait.window    — wait for a window with matching title to appear
///   wait.property  — wait for a UIA element's property to reach a specific value
///   wait.gone      — wait for a window/element to disappear
/// </summary>
public class WaitOps
{
    private UIA3Automation? _automation;

    public void Register(CommandRouter router)
    {
        router.RegisterNoSta("wait.element", HandleWaitElement);
        router.RegisterNoSta("wait.window", HandleWaitWindow);
        router.RegisterNoSta("wait.property", HandleWaitProperty);
        router.RegisterNoSta("wait.gone", HandleWaitGone);
    }

    private UIA3Automation GetAutomation()
    {
        return _automation ??= new UIA3Automation();
    }

    // ── wait.element ────────────────────────────────────────────────

    /// <summary>
    /// Wait for a UIA element to appear within a window.
    /// Args: name or automation_id (at least one required), window (title substring, optional),
    ///       timeout (default 10000ms), poll_ms (default 250ms)
    /// </summary>
    private string HandleWaitElement(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>();
        var automationId = args["automation_id"]?.GetValue<string>();
        var windowTitle = args["window"]?.GetValue<string>();
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 10000;
        var pollMs = args["poll_ms"]?.GetValue<int>() ?? 250;

        if (name == null && automationId == null)
            return Response.Error("wait.element requires 'name' or 'automation_id'");

        var automation = GetAutomation();
        var sw = Stopwatch.StartNew();

        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            try
            {
                var parent = FindParentElement(automation, windowTitle);
                if (parent != null)
                {
                    AutomationElement? element = null;

                    if (automationId != null)
                        element = parent.FindFirstDescendant(cf => cf.ByAutomationId(automationId));

                    if (element == null && name != null)
                        element = parent.FindFirstDescendant(cf => cf.ByName(name));

                    if (element != null)
                    {
                        return Response.Ok(new
                        {
                            found = true,
                            elapsed_ms = sw.ElapsedMilliseconds,
                            element_name = SafeGet(() => element.Name),
                            automation_id = SafeGet(() => element.AutomationId),
                            control_type = SafeGet(() => element.ControlType.ToString()),
                            is_enabled = element.IsEnabled,
                        });
                    }
                }
            }
            catch { /* ignore transient UIA errors */ }

            Thread.Sleep(pollMs);
        }

        var desc = automationId != null ? $"automation_id='{automationId}'" : $"name='{name}'";
        return Response.Error($"Timed out after {timeoutMs}ms waiting for element {desc}");
    }

    // ── wait.window ─────────────────────────────────────────────────

    /// <summary>
    /// Wait for a window with a matching title to appear.
    /// Args: title (required — substring match), timeout (default 10000ms), poll_ms (default 250ms)
    /// </summary>
    private string HandleWaitWindow(JsonObject args)
    {
        var title = args["title"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(title))
            return Response.Error("wait.window requires 'title'");

        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 10000;
        var pollMs = args["poll_ms"]?.GetValue<int>() ?? 250;

        var sw = Stopwatch.StartNew();

        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            var hwnd = FindWindowByTitle(title);
            if (hwnd != nint.Zero)
            {
                var windowTitle = GetWindowTitle(hwnd);
                return Response.Ok(new
                {
                    found = true,
                    elapsed_ms = sw.ElapsedMilliseconds,
                    title = windowTitle,
                    hwnd = hwnd.ToInt64(),
                });
            }

            Thread.Sleep(pollMs);
        }

        return Response.Error($"Timed out after {timeoutMs}ms waiting for window with title containing '{title}'");
    }

    // ── wait.property ───────────────────────────────────────────────

    /// <summary>
    /// Wait for a UIA element's property (Name or Value) to reach a specific value.
    /// Args: name or automation_id (element identifier), property (default "Name"),
    ///       value (expected value), window (optional), timeout, poll_ms
    /// </summary>
    private string HandleWaitProperty(JsonObject args)
    {
        var name = args["name"]?.GetValue<string>();
        var automationId = args["automation_id"]?.GetValue<string>();
        var property = args["property"]?.GetValue<string>() ?? "Name";
        var expectedValue = args["value"]?.GetValue<string>();
        var windowTitle = args["window"]?.GetValue<string>();
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 10000;
        var pollMs = args["poll_ms"]?.GetValue<int>() ?? 250;

        if (name == null && automationId == null)
            return Response.Error("wait.property requires 'name' or 'automation_id'");
        if (expectedValue == null)
            return Response.Error("wait.property requires 'value'");

        var automation = GetAutomation();
        var sw = Stopwatch.StartNew();
        string? lastValue = null;

        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            try
            {
                var parent = FindParentElement(automation, windowTitle);
                if (parent != null)
                {
                    AutomationElement? element = null;

                    if (automationId != null)
                        element = parent.FindFirstDescendant(cf => cf.ByAutomationId(automationId));
                    if (element == null && name != null)
                        element = parent.FindFirstDescendant(cf => cf.ByName(name));

                    if (element != null)
                    {
                        lastValue = property.ToLowerInvariant() switch
                        {
                            "name" => SafeGet(() => element.Name),
                            "value" => SafeGet(() => element.Patterns.Value.PatternOrDefault?.Value?.ToString() ?? ""),
                            "helptext" => SafeGet(() => element.HelpText),
                            _ => SafeGet(() => element.Name),
                        };

                        if (string.Equals(lastValue, expectedValue, StringComparison.OrdinalIgnoreCase))
                        {
                            return Response.Ok(new
                            {
                                matched = true,
                                elapsed_ms = sw.ElapsedMilliseconds,
                                property,
                                value = lastValue,
                            });
                        }
                    }
                }
            }
            catch { }

            Thread.Sleep(pollMs);
        }

        var desc = automationId != null ? $"automation_id='{automationId}'" : $"name='{name}'";
        return Response.Error(
            $"Timed out after {timeoutMs}ms waiting for {desc} property '{property}' " +
            $"to equal '{expectedValue}'. Last seen value: '{lastValue ?? "(element not found)"}'");
    }

    // ── wait.gone ───────────────────────────────────────────────────

    /// <summary>
    /// Wait for a window or element to disappear.
    /// Args: title (window title substring) or name/automation_id (element),
    ///       window (parent window for element mode), timeout, poll_ms
    /// </summary>
    private string HandleWaitGone(JsonObject args)
    {
        var title = args["title"]?.GetValue<string>();
        var name = args["name"]?.GetValue<string>();
        var automationId = args["automation_id"]?.GetValue<string>();
        var windowTitle = args["window"]?.GetValue<string>();
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 10000;
        var pollMs = args["poll_ms"]?.GetValue<int>() ?? 250;

        if (title == null && name == null && automationId == null)
            return Response.Error("wait.gone requires 'title' (window) or 'name'/'automation_id' (element)");

        var automation = GetAutomation();
        var sw = Stopwatch.StartNew();

        // Window mode: wait for window to disappear
        if (title != null && name == null && automationId == null)
        {
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                var hwnd = FindWindowByTitle(title);
                if (hwnd == nint.Zero)
                {
                    return Response.Ok(new
                    {
                        gone = true,
                        elapsed_ms = sw.ElapsedMilliseconds,
                        target = title,
                        type = "window",
                    });
                }
                Thread.Sleep(pollMs);
            }
            return Response.Error($"Timed out after {timeoutMs}ms waiting for window '{title}' to disappear");
        }

        // Element mode: wait for element to disappear
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            try
            {
                var parent = FindParentElement(automation, windowTitle ?? title);
                if (parent == null)
                {
                    // Parent window is gone — element is definitely gone
                    return Response.Ok(new
                    {
                        gone = true,
                        elapsed_ms = sw.ElapsedMilliseconds,
                        type = "element",
                        reason = "parent_window_gone",
                    });
                }

                AutomationElement? element = null;
                if (automationId != null)
                    element = parent.FindFirstDescendant(cf => cf.ByAutomationId(automationId));
                if (element == null && name != null)
                    element = parent.FindFirstDescendant(cf => cf.ByName(name));

                if (element == null)
                {
                    return Response.Ok(new
                    {
                        gone = true,
                        elapsed_ms = sw.ElapsedMilliseconds,
                        type = "element",
                    });
                }
            }
            catch
            {
                // UIA error likely means the element/window is gone
                return Response.Ok(new
                {
                    gone = true,
                    elapsed_ms = sw.ElapsedMilliseconds,
                    type = "element",
                    reason = "uia_error",
                });
            }

            Thread.Sleep(pollMs);
        }

        var desc = automationId != null ? $"automation_id='{automationId}'" : $"name='{name}'";
        return Response.Error($"Timed out after {timeoutMs}ms waiting for element {desc} to disappear");
    }

    // ── Helpers ──────────────────────────────────────────────────────

    private static string SafeGet(Func<string?> getter)
    {
        try { return getter() ?? ""; } catch { return ""; }
    }

    /// <summary>
    /// Find a parent element (window) by title. If no title given, uses the desktop.
    /// </summary>
    private AutomationElement? FindParentElement(UIA3Automation automation, string? windowTitle)
    {
        if (windowTitle == null)
            return automation.GetDesktop();

        var desktop = automation.GetDesktop();
        var children = desktop.FindAllChildren();
        foreach (var child in children)
        {
            try
            {
                var title = SafeGet(() => child.Name);
                if (title.Contains(windowTitle, StringComparison.OrdinalIgnoreCase))
                    return child;
            }
            catch { }
        }
        return null;
    }

    // ── Win32 P/Invoke ──────────────────────────────────────────────

    private delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(nint hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

    private static nint FindWindowByTitle(string titleSubstring)
    {
        nint found = nint.Zero;
        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd)) return true;
            var sb = new StringBuilder(512);
            GetWindowText(hWnd, sb, sb.Capacity);
            if (sb.ToString().Contains(titleSubstring, StringComparison.OrdinalIgnoreCase))
            {
                found = hWnd;
                return false;
            }
            return true;
        }, 0);
        return found;
    }

    private static string GetWindowTitle(nint hwnd)
    {
        var sb = new StringBuilder(512);
        GetWindowText(hwnd, sb, sb.Capacity);
        return sb.ToString();
    }
}
