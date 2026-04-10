using System.Diagnostics;
using System.Text.Json.Nodes;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using XRai.Core;

namespace XRai.UI;

/// <summary>
/// Generalized app attach — connect to any running Windows application via UIA.
/// No ExcelSession dependency; works with any process.
/// </summary>
public class AppAttachOps
{
    private UIA3Automation? _automation;
    private Application? _attachedApp;
    private Window? _attachedWindow;

    public AppAttachOps() { }

    public void Register(CommandRouter router)
    {
        router.Register("app.launch", HandleAppLaunch);
        router.Register("app.list", HandleAppList);
        router.Register("app.attach", HandleAppAttach);
    }

    private string HandleAppLaunch(JsonObject args)
    {
        var path = args["path"]?.GetValue<string>()
            ?? throw new ArgumentException("app.launch requires 'path'");
        var arguments = args["args"]?.GetValue<string>();
        var timeoutSeconds = args["timeout"]?.GetValue<int>() ?? 10;

        var psi = new ProcessStartInfo(path)
        {
            UseShellExecute = true,
        };
        if (!string.IsNullOrEmpty(arguments)) psi.Arguments = arguments;

        var proc = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start process");

        // Wait for the main window to appear
        var deadline = DateTime.UtcNow.AddSeconds(timeoutSeconds);
        while (DateTime.UtcNow < deadline)
        {
            proc.Refresh();
            if (proc.MainWindowHandle != nint.Zero) break;
            Thread.Sleep(200);
        }

        var hwnd = proc.MainWindowHandle.ToInt64();
        var title = SafeGet(() => proc.MainWindowTitle) ?? "";

        return Response.Ok(new
        {
            pid = proc.Id,
            name = proc.ProcessName,
            hwnd,
            window_title = title
        });
    }

    private string HandleAppList(JsonObject args)
    {
        var apps = new List<object>();
        var seen = new HashSet<int>();

        foreach (var proc in Process.GetProcesses())
        {
            try
            {
                if (proc.MainWindowHandle == nint.Zero) continue;
                var title = proc.MainWindowTitle;
                if (string.IsNullOrEmpty(title)) continue;
                if (!seen.Add(proc.Id)) continue;

                apps.Add(new
                {
                    pid = proc.Id,
                    name = proc.ProcessName,
                    title,
                    hwnd = proc.MainWindowHandle.ToInt64()
                });
            }
            catch { }
        }

        return Response.Ok(new { count = apps.Count, apps });
    }

    private string HandleAppAttach(JsonObject args)
    {
        var pidNode = args["pid"];
        var title = args["title"]?.GetValue<string>();

        Process? target = null;

        if (pidNode != null)
        {
            var pid = pidNode.GetValue<int>();
            target = Process.GetProcessById(pid);
        }
        else if (title != null)
        {
            // Find process by window title (partial match)
            foreach (var proc in Process.GetProcesses())
            {
                try
                {
                    var wt = proc.MainWindowTitle;
                    if (!string.IsNullOrEmpty(wt) && wt.Contains(title, StringComparison.OrdinalIgnoreCase))
                    {
                        target = proc;
                        break;
                    }
                }
                catch { }
            }
            if (target == null)
                throw new ArgumentException($"No process found with window title containing '{title}'");
        }
        else
        {
            throw new ArgumentException("app.attach requires 'pid' or 'title'");
        }

        // Attach FlaUI
        _automation?.Dispose();
        _automation = new UIA3Automation();
        _attachedApp = Application.Attach(target);
        _attachedWindow = _attachedApp.GetMainWindow(_automation, TimeSpan.FromSeconds(5));

        if (_attachedWindow == null)
            throw new InvalidOperationException($"Could not get main window for process {target.Id}");

        // Build a summary of the UIA tree root
        var root = _attachedWindow;
        var children = new JsonArray();
        foreach (var child in root.FindAllChildren())
        {
            children.Add(new JsonObject
            {
                ["name"] = SafeGet(() => child.Name) ?? "",
                ["type"] = SafeGet(() => child.ControlType.ToString()) ?? "",
                ["automation_id"] = SafeGet(() => child.AutomationId) ?? "",
            });
            if (children.Count >= 50) break; // cap to avoid token bloat
        }

        return Response.Ok(new
        {
            attached = true,
            pid = target.Id,
            name = target.ProcessName,
            window_title = SafeGet(() => root.Title) ?? "",
            hwnd = root.Properties.NativeWindowHandle.ValueOrDefault.ToInt64(),
            child_count = children.Count,
            children
        });
    }

    private static string? SafeGet(Func<string?> getter)
    {
        try { return getter(); } catch { return null; }
    }
}
