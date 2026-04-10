using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.Com;

/// <summary>
/// General-purpose Windows desktop automation commands. These are app-agnostic —
/// they do NOT depend on ExcelSession and work with any Windows application.
/// Provides clipboard, process management, window control, keyboard/mouse input,
/// and system info.
/// </summary>
public class DesktopOps
{
    // ── P/Invoke declarations ────────────────────────────────────────

    private delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(nint hWnd, nint hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(nint hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(nint hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint FindWindow(string? lpClassName, string lpWindowName);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(nint hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, nint dwExtraInfo);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(nint hWnd, uint Msg, nint wParam, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(nint hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(nint hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint hWnd, out uint lpdwProcessId);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left, Top, Right, Bottom;
    }

    // mouse_event flags
    private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
    private const uint MOUSEEVENTF_LEFTUP = 0x0004;
    private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    private const uint MOUSEEVENTF_WHEEL = 0x0800;
    private const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
    private const uint MOUSEEVENTF_MOVE = 0x0001;

    // ShowWindow commands
    private const int SW_MINIMIZE = 6;
    private const int SW_MAXIMIZE = 3;
    private const int SW_RESTORE = 9;

    // SetWindowPos flags
    private const uint SWP_NOZORDER = 0x0004;

    // WM_CLOSE
    private const uint WM_CLOSE = 0x0010;

    public DesktopOps() { }

    public void Register(CommandRouter router)
    {
        // Clipboard
        router.Register("clipboard.read", HandleClipboardRead);
        router.Register("clipboard.write", HandleClipboardWrite);
        router.Register("clipboard.clear", HandleClipboardClear);
        router.Register("clipboard.formats", HandleClipboardFormats);

        // Process management
        router.Register("process.list", HandleProcessList);
        router.Register("process.start", HandleProcessStart);
        router.Register("process.kill", HandleProcessKill);
        router.Register("process.wait", HandleProcessWait);
        router.Register("process.info", HandleProcessInfo);

        // Window management
        router.Register("window.list", HandleWindowList);
        router.Register("window.move", HandleWindowMove);
        router.Register("window.minimize", HandleWindowMinimize);
        router.Register("window.maximize", HandleWindowMaximize);
        router.Register("window.restore", HandleWindowRestore);
        router.Register("window.close", HandleWindowClose);
        router.Register("window.focus", HandleWindowFocus);

        // Raw keyboard
        router.Register("keys.send", HandleKeysSend);

        // Raw mouse
        router.Register("mouse.click", HandleMouseClick);
        router.Register("mouse.move", HandleMouseMove);
        router.Register("mouse.scroll", HandleMouseScroll);

        // System info
        router.Register("system.info", HandleSystemInfo);
    }

    // ── Clipboard ────────────────────────────────────────────────────

    private string HandleClipboardRead(JsonObject args)
    {
        string? text = null;
        RunOnStaThread(() =>
        {
            text = System.Windows.Forms.Clipboard.GetText();
        });
        return Response.Ok(new { text = text ?? "" });
    }

    private string HandleClipboardWrite(JsonObject args)
    {
        var text = args["text"]?.GetValue<string>()
            ?? throw new ArgumentException("clipboard.write requires 'text'");
        RunOnStaThread(() =>
        {
            System.Windows.Forms.Clipboard.SetText(text);
        });
        return Response.Ok(new { written = true, length = text.Length });
    }

    private string HandleClipboardClear(JsonObject args)
    {
        RunOnStaThread(() =>
        {
            System.Windows.Forms.Clipboard.Clear();
        });
        return Response.Ok(new { cleared = true });
    }

    private string HandleClipboardFormats(JsonObject args)
    {
        string[]? formats = null;
        RunOnStaThread(() =>
        {
            var data = System.Windows.Forms.Clipboard.GetDataObject();
            formats = data?.GetFormats();
        });
        return Response.Ok(new { formats = formats ?? Array.Empty<string>() });
    }

    /// <summary>
    /// Clipboard operations require STA. If we're already on an STA thread
    /// (the StaComWorker), just run directly. Otherwise spin up a temporary
    /// STA thread.
    /// </summary>
    private static void RunOnStaThread(Action action)
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
        {
            action();
            return;
        }
        Exception? caught = null;
        var t = new Thread(() =>
        {
            try { action(); }
            catch (Exception ex) { caught = ex; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        t.Join();
        if (caught != null) throw caught;
    }

    // ── Process management ───────────────────────────────────────────

    private string HandleProcessList(JsonObject args)
    {
        var procs = Process.GetProcesses()
            .OrderByDescending(p => SafeGet(() => p.WorkingSet64))
            .Take(100)
            .Select(p => new
            {
                name = SafeGet(() => p.ProcessName) ?? "",
                pid = p.Id,
                memory_mb = Math.Round(SafeGet(() => p.WorkingSet64) / 1024.0 / 1024.0, 1),
                window_title = SafeGet(() => p.MainWindowTitle) ?? ""
            })
            .ToArray();
        return Response.Ok(new { count = procs.Length, processes = procs });
    }

    private string HandleProcessStart(JsonObject args)
    {
        var path = args["path"]?.GetValue<string>()
            ?? throw new ArgumentException("process.start requires 'path'");
        var arguments = args["args"]?.GetValue<string>();
        var workingDir = args["working_directory"]?.GetValue<string>();

        var psi = new ProcessStartInfo(path)
        {
            UseShellExecute = true,
        };
        if (!string.IsNullOrEmpty(arguments)) psi.Arguments = arguments;
        if (!string.IsNullOrEmpty(workingDir)) psi.WorkingDirectory = workingDir;

        var proc = Process.Start(psi)
            ?? throw new InvalidOperationException("Failed to start process");
        return Response.Ok(new { pid = proc.Id, name = proc.ProcessName });
    }

    private string HandleProcessKill(JsonObject args)
    {
        var pidNode = args["pid"];
        var nameNode = args["name"];

        var killed = new List<int>();

        if (pidNode != null)
        {
            var pid = pidNode.GetValue<int>();
            var proc = Process.GetProcessById(pid);
            proc.Kill(entireProcessTree: true);
            proc.WaitForExit(5000);
            killed.Add(pid);
        }
        else if (nameNode != null)
        {
            var name = nameNode.GetValue<string>()!;
            foreach (var proc in Process.GetProcessesByName(name))
            {
                try
                {
                    proc.Kill(entireProcessTree: true);
                    proc.WaitForExit(5000);
                    killed.Add(proc.Id);
                }
                catch { }
            }
        }
        else
        {
            throw new ArgumentException("process.kill requires 'pid' or 'name'");
        }

        return Response.Ok(new { killed_pids = killed, killed_count = killed.Count });
    }

    private string HandleProcessWait(JsonObject args)
    {
        var pid = args["pid"]?.GetValue<int>()
            ?? throw new ArgumentException("process.wait requires 'pid'");
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 30000;

        var proc = Process.GetProcessById(pid);
        var exited = proc.WaitForExit(timeoutMs);
        return Response.Ok(new
        {
            pid,
            exited,
            exit_code = exited ? proc.ExitCode : (int?)null
        });
    }

    private string HandleProcessInfo(JsonObject args)
    {
        var pid = args["pid"]?.GetValue<int>()
            ?? throw new ArgumentException("process.info requires 'pid'");

        var proc = Process.GetProcessById(pid);
        return Response.Ok(new
        {
            pid = proc.Id,
            name = proc.ProcessName,
            memory_mb = Math.Round(proc.WorkingSet64 / 1024.0 / 1024.0, 1),
            cpu_time_seconds = Math.Round(SafeGet(() => proc.TotalProcessorTime.TotalSeconds), 1),
            handle_count = SafeGet(() => proc.HandleCount),
            start_time = SafeGet(() => proc.StartTime.ToString("o")) ?? "",
            window_title = SafeGet(() => proc.MainWindowTitle) ?? "",
            responding = SafeGet(() => proc.Responding),
        });
    }

    // ── Window management ────────────────────────────────────────────

    private string HandleWindowList(JsonObject args)
    {
        var windows = new List<object>();
        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd)) return true;

            var titleBuf = new StringBuilder(256);
            GetWindowText(hWnd, titleBuf, titleBuf.Capacity);
            var title = titleBuf.ToString();
            if (string.IsNullOrEmpty(title)) return true;

            var classBuf = new StringBuilder(256);
            GetClassName(hWnd, classBuf, classBuf.Capacity);

            GetWindowThreadProcessId(hWnd, out uint procId);
            GetWindowRect(hWnd, out RECT rect);

            windows.Add(new
            {
                hwnd = hWnd.ToInt64(),
                title,
                class_name = classBuf.ToString(),
                pid = (int)procId,
                bounds = new { left = rect.Left, top = rect.Top, right = rect.Right, bottom = rect.Bottom }
            });
            return true;
        }, nint.Zero);

        return Response.Ok(new { count = windows.Count, windows });
    }

    private nint ResolveHwnd(JsonObject args)
    {
        var hwndNode = args["hwnd"];
        if (hwndNode != null)
            return (nint)hwndNode.GetValue<long>();

        var title = args["title"]?.GetValue<string>();
        if (title != null)
        {
            var found = FindWindow(null, title);
            if (found == nint.Zero)
                throw new ArgumentException($"No window found with title '{title}'");
            return found;
        }
        throw new ArgumentException("Requires 'hwnd' or 'title'");
    }

    private string HandleWindowMove(JsonObject args)
    {
        var hwnd = ResolveHwnd(args);
        var x = args["x"]?.GetValue<int>() ?? 0;
        var y = args["y"]?.GetValue<int>() ?? 0;
        var width = args["width"]?.GetValue<int>() ?? 0;
        var height = args["height"]?.GetValue<int>() ?? 0;

        // If width/height not specified, keep current size
        if (width == 0 || height == 0)
        {
            GetWindowRect(hwnd, out RECT currentRect);
            if (width == 0) width = currentRect.Right - currentRect.Left;
            if (height == 0) height = currentRect.Bottom - currentRect.Top;
        }

        SetWindowPos(hwnd, nint.Zero, x, y, width, height, SWP_NOZORDER);
        return Response.Ok(new { hwnd = hwnd.ToInt64(), x, y, width, height });
    }

    private string HandleWindowMinimize(JsonObject args)
    {
        var hwnd = ResolveHwnd(args);
        ShowWindow(hwnd, SW_MINIMIZE);
        return Response.Ok(new { hwnd = hwnd.ToInt64(), state = "minimized" });
    }

    private string HandleWindowMaximize(JsonObject args)
    {
        var hwnd = ResolveHwnd(args);
        ShowWindow(hwnd, SW_MAXIMIZE);
        return Response.Ok(new { hwnd = hwnd.ToInt64(), state = "maximized" });
    }

    private string HandleWindowRestore(JsonObject args)
    {
        var hwnd = ResolveHwnd(args);
        ShowWindow(hwnd, SW_RESTORE);
        return Response.Ok(new { hwnd = hwnd.ToInt64(), state = "restored" });
    }

    private string HandleWindowClose(JsonObject args)
    {
        var hwnd = ResolveHwnd(args);
        PostMessage(hwnd, WM_CLOSE, nint.Zero, nint.Zero);
        return Response.Ok(new { hwnd = hwnd.ToInt64(), closed = true });
    }

    private string HandleWindowFocus(JsonObject args)
    {
        var hwnd = ResolveHwnd(args);
        SetForegroundWindow(hwnd);
        return Response.Ok(new { hwnd = hwnd.ToInt64(), focused = true });
    }

    // ── Raw keyboard ─────────────────────────────────────────────────

    private string HandleKeysSend(JsonObject args)
    {
        var keys = args["keys"]?.GetValue<string>()
            ?? throw new ArgumentException("keys.send requires 'keys'");

        RunOnStaThread(() =>
        {
            System.Windows.Forms.SendKeys.SendWait(keys);
        });
        return Response.Ok(new { sent = keys });
    }

    // ── Raw mouse ────────────────────────────────────────────────────

    private string HandleMouseClick(JsonObject args)
    {
        var x = args["x"]?.GetValue<int>()
            ?? throw new ArgumentException("mouse.click requires 'x'");
        var y = args["y"]?.GetValue<int>()
            ?? throw new ArgumentException("mouse.click requires 'y'");
        var rightClick = args["right_click"]?.GetValue<bool>() ?? false;
        var doubleClick = args["double_click"]?.GetValue<bool>() ?? false;

        SetCursorPos(x, y);

        if (rightClick)
        {
            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, nint.Zero);
            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, nint.Zero);
        }
        else
        {
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, nint.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, nint.Zero);
            if (doubleClick)
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, nint.Zero);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, nint.Zero);
            }
        }

        return Response.Ok(new { x, y, right_click = rightClick, double_click = doubleClick });
    }

    private string HandleMouseMove(JsonObject args)
    {
        var x = args["x"]?.GetValue<int>()
            ?? throw new ArgumentException("mouse.move requires 'x'");
        var y = args["y"]?.GetValue<int>()
            ?? throw new ArgumentException("mouse.move requires 'y'");

        SetCursorPos(x, y);
        return Response.Ok(new { x, y });
    }

    private string HandleMouseScroll(JsonObject args)
    {
        var amount = args["amount"]?.GetValue<int>()
            ?? throw new ArgumentException("mouse.scroll requires 'amount'");

        // Positive = scroll up, negative = scroll down. WHEEL_DELTA = 120.
        mouse_event(MOUSEEVENTF_WHEEL, 0, 0, (uint)(amount * 120), nint.Zero);
        return Response.Ok(new { scrolled = amount });
    }

    // ── System info ──────────────────────────────────────────────────

    private string HandleSystemInfo(JsonObject args)
    {
        return Response.Ok(new
        {
            os_version = Environment.OSVersion.VersionString,
            architecture = RuntimeInformation.OSArchitecture.ToString(),
            dotnet_version = RuntimeInformation.FrameworkDescription,
            display_count = System.Windows.Forms.Screen.AllScreens.Length,
            total_memory_gb = Math.Round(GC.GetGCMemoryInfo().TotalAvailableMemoryBytes / 1024.0 / 1024.0 / 1024.0, 1),
            computer_name = Environment.MachineName,
            user_name = Environment.UserName,
            processor_count = Environment.ProcessorCount,
        });
    }

    // ── Helpers ──────────────────────────────────────────────────────

    private static T SafeGet<T>(Func<T> getter, T defaultValue = default!)
    {
        try { return getter(); }
        catch { return defaultValue; }
    }

    private static string? SafeGet(Func<string?> getter)
    {
        try { return getter(); }
        catch { return null; }
    }
}
