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
/// Drives native Windows folder pickers and Open/Save file dialogs via FlaUI.
/// These dialogs are UIA-accessible but are NOT COM modals — they're spawned as
/// separate top-level windows with ClassName "#32770" (standard Win32 dialog).
/// RibbonDriver's dialog.* commands only see COM modal children of the Excel
/// main window, so this driver fills the critical gap for agent-driven add-in
/// setup flows (picking vault folders, export paths, config files, etc.).
///
/// Commands:
///   dialog.wait         Block until a dialog with matching title appears
///   folder.dialog.pick  Type path + click Select Folder in any open folder picker
///   file.dialog.pick    Type filename + click Open/Save in any open file dialog
/// </summary>
public class FileDialogDriver
{
    private UIA3Automation? _automation;

    // Win32 fallback — some WinForms and .NET 6+ modal dialogs don't appear in the
    // UIA desktop walker's FindAllChildren result (their accessibility exposure
    // can be flaky). EnumWindows finds EVERY top-level window regardless of how
    // well-behaved it is. We use this as a fallback when the UIA walk comes up empty.
    private delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(nint hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(nint hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint SendMessage(nint hWnd, uint msg, nint wParam, nint lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
    private static extern nint SendMessageText(nint hWnd, uint msg, nint wParam, string lParam);

    private const uint WM_SETTEXT = 0x000C;
    private const uint WM_KEYDOWN = 0x0100;
    private const uint WM_KEYUP = 0x0101;
    private const uint VK_RETURN = 0x0D;

    // Shell COM interfaces for programmatic folder dialog navigation
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHCreateItemFromParsingName(
        string pszPath, IntPtr pbc,
        [MarshalAs(UnmanagedType.LPStruct)] Guid riid,
        out IntPtr ppv);

    [DllImport("user32.dll")]
    private static extern nint SendMessageTimeout(
        nint hWnd, uint Msg, nint wParam, nint lParam,
        uint fuFlags, uint uTimeout, out nint lpdwResult);

    // CDM_SETCONTROLTEXT for common dialog controls
    private const uint CDM_FIRST = 0x0601;
    private const uint CDM_SETCONTROLTEXT = CDM_FIRST + 0x0004;

    public void Register(CommandRouter router)
    {
        router.Register("dialog.wait", HandleWait);
        router.Register("dialog.list", HandleDialogList);
        router.Register("folder.dialog.pick", HandleFolderPick);
        router.Register("folder.dialog.navigate", HandleFolderNavigate);
        router.Register("folder.dialog.set_path", HandleFolderSetPath);
        router.Register("file.dialog.pick", HandleFilePick);
    }

    private UIA3Automation GetAutomation()
    {
        return _automation ??= new UIA3Automation();
    }

    private static string SafeGet(Func<string?> getter)
    {
        try { return getter() ?? ""; } catch { return ""; }
    }

    // ── Win32 enumeration fallback ───────────────────────────────────

    /// <summary>
    /// Walk every top-level window via EnumWindows, filter to those owned by
    /// the Excel process, and return their HWNDs + titles + class names.
    /// This catches WinForms and other dialogs that the UIA desktop walker
    /// sometimes misses.
    /// </summary>
    private List<(nint Hwnd, string Title, string ClassName)> EnumerateWin32Windows(HashSet<int> excelPids)
    {
        var result = new List<(nint, string, string)>();
        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd)) return true;
            GetWindowThreadProcessId(hWnd, out uint pid);
            if (!excelPids.Contains((int)pid)) return true;

            var titleSb = new StringBuilder(512);
            GetWindowText(hWnd, titleSb, titleSb.Capacity);
            var title = titleSb.ToString();

            var classSb = new StringBuilder(256);
            GetClassName(hWnd, classSb, classSb.Capacity);
            var className = classSb.ToString();

            result.Add((hWnd, title, className));
            return true;
        }, 0);
        return result;
    }

    /// <summary>
    /// Wrap a Win32 HWND as a FlaUI AutomationElement by asking UIA for its
    /// element-from-handle lookup. This gives us full UIA access (descendants,
    /// patterns, children) on a window we found via Win32 enumeration.
    /// </summary>
    private AutomationElement? Win32HwndToUiaElement(nint hwnd)
    {
        try
        {
            return GetAutomation().FromHandle(hwnd);
        }
        catch
        {
            return null;
        }
    }

    // ── dialog.wait ──────────────────────────────────────────────────

    /// <summary>
    /// Poll for a top-level window whose title contains the given substring
    /// and belongs to an Excel process. Returns the window's UIA tree once
    /// found, or errors on timeout.
    /// </summary>
    private string HandleWait(JsonObject args)
    {
        // Match by EITHER title substring OR kind (folder / file_open / file_save / any).
        // kind is preferred when the app sets UseDescriptionForTitle=true and the agent
        // has no idea what the actual title text will be.
        var titleFilter = args["title"]?.GetValue<string>();
        var kind = args["kind"]?.GetValue<string>();

        if (titleFilter == null && kind == null)
            throw new ArgumentException("dialog.wait requires 'title' or 'kind' (or both)");

        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 5000;
        var pollMs = args["poll_ms"]?.GetValue<int>() ?? 200;

        var automation = GetAutomation();
        var sw = Stopwatch.StartNew();
        var excelPids = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        if (excelPids.Count == 0)
            return Response.Error("No Excel process found");

        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            try
            {
                var dialog = FindDialog(automation, titleFilter, kind, excelPids);
                if (dialog != null)
                {
                    return Response.Ok(new
                    {
                        found = true,
                        title = SafeGet(() => dialog.Name),
                        class_name = SafeGet(() => dialog.ClassName),
                        automation_id = SafeGet(() => dialog.AutomationId),
                        elapsed_ms = sw.ElapsedMilliseconds,
                        buttons = DumpButtons(dialog),
                    });
                }
            }
            catch { /* ignore transient UIA errors */ }

            Thread.Sleep(pollMs);
        }

        // Timeout — return helpful error with the list of dialogs we DID see
        var candidates = ListCandidateDialogs(automation, excelPids);
        var errorObj = new JsonObject
        {
            ["ok"] = false,
            ["error"] = titleFilter != null
                ? $"No dialog with title containing '{titleFilter}' appeared within {timeoutMs}ms"
                : $"No dialog of kind '{kind}' appeared within {timeoutMs}ms",
            ["suggestion"] = "Use {\"cmd\":\"dialog.list\"} to see all currently-open dialogs with their titles/classes. " +
                "For folder pickers that set UseDescriptionForTitle=true, use {\"kind\":\"folder\"} instead of a title filter.",
            ["candidates"] = candidates,
        };
        return errorObj.ToJsonString();
    }

    /// <summary>
    /// List every top-level Excel-owned dialog-like window with title/class/buttons.
    /// Use this when dialog.wait times out to see what's actually open.
    /// </summary>
    private string HandleDialogList(JsonObject args)
    {
        var automation = GetAutomation();
        var excelPids = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        if (excelPids.Count == 0)
            return Response.Error("No Excel process found");

        var candidates = ListCandidateDialogs(automation, excelPids);
        return Response.Ok(new { dialogs = candidates, count = candidates.Count });
    }

    private JsonArray ListCandidateDialogs(UIA3Automation automation, HashSet<int> excelPids)
    {
        var result = new JsonArray();
        var seenHwnds = new HashSet<long>();

        // Primary: UIA desktop walk
        try
        {
            var desktop = automation.GetDesktop();
            var children = desktop.FindAllChildren();
            foreach (var child in children)
            {
                try
                {
                    if (!excelPids.Contains(child.Properties.ProcessId.Value)) continue;

                    var cls = SafeGet(() => child.ClassName);
                    var title = SafeGet(() => child.Name);
                    if (cls == "XLMAIN") continue;

                    long hwnd = 0;
                    try { hwnd = child.Properties.NativeWindowHandle.ValueOrDefault.ToInt64(); } catch { }
                    if (hwnd != 0) seenHwnds.Add(hwnd);

                    result.Add(new JsonObject
                    {
                        ["title"] = title,
                        ["class_name"] = cls,
                        ["automation_id"] = SafeGet(() => child.AutomationId),
                        ["kind_guess"] = GuessDialogKind(cls, title),
                        ["button_count"] = GetButtonNames(child).Length,
                        ["hwnd"] = hwnd,
                        ["source"] = "uia",
                    });
                }
                catch { }
            }
        }
        catch { }

        // Fallback: Win32 EnumWindows — catches WinForms dialogs UIA might miss
        try
        {
            var win32Windows = EnumerateWin32Windows(excelPids);
            foreach (var (hwnd, title, cls) in win32Windows)
            {
                if (cls == "XLMAIN") continue;
                if (seenHwnds.Contains(hwnd.ToInt64())) continue; // already seen via UIA

                result.Add(new JsonObject
                {
                    ["title"] = title,
                    ["class_name"] = cls,
                    ["automation_id"] = "",
                    ["kind_guess"] = GuessDialogKind(cls, title),
                    ["button_count"] = 0,
                    ["hwnd"] = hwnd.ToInt64(),
                    ["source"] = "win32",
                });
            }
        }
        catch { }

        return result;
    }

    private static string GuessDialogKind(string className, string title)
    {
        // Standard Win32 dialog class
        if (className != "#32770" && !className.Contains("Dialog", StringComparison.OrdinalIgnoreCase))
            return "other";

        var lower = title.ToLowerInvariant();

        // Folder browser signatures (Windows Vista+ IFileDialog and legacy SHBrowseForFolder)
        if (lower.Contains("browse for folder") ||
            lower.Contains("select folder") ||
            lower.Contains("choose folder") ||
            lower.Contains("select a folder") ||
            lower.Contains("pick a folder"))
            return "folder";

        // Open file
        if (lower.StartsWith("open") || lower.Contains("choose file") || lower.Contains("select file"))
            return "file_open";

        // Save file
        if (lower.StartsWith("save") || lower.Contains("save as"))
            return "file_save";

        return "dialog";
    }

    // ── folder.dialog.pick ───────────────────────────────────────────

    /// <summary>
    /// Drive an already-open folder picker dialog:
    ///   1. Find the dialog (any ClassName #32770 with "Folder" in title, or matching a custom title filter)
    ///   2. Find the path Edit field (usually AutomationId "1148" or first Edit descendant)
    ///   3. Type the target path
    ///   4. Click the confirm button (AutomationId "1", or by name "Select Folder" / "OK")
    /// </summary>
    private string HandleFolderPick(JsonObject args)
    {
        var path = args["path"]?.GetValue<string>()
            ?? throw new ArgumentException("folder.dialog.pick requires 'path'");
        var titleFilter = args["title"]?.GetValue<string>();
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 3000;

        var automation = GetAutomation();
        var excelPids = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        if (excelPids.Count == 0)
            return Response.Error("No Excel process found");

        // Find the dialog — wait briefly if not immediately present
        AutomationElement? dialog = null;
        var sw = Stopwatch.StartNew();
        while (dialog == null && sw.ElapsedMilliseconds < timeoutMs)
        {
            dialog = FindFolderPicker(automation, titleFilter, excelPids);
            if (dialog == null) Thread.Sleep(200);
        }

        if (dialog == null)
            return Response.Error(
                $"No folder picker dialog found" +
                (titleFilter != null ? $" (title contains '{titleFilter}')" : "") +
                ". Ensure the button that opens the picker has been clicked first. " +
                "Use {\"cmd\":\"dialog.wait\",\"title\":\"...\"} before this command to block until it appears.");

        // Find the path input field
        var pathInput = FindPathInput(dialog);
        if (pathInput == null)
            return Response.Error($"Path input field not found in dialog '{SafeGet(() => dialog.Name)}'. " +
                "Buttons available: " + string.Join(", ", GetButtonNames(dialog)));

        // Type the path
        try
        {
            var valuePattern = pathInput.Patterns.Value.PatternOrDefault;
            if (valuePattern != null)
            {
                valuePattern.SetValue(path);
            }
            else
            {
                pathInput.Focus();
                pathInput.AsTextBox().Enter(path);
            }
        }
        catch (Exception ex)
        {
            return Response.Error($"Failed to type path into folder picker: {ex.Message}");
        }

        // Short delay to let the dialog update its internal state
        Thread.Sleep(150);

        // Click the confirm button
        var confirm = FindConfirmButton(dialog, ["Select Folder", "OK", "Open", "Choose"]);
        if (confirm == null)
            return Response.Error("Confirm button not found in folder picker. " +
                "Buttons available: " + string.Join(", ", GetButtonNames(dialog)));

        try
        {
            var invoke = confirm.Patterns.Invoke.PatternOrDefault;
            if (invoke != null) invoke.Invoke();
            else confirm.Click();
        }
        catch (Exception ex)
        {
            return Response.Error($"Failed to click confirm button: {ex.Message}");
        }

        return Response.Ok(new
        {
            picked = true,
            path,
            dialog_title = SafeGet(() => dialog.Name),
            confirm_button = SafeGet(() => confirm.Name)
        });
    }

    // ── folder.dialog.set_path ─────────────────────────────────────────

    /// <summary>
    /// Programmatically navigate an open folder picker to a specific path.
    /// Uses character-by-character keyboard simulation into the address bar
    /// which bypasses ALL JSON escaping issues (backslashes are sent as
    /// actual keystrokes, not as text strings that get re-interpreted).
    ///
    /// Steps:
    ///   1. Find the folder picker dialog
    ///   2. Click the breadcrumb bar to activate address bar edit mode
    ///   3. Clear the current text (Ctrl+A, Delete)
    ///   4. Type the path character by character using SendInput
    ///   5. Press Enter to navigate
    ///   6. Optionally click "Select Folder" if pick:true
    /// </summary>
    private string HandleFolderSetPath(JsonObject args)
    {
        var path = args["path"]?.GetValue<string>()
            ?? throw new ArgumentException("folder.dialog.set_path requires 'path'");
        var pick = args["pick"]?.GetValue<bool>() ?? false;
        var titleFilter = args["title"]?.GetValue<string>();
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 5000;

        var automation = GetAutomation();
        var excelPids = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        if (excelPids.Count == 0)
            return Response.Error("No Excel process found");

        // Find the dialog
        AutomationElement? dialog = null;
        var sw = Stopwatch.StartNew();
        while (dialog == null && sw.ElapsedMilliseconds < timeoutMs)
        {
            dialog = FindFolderPicker(automation, titleFilter, excelPids);
            if (dialog == null) Thread.Sleep(200);
        }

        if (dialog == null)
            return Response.Error(
                "No folder picker dialog found. " +
                "Ensure the folder picker button has been clicked first.");

        try
        {
            // Step 1: Click the breadcrumb / address bar area to activate edit mode
            // AutomationId "1001" is the breadcrumb ToolBar in IFileDialog
            var breadcrumb = dialog.FindFirstDescendant(cf => cf.ByAutomationId("1001"));
            if (breadcrumb != null)
            {
                try { breadcrumb.Click(); } catch { }
                Thread.Sleep(400);
            }

            // Step 2: Find the address bar edit control
            var addressEdit = dialog.FindFirstDescendant(cf =>
                cf.ByControlType(ControlType.Edit).And(cf.ByAutomationId("41477")));

            // Fallback: try the ComboBoxEx32 address bar
            if (addressEdit == null)
            {
                addressEdit = dialog.FindFirstDescendant(cf =>
                    cf.ByControlType(ControlType.ComboBox).And(cf.ByAutomationId("41477")));
            }

            // Fallback: try the standard path edit (AutomationId 1152)
            if (addressEdit == null)
            {
                addressEdit = dialog.FindFirstDescendant(cf => cf.ByAutomationId("1152"));
            }

            // Fallback: any Edit control
            if (addressEdit == null)
            {
                var edits = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Edit));
                addressEdit = edits.FirstOrDefault(e => e.IsEnabled && !e.IsOffscreen);
            }

            if (addressEdit == null)
                return Response.Error("Could not find address bar or path edit in folder picker. " +
                    "Buttons: " + string.Join(", ", GetButtonNames(dialog)));

            // Step 3: Focus the edit and clear it
            addressEdit.Focus();
            Thread.Sleep(100);

            // Ctrl+A to select all, then Delete to clear
            FlaUI.Core.Input.Keyboard.TypeSimultaneously(
                FlaUI.Core.WindowsAPI.VirtualKeyShort.CONTROL,
                FlaUI.Core.WindowsAPI.VirtualKeyShort.KEY_A);
            Thread.Sleep(50);
            FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.DELETE);
            Thread.Sleep(100);

            // Step 4: Type the path character by character
            // This bypasses ALL escaping issues — backslashes are real keystrokes
            FlaUI.Core.Input.Keyboard.Type(path);
            Thread.Sleep(300);

            // Step 5: Press Enter to navigate
            FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.ENTER);
            Thread.Sleep(800); // Let the dialog navigate

            string? pickResult = null;

            // Step 6: Optionally click Select Folder
            if (pick)
            {
                Thread.Sleep(500);
                var confirm = FindConfirmButton(dialog, ["Select Folder", "OK", "Open", "Choose"]);
                if (confirm != null)
                {
                    try
                    {
                        var invoke = confirm.Patterns.Invoke.PatternOrDefault;
                        if (invoke != null) invoke.Invoke();
                        else confirm.Click();
                        pickResult = SafeGet(() => confirm.Name);
                    }
                    catch { }
                }
            }

            return Response.Ok(new
            {
                navigated = true,
                path,
                method = "keyboard_type",
                picked = pick && pickResult != null,
                confirm_button = pickResult,
                dialog_title = SafeGet(() => dialog.Name),
            });
        }
        catch (Exception ex)
        {
            return Response.Error($"folder.dialog.set_path failed: {ex.Message}");
        }
    }

    // ── folder.dialog.navigate ─────────────────────────────────────────

    /// <summary>
    /// Navigate an open folder picker to a path by typing into its address bar
    /// (not the filename Edit field). This is the reliable way to set the folder
    /// path — the filename Edit field interprets text as navigation input and
    /// mangles backslashes. The address bar accepts full paths and navigates on Enter.
    ///
    /// Steps:
    ///   1. Find the open folder picker dialog
    ///   2. Find the breadcrumb/address bar (ToolBar with AutomationId "1001")
    ///   3. Click it to activate edit mode (reveals an Edit control)
    ///   4. Type the path + press Enter to navigate
    ///   5. Return success (caller clicks "Select Folder" separately)
    /// </summary>
    private string HandleFolderNavigate(JsonObject args)
    {
        var path = args["path"]?.GetValue<string>()
            ?? throw new ArgumentException("folder.dialog.navigate requires 'path'");
        var titleFilter = args["title"]?.GetValue<string>();
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 3000;

        var automation = GetAutomation();
        var excelPids = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        if (excelPids.Count == 0)
            return Response.Error("No Excel process found");

        AutomationElement? dialog = null;
        var sw = Stopwatch.StartNew();
        while (dialog == null && sw.ElapsedMilliseconds < timeoutMs)
        {
            dialog = FindFolderPicker(automation, titleFilter, excelPids);
            if (dialog == null) Thread.Sleep(200);
        }

        if (dialog == null)
            return Response.Error(
                "No folder picker dialog found" +
                (titleFilter != null ? $" (title contains '{titleFilter}')" : "") +
                ". Use {\"cmd\":\"dialog.wait\",\"kind\":\"folder\"} first.");

        // Strategy 1: Find the breadcrumb ToolBar (AutomationId "1001" in IFileDialog)
        // and click it to reveal the address bar Edit. Then type the path and press Enter.
        try
        {
            var breadcrumb = dialog.FindFirstDescendant(cf => cf.ByAutomationId("1001"));
            if (breadcrumb != null)
            {
                // Click the breadcrumb to switch to edit mode
                try { breadcrumb.Click(); } catch { }
                Thread.Sleep(300);

                // Now look for the Edit control that appeared (address bar edit mode)
                // It's typically a ComboBox with an Edit child, or a direct Edit
                var addressEdit = dialog.FindFirstDescendant(cf =>
                    cf.ByControlType(ControlType.Edit).And(cf.ByAutomationId("41477")));

                // Fallback: any Edit that appeared after clicking the breadcrumb
                addressEdit ??= dialog.FindFirstDescendant(cf =>
                    cf.ByControlType(ControlType.ComboBox).And(cf.ByAutomationId("41477")));

                if (addressEdit != null)
                {
                    var vp = addressEdit.Patterns.Value.PatternOrDefault;
                    if (vp != null)
                    {
                        vp.SetValue(path);
                    }
                    else
                    {
                        // Win32 fallback: WM_SETTEXT on the native handle
                        var hwnd = addressEdit.Properties.NativeWindowHandle.ValueOrDefault;
                        if (hwnd != nint.Zero)
                            SendMessageText(hwnd, WM_SETTEXT, 0, path);
                        else
                        {
                            addressEdit.Focus();
                            addressEdit.AsTextBox().Enter(path);
                        }
                    }

                    // Press Enter to navigate
                    Thread.Sleep(100);
                    var editHwnd = addressEdit.Properties.NativeWindowHandle.ValueOrDefault;
                    if (editHwnd != nint.Zero)
                    {
                        SendMessage(editHwnd, WM_KEYDOWN, (nint)VK_RETURN, 0);
                        SendMessage(editHwnd, WM_KEYUP, (nint)VK_RETURN, 0);
                    }
                    else
                    {
                        // UIA keyboard fallback
                        FlaUI.Core.Input.Keyboard.Press(FlaUI.Core.WindowsAPI.VirtualKeyShort.ENTER);
                    }

                    Thread.Sleep(500); // Let the dialog navigate

                    return Response.Ok(new
                    {
                        navigated = true,
                        path,
                        method = "address_bar",
                        dialog_title = SafeGet(() => dialog.Name),
                    });
                }
            }
        }
        catch { /* fall through to strategy 2 */ }

        // Strategy 2: No breadcrumb found (legacy SHBrowseForFolder or unusual dialog).
        // Try to type directly into the path Edit field (AutomationId "1148")
        // using UIA ValuePattern (not WM_SETTEXT which mangles backslashes in
        // folder browser edit fields).
        var pathInput = FindPathInput(dialog);
        if (pathInput != null)
        {
            try
            {
                var vp = pathInput.Patterns.Value.PatternOrDefault;
                if (vp != null)
                {
                    vp.SetValue(path);
                    Thread.Sleep(100);
                    // Press Enter to navigate/confirm
                    var hwnd = pathInput.Properties.NativeWindowHandle.ValueOrDefault;
                    if (hwnd != nint.Zero)
                    {
                        SendMessage(hwnd, WM_KEYDOWN, (nint)VK_RETURN, 0);
                        SendMessage(hwnd, WM_KEYUP, (nint)VK_RETURN, 0);
                    }
                    Thread.Sleep(500);

                    return Response.Ok(new
                    {
                        navigated = true,
                        path,
                        method = "path_edit_value_pattern",
                        dialog_title = SafeGet(() => dialog.Name),
                    });
                }
            }
            catch { }
        }

        return Response.Error(
            $"Could not navigate folder picker to '{path}'. " +
            "No address bar (AutomationId 1001) or path edit field found. " +
            "Buttons: " + string.Join(", ", GetButtonNames(dialog)));
    }

    // ── file.dialog.pick ─────────────────────────────────────────────

    /// <summary>
    /// Drive an already-open Open/Save file dialog:
    ///   action:"open" → finds Open File dialog, types filename, clicks Open
    ///   action:"save" → finds Save As dialog, types filename, clicks Save
    /// </summary>
    private string HandleFilePick(JsonObject args)
    {
        var path = args["path"]?.GetValue<string>()
            ?? throw new ArgumentException("file.dialog.pick requires 'path'");
        var action = args["action"]?.GetValue<string>() ?? "open";
        var titleFilter = args["title"]?.GetValue<string>();
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 3000;

        if (action != "open" && action != "save")
            return Response.Error("file.dialog.pick requires 'action' to be 'open' or 'save'");

        var automation = GetAutomation();
        var excelPids = Process.GetProcessesByName("EXCEL").Select(p => p.Id).ToHashSet();
        if (excelPids.Count == 0)
            return Response.Error("No Excel process found");

        // Default title filters per action
        titleFilter ??= action == "save" ? "Save" : "Open";

        AutomationElement? dialog = null;
        var sw = Stopwatch.StartNew();
        while (dialog == null && sw.ElapsedMilliseconds < timeoutMs)
        {
            dialog = FindFileDialog(automation, titleFilter, excelPids);
            if (dialog == null) Thread.Sleep(200);
        }

        if (dialog == null)
            return Response.Error(
                $"No {action} file dialog found (title contains '{titleFilter}'). " +
                "Ensure the button that opens the dialog has been clicked first. " +
                "Use {\"cmd\":\"dialog.wait\",\"title\":\"...\"} before this command.");

        // Find the filename Edit field. In Open/Save dialogs it's usually labeled
        // "File name:" — we look for an Edit control by that name first, then fall
        // back to the first enabled Edit descendant.
        var fileInput = FindFileNameInput(dialog);
        if (fileInput == null)
            return Response.Error($"Filename input not found in dialog '{SafeGet(() => dialog.Name)}'. " +
                "Buttons available: " + string.Join(", ", GetButtonNames(dialog)));

        try
        {
            var valuePattern = fileInput.Patterns.Value.PatternOrDefault;
            if (valuePattern != null)
            {
                valuePattern.SetValue(path);
            }
            else
            {
                fileInput.Focus();
                fileInput.AsTextBox().Enter(path);
            }
        }
        catch (Exception ex)
        {
            return Response.Error($"Failed to type filename: {ex.Message}");
        }

        Thread.Sleep(150);

        var confirmButtonNames = action == "save"
            ? new[] { "Save", "OK" }
            : new[] { "Open", "OK" };

        var confirm = FindConfirmButton(dialog, confirmButtonNames);
        if (confirm == null)
            return Response.Error($"{action} confirm button not found in dialog. " +
                "Buttons available: " + string.Join(", ", GetButtonNames(dialog)));

        try
        {
            var invoke = confirm.Patterns.Invoke.PatternOrDefault;
            if (invoke != null) invoke.Invoke();
            else confirm.Click();
        }
        catch (Exception ex)
        {
            return Response.Error($"Failed to click {action} button: {ex.Message}");
        }

        return Response.Ok(new
        {
            picked = true,
            action,
            path,
            dialog_title = SafeGet(() => dialog.Name),
            confirm_button = SafeGet(() => confirm.Name)
        });
    }

    // ── Dialog discovery helpers ─────────────────────────────────────

    /// <summary>
    /// Unified dialog finder: matches by title substring, kind, or both.
    /// If both are provided, BOTH must match. If neither matches, returns null.
    /// kind values: "folder", "file_open", "file_save", "any" (any dialog-class window)
    /// </summary>
    private AutomationElement? FindDialog(UIA3Automation automation, string? titleFilter, string? kind, HashSet<int> excelPids)
    {
        // Step 1: UIA desktop walk (fast path, captures most WPF dialogs)
        var desktop = automation.GetDesktop();
        var children = desktop.FindAllChildren();

        foreach (var child in children)
        {
            try
            {
                if (!excelPids.Contains(child.Properties.ProcessId.Value)) continue;
                var cls = SafeGet(() => child.ClassName);
                var title = SafeGet(() => child.Name);
                if (cls == "XLMAIN") continue;

                if (MatchesFilters(child, cls, title, titleFilter, kind))
                    return child;
            }
            catch { }
        }

        // Step 2: Win32 EnumWindows fallback. UIA's desktop walker misses some
        // WinForms/InputBox dialogs because their accessibility tree isn't fully
        // populated until after they process messages. Win32 sees them immediately.
        var win32Windows = EnumerateWin32Windows(excelPids);
        foreach (var (hwnd, winTitle, winClass) in win32Windows)
        {
            if (winClass == "XLMAIN") continue;

            // Fast pre-check on Win32 attributes before the expensive UIA lookup
            if (titleFilter != null && !winTitle.Contains(titleFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            // Kind pre-check on the Win32 metadata
            if (kind != null)
            {
                var preKind = GuessDialogKind(winClass, winTitle);
                bool preMatches = kind.ToLowerInvariant() switch
                {
                    "any" => winClass == "#32770" || winClass.Contains("Dialog", StringComparison.OrdinalIgnoreCase),
                    "folder" => preKind == "folder",  // structural check needs UIA element, done below
                    "file_open" => preKind == "file_open",
                    "file_save" => preKind == "file_save",
                    _ => false,
                };
                // If title filter also provided, we can accept a pre-kind mismatch
                // and rely on the structural check via the UIA element. Otherwise
                // skip early.
                if (!preMatches && titleFilter == null && kind != "folder") continue;
            }

            // Get the UIA element for this HWND and do the full check
            var element = Win32HwndToUiaElement(hwnd);
            if (element == null) continue;

            if (MatchesFilters(element, winClass, winTitle, titleFilter, kind))
                return element;
        }

        return null;
    }

    private bool MatchesFilters(AutomationElement element, string cls, string title, string? titleFilter, string? kind)
    {
        if (kind != null)
        {
            var detectedKind = GuessDialogKind(cls, title);
            bool kindMatches = kind.ToLowerInvariant() switch
            {
                "any" => cls == "#32770" || cls.Contains("Dialog", StringComparison.OrdinalIgnoreCase),
                "folder" => detectedKind == "folder" || IsLikelyFolderPicker(element),
                "file_open" => detectedKind == "file_open",
                "file_save" => detectedKind == "file_save",
                _ => false,
            };
            if (!kindMatches) return false;
        }

        if (titleFilter != null)
        {
            if (string.IsNullOrWhiteSpace(title)) return false;
            if (!title.Contains(titleFilter, StringComparison.OrdinalIgnoreCase)) return false;
        }

        return titleFilter != null || kind != null;
    }

    /// <summary>
    /// When the app sets UseDescriptionForTitle=true, the folder picker's title
    /// becomes the arbitrary Description text (e.g. "Select your CellVault root").
    /// Title-based matching fails. Fallback: detect by structural signatures —
    /// the presence of the path Edit field with AutomationId "1148" or a treeview
    /// with the SHBrowseForFolder class, plus a Select/OK button.
    /// </summary>
    private bool IsLikelyFolderPicker(AutomationElement candidate)
    {
        try
        {
            // Signature 1: AutomationId "1148" Edit field (Vista+ IFileDialog folder mode)
            var pathEdit = candidate.FindFirstDescendant(cf => cf.ByAutomationId("1148"));
            if (pathEdit != null && pathEdit.ControlType == ControlType.Edit)
                return true;

            // Signature 2: Has "Select Folder" or "Select" button by name
            var buttons = candidate.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
            foreach (var btn in buttons)
            {
                var name = SafeGet(() => btn.Name).Replace("&", "");
                if (string.Equals(name, "Select Folder", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, "Select", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(name, "Choose Folder", StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            // Signature 3: Legacy SHBrowseForFolder uses a TreeView + "Make New Folder" button
            var makeNew = candidate.FindFirstDescendant(cf => cf.ByName("Make New Folder"));
            if (makeNew != null) return true;
        }
        catch { }
        return false;
    }

    private AutomationElement? FindFolderPicker(UIA3Automation automation, string? titleFilter, HashSet<int> excelPids)
    {
        // Delegate to unified finder with kind:"folder" when no title filter given.
        // This catches UseDescriptionForTitle=true cases via IsLikelyFolderPicker.
        return FindDialog(automation, titleFilter, kind: titleFilter == null ? "folder" : null, excelPids);
    }

    private AutomationElement? FindFileDialog(UIA3Automation automation, string titleFilter, HashSet<int> excelPids)
    {
        var desktop = automation.GetDesktop();
        var children = desktop.FindAllChildren();

        foreach (var child in children)
        {
            try
            {
                if (!excelPids.Contains(child.Properties.ProcessId.Value)) continue;

                var cls = SafeGet(() => child.ClassName);
                var title = SafeGet(() => child.Name);

                if (cls != "#32770" && !cls.Contains("Dialog", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (title.Contains(titleFilter, StringComparison.OrdinalIgnoreCase))
                    return child;
            }
            catch { }
        }
        return null;
    }

    private AutomationElement? FindPathInput(AutomationElement dialog)
    {
        // Primary: AutomationId "1148" (standard Vista+ folder picker path field)
        var byId = dialog.FindFirstDescendant(cf => cf.ByAutomationId("1148"));
        if (byId != null && byId.ControlType == ControlType.Edit) return byId;

        // Secondary: look for any Edit control, prefer enabled/visible
        var edits = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Edit));
        foreach (var edit in edits)
        {
            try
            {
                if (edit.IsEnabled && !edit.IsOffscreen) return edit;
            }
            catch { }
        }
        return edits.FirstOrDefault();
    }

    private AutomationElement? FindFileNameInput(AutomationElement dialog)
    {
        // Try by name first — "File name:" or similar label
        var byName = dialog.FindFirstDescendant(cf =>
            cf.ByControlType(ControlType.Edit).And(cf.ByName("File name:")));
        if (byName != null) return byName;

        // Try by name containing "name"
        var edits = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Edit));
        foreach (var edit in edits)
        {
            var name = SafeGet(() => edit.Name);
            if (name.Contains("name", StringComparison.OrdinalIgnoreCase) && edit.IsEnabled)
                return edit;
        }

        // Fallback: first enabled visible Edit
        foreach (var edit in edits)
        {
            try
            {
                if (edit.IsEnabled && !edit.IsOffscreen) return edit;
            }
            catch { }
        }
        return edits.FirstOrDefault();
    }

    private AutomationElement? FindConfirmButton(AutomationElement dialog, string[] preferredNames)
    {
        // Primary: AutomationId "1" (standard default button)
        var byId = dialog.FindFirstDescendant(cf => cf.ByAutomationId("1"));
        if (byId != null && byId.ControlType == ControlType.Button) return byId;

        // Secondary: by name in priority order
        var buttons = dialog.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
        foreach (var preferred in preferredNames)
        {
            foreach (var btn in buttons)
            {
                var name = SafeGet(() => btn.Name).Replace("&", "");
                if (string.Equals(name, preferred, StringComparison.OrdinalIgnoreCase) && btn.IsEnabled)
                    return btn;
            }
        }
        return null;
    }

    private JsonArray DumpButtons(AutomationElement element)
    {
        var result = new JsonArray();
        try
        {
            var buttons = element.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
            foreach (var btn in buttons)
            {
                var name = SafeGet(() => btn.Name);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    result.Add(new JsonObject
                    {
                        ["name"] = name,
                        ["automation_id"] = SafeGet(() => btn.AutomationId),
                        ["enabled"] = btn.IsEnabled,
                    });
                }
            }
        }
        catch { }
        return result;
    }

    private string[] GetButtonNames(AutomationElement element)
    {
        try
        {
            return element.FindAllDescendants(cf => cf.ByControlType(ControlType.Button))
                .Select(b => SafeGet(() => b.Name))
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToArray();
        }
        catch { return []; }
    }
}
