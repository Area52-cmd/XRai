using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Nodes;
using XRai.Core;

namespace XRai.UI;

/// <summary>
/// Win32-level dialog driver. Unlike RibbonDriver.dialog.read which only sees
/// COM/UIA modal children of Excel's main window, this enumerates ALL top-level
/// windows in the Excel process using raw Win32 APIs. This is the only way to
/// see and dismiss the "Excel is waiting for another application to complete
/// an OLE action" dialog and similar native modals.
/// </summary>
public class Win32DialogDriver : IDialogWatchdog, ITimeoutDiagnostics
{
    // ── Win32 P/Invoke ────────────────────────────────────────────────

    private delegate bool EnumWindowsProc(nint hWnd, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(nint hWndParent, EnumWindowsProc lpEnumFunc, nint lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(nint hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(nint hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(nint hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(nint hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(nint hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(nint hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindowEnabled(nint hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern nint SendMessage(nint hWnd, uint msg, nint wParam, nint lParam);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(nint hWnd, uint msg, nint wParam, nint lParam);

    [DllImport("user32.dll")]
    private static extern nint SetForegroundWindow(nint hWnd);

    private const uint WM_CLOSE = 0x0010;
    private const uint WM_COMMAND = 0x0111;
    private const uint BM_CLICK = 0x00F5;
    private const uint WM_SETTEXT = 0x000C;
    private const uint WM_GETTEXT = 0x000D;
    private const uint WM_GETTEXTLENGTH = 0x000E;
    private const uint EM_SETSEL = 0x00B1;
    private const uint EM_REPLACESEL = 0x00C2;
    private const uint WM_LBUTTONDOWN = 0x0201;
    private const uint WM_LBUTTONUP = 0x0202;
    private const ushort BN_CLICKED = 0;

    [DllImport("user32.dll")]
    private static extern nint GetParent(nint hWnd);

    // Unicode string overload for WM_SETTEXT. The Windows function is SendMessageW —
    // the "SendMessageString" name I used previously doesn't exist in user32.dll.
    // Must specify EntryPoint explicitly because P/Invoke defaults to the method name
    // when looking up the native symbol.
    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW", SetLastError = true)]
    private static extern nint SendMessageSetText(nint hWnd, uint Msg, nint wParam, string lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW", SetLastError = true)]
    private static extern nint SendMessageInt(nint hWnd, uint Msg, nint wParam, nint lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageW")]
    private static extern int SendMessageGetText(nint hWnd, uint Msg, int wParam, StringBuilder lParam);

    [DllImport("user32.dll")]
    private static extern bool SetFocus(nint hWnd);

    // Known nuisance dialog patterns (title contains ... → click this button).
    // Order matters: more specific patterns first so they win over generic ones.
    private static readonly (string TitlePattern, string PreferredButton)[] NuisanceDialogs =
    [
        // OLE / server-busy family
        ("Server Busy", "Switch To..."),     // Classic OLE server-busy dialog
        ("application.complete", "OK"),      // "waiting for another application to complete an OLE action"
        ("Excel is waiting", "OK"),

        // Workbook.Open dialog family — these are the ones that block workbook.open
        // behind a NUIDialog modal and strand the STA thread.
        ("This workbook contains links", "Don't Update"),  // "Update Links" prompt → don't update
        ("update the links", "Don't Update"),
        ("file format and extension", "Yes"),              // format mismatch warning → trust and open
        ("format doesn't match", "Yes"),
        ("Protected View", "Enable Editing"),              // protected view bar promoted to dialog
        ("Security Warning", "Enable Content"),            // macro warning
        ("document recovery", "Close"),                    // recovery pane as modal
        ("read-only", "Yes"),                              // "Open as read-only?" → yes
        ("locked for editing", "Notify"),                  // file-in-use dialog → notify when available
        ("compatibility", "Continue"),                     // compat checker
        ("document was created", "Continue"),              // version warning

        // Generic catch-all — must be LAST so specific patterns above win.
        // Matches any dialog whose title is literally "Microsoft Excel"
        // (the one-liner MsgBox Excel uses for "waiting for another application").
        ("Microsoft Excel", "OK"),
    ];

    // Button-click priority for unknown NUIDialogs. When a dialog matches by
    // class name "NUIDialog" but no title pattern hits, we try these buttons
    // in order. The list is ordered "safest → most aggressive": prefer
    // buttons that DO NOT mutate data or trust untrusted content.
    private static readonly string[] NuiDialogButtonPriority =
    [
        "Don't Update",   // links prompt — preserve stale data, don't phone home
        "No",             // "save changes?" → no
        "Close",          // recovery pane
        "Cancel",         // generic cancel
        "OK",             // acknowledgment
        "Yes",            // format-mismatch / read-only recommendation
        "Continue",       // compat warnings
        "Enable Editing", // protected view (only last because it trusts content)
        "Enable Content", // macro warning (last for the same reason)
    ];

    // ── Autodismiss background thread state ──────────────────────────

    private Thread? _autodismissThread;
    private CancellationTokenSource? _autodismissCts;
    private readonly object _autodismissLock = new();
    private volatile bool _autodismissEnabled;
    private int _autodismissIntervalMs = 1000;
    private int _autodismissDismissCount;
    private string? _lastDismissedTitle;
    private DateTime? _lastDismissedAt;

    public void Register(CommandRouter router)
    {
        router.Register("win32.dialog.list", HandleList);
        router.Register("win32.dialog.dismiss", HandleDismiss);
        router.Register("win32.dialog.click", HandleClick);
        router.Register("win32.dialog.type", HandleWin32Type);
        router.Register("win32.dialog.read", HandleWin32Read);
        router.Register("excel.autodismiss", HandleAutodismiss);
        router.Register("excel.autodismiss.status", HandleAutodismissStatus);
        router.Register("dialog.auto_click", HandleAutoClick);
    }

    // ── Targeted dialog auto-click ──────────────────────────────────

    /// <summary>
    /// Pre-register a button click that fires the instant a matching dialog
    /// appears. Runs a background watcher thread that polls every 50ms for
    /// a dialog whose title contains the given substring, clicks the specified
    /// button, then stops. Use this BEFORE triggering an action that opens a
    /// modal (e.g. pane.click on a delete button that shows MessageBox.Show).
    ///
    /// The watcher runs on a separate thread from the UI thread, so it catches
    /// synchronous MessageBox.Show dialogs that block InvokeOnUI.
    ///
    /// Usage:
    ///   {"cmd":"dialog.auto_click","title":"Delete","button":"Yes","timeout":5000}
    ///   {"cmd":"pane.click","control":"DeleteButton"}
    /// </summary>
    private string HandleAutoClick(JsonObject args)
    {
        var title = args["title"]?.GetValue<string>();
        var button = args["button"]?.GetValue<string>()
            ?? throw new ArgumentException("dialog.auto_click requires 'button'");
        var timeoutMs = args["timeout"]?.GetValue<int>() ?? 10000;
        var pollMs = args["poll_ms"]?.GetValue<int>() ?? 50;

        // Start a background watcher thread
        var thread = new Thread(() =>
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    var windows = EnumerateExcelWindows();
                    foreach (var w in windows.Where(w => !w.IsMainWindow && IsLikelyDialog(w)))
                    {
                        // Match by title if provided, otherwise match any dialog
                        if (title != null &&
                            !w.Title.Contains(title, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var buttons = EnumerateButtons(w.Handle);
                        var target = buttons.FirstOrDefault(b =>
                            b.Text.Replace("&", "").Equals(button, StringComparison.OrdinalIgnoreCase));

                        if (target.Handle != nint.Zero)
                        {
                            ClickWin32Button(target.Handle);
                            _lastAutoClickTitle = w.Title;
                            _lastAutoClickButton = target.Text;
                            _lastAutoClickAt = DateTime.UtcNow;
                            return; // Done — dialog found and clicked
                        }
                    }
                }
                catch { }

                Thread.Sleep(pollMs);
            }
        })
        {
            IsBackground = true,
            Name = "xrai-dialog-auto-click"
        };
        thread.Start();

        return Response.Ok(new
        {
            watching = true,
            title = title ?? "(any dialog)",
            button,
            timeout_ms = timeoutMs,
            poll_ms = pollMs,
            hint = "Watcher started. Now trigger the action that opens the dialog."
        });
    }

    private volatile string? _lastAutoClickTitle;
    private volatile string? _lastAutoClickButton;
    private DateTime? _lastAutoClickAt;

    // ── Window enumeration ───────────────────────────────────────────

    private List<WindowInfo> EnumerateExcelWindows()
    {
        var excelPids = Process.GetProcessesByName("EXCEL")
            .Select(p => (uint)p.Id)
            .ToHashSet();

        if (excelPids.Count == 0) return [];

        var result = new List<WindowInfo>();

        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd)) return true;
            GetWindowThreadProcessId(hWnd, out uint pid);
            if (!excelPids.Contains(pid)) return true;

            var title = GetTitle(hWnd);
            var className = GetClass(hWnd);

            // Skip the main Excel XLMAIN windows — those are the workbook windows
            // We want dialogs, tool windows, message boxes.
            // Common Win32 dialog classes: #32770 (standard), NUIDialog, bosa_sdm_XL9
            result.Add(new WindowInfo
            {
                Handle = hWnd,
                Title = title,
                ClassName = className,
                Pid = pid,
                IsMainWindow = className == "XLMAIN",
                Enabled = IsWindowEnabled(hWnd),
            });

            return true;
        }, 0);

        return result;
    }

    private static string GetTitle(nint hWnd)
    {
        int len = GetWindowTextLength(hWnd);
        if (len == 0) return "";
        var sb = new StringBuilder(len + 1);
        GetWindowText(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private static string GetClass(nint hWnd)
    {
        var sb = new StringBuilder(256);
        GetClassName(hWnd, sb, sb.Capacity);
        return sb.ToString();
    }

    private List<ButtonInfo> EnumerateButtons(nint hWndParent)
    {
        var buttons = new List<ButtonInfo>();
        EnumChildWindows(hWndParent, (hWnd, _) =>
        {
            var className = GetClass(hWnd);
            // Match:
            //   "Button"                                  — native Win32
            //   "WindowsForms10.BUTTON.app.0.141b42a_r9_ad1" — WinForms Button (any suffix)
            //   Any class containing ".BUTTON." (case-insensitive) — generalized WinForms
            bool isButton =
                className.Equals("Button", StringComparison.OrdinalIgnoreCase) ||
                className.Contains(".BUTTON.", StringComparison.OrdinalIgnoreCase) ||
                className.StartsWith("WindowsForms", StringComparison.OrdinalIgnoreCase) &&
                    className.Contains("BUTTON", StringComparison.OrdinalIgnoreCase);

            if (isButton)
            {
                var text = GetTitle(hWnd);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    buttons.Add(new ButtonInfo { Handle = hWnd, Text = text });
                }
            }
            return true;
        }, 0);
        return buttons;
    }

    // ── Command handlers ─────────────────────────────────────────────

    private string HandleList(JsonObject args)
    {
        var windows = EnumerateExcelWindows();
        var result = new JsonArray();

        foreach (var w in windows)
        {
            if (w.IsMainWindow) continue; // skip main Excel window, focus on dialogs

            var buttons = EnumerateButtons(w.Handle);
            var btnArray = new JsonArray();
            foreach (var b in buttons)
                btnArray.Add(b.Text);

            result.Add(new JsonObject
            {
                ["hwnd"] = w.Handle.ToInt64(),
                ["title"] = w.Title,
                ["class"] = w.ClassName,
                ["pid"] = (long)w.Pid,
                ["enabled"] = w.Enabled,
                ["buttons"] = btnArray,
                ["is_likely_dialog"] = IsLikelyDialog(w),
            });
        }

        return Response.Ok(new { windows = result, count = result.Count });
    }

    private string HandleDismiss(JsonObject args)
    {
        var windows = EnumerateExcelWindows();
        var dialogs = windows.Where(w => !w.IsMainWindow && IsLikelyDialog(w)).ToList();

        if (dialogs.Count == 0)
            return Response.Ok(new { dismissed = false, reason = "No Win32 dialogs detected in Excel process" });

        var dismissed = new JsonArray();
        var priorities = new[] { "OK", "Yes", "Continue", "Retry", "Switch To...", "Close", "Cancel", "No" };

        foreach (var dlg in dialogs)
        {
            var buttons = EnumerateButtons(dlg.Handle);
            ButtonInfo? target = null;

            foreach (var preferred in priorities)
            {
                target = buttons.FirstOrDefault(b =>
                    b.Text.Replace("&", "").Equals(preferred, StringComparison.OrdinalIgnoreCase));
                if (target != null) break;
            }

            if (target != null)
            {
                SendMessage(target.Value.Handle, BM_CLICK, 0, 0);
                dismissed.Add(new JsonObject
                {
                    ["hwnd"] = dlg.Handle.ToInt64(),
                    ["title"] = dlg.Title,
                    ["clicked"] = target.Value.Text,
                });
            }
            else
            {
                // Last resort: send WM_CLOSE
                PostMessage(dlg.Handle, WM_CLOSE, 0, 0);
                dismissed.Add(new JsonObject
                {
                    ["hwnd"] = dlg.Handle.ToInt64(),
                    ["title"] = dlg.Title,
                    ["clicked"] = "(WM_CLOSE)",
                });
            }
        }

        return Response.Ok(new { dismissed = true, count = dismissed.Count, details = dismissed });
    }

    /// <summary>
    /// Public entry point so RibbonDriver's dialog.click can fall through to the
    /// Win32 enumeration path without re-instantiating a driver. Returns the same
    /// JSON shape HandleClick returns on success, or null if no matching button
    /// was found (caller should synthesize its own error so it can aggregate
    /// information from both the UIA and Win32 tiers).
    /// </summary>
    public string? TryClickButton(JsonObject args)
    {
        var result = HandleClick(args);
        // HandleClick returns Response.Error(...) for "not found" — detect that by
        // parsing the ok field and returning null so the caller can fall through.
        try
        {
            var node = System.Text.Json.Nodes.JsonNode.Parse(result);
            if (node?["ok"]?.GetValue<bool>() == true) return result;
            return null;
        }
        catch { return null; }
    }

    private string HandleClick(JsonObject args)
    {
        var buttonText = args["button"]?.GetValue<string>()
            ?? throw new ArgumentException("win32.dialog.click requires 'button'");
        var titleFilter = args["title"]?.GetValue<string>();

        var windows = EnumerateExcelWindows();
        foreach (var w in windows.Where(w => !w.IsMainWindow))
        {
            if (titleFilter != null &&
                !w.Title.Contains(titleFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            var buttons = EnumerateButtons(w.Handle);
            var target = buttons.FirstOrDefault(b =>
                b.Text.Replace("&", "").Equals(buttonText, StringComparison.OrdinalIgnoreCase));

            if (target.Handle != nint.Zero)
            {
                SendMessage(target.Handle, BM_CLICK, 0, 0);
                return Response.Ok(new
                {
                    clicked = true,
                    dialog_title = w.Title,
                    button = buttonText
                });
            }
        }

        return Response.Error($"Win32 button not found: '{buttonText}'" +
            (titleFilter != null ? $" (in dialog matching '{titleFilter}')" : ""));
    }

    // ── Type into Win32/WinForms edit controls inside dialogs ────────

    /// <summary>
    /// Type text into an Edit (or ComboBox) control inside a Win32 dialog.
    /// Matches the target dialog by optional title substring, then locates the
    /// Edit control by one of:
    ///   - control_id (Win32 control ID, e.g. 1148 for file-dialog path field)
    ///   - index (zero-based among Edit controls, default 0 = first edit)
    ///   - class_name (e.g. "Edit" for WinForms, "RICHEDIT50W" for rich edit)
    /// Sends WM_SETTEXT to set the text directly, which works for native Win32
    /// Edit controls and WinForms TextBox/ComboBox alike without focus games.
    /// </summary>
    private string HandleWin32Type(JsonObject args)
    {
        var text = args["text"]?.GetValue<string>()
            ?? throw new ArgumentException("win32.dialog.type requires 'text'");
        var titleFilter = args["title"]?.GetValue<string>();
        var controlId = args["control_id"]?.GetValue<int>();
        var index = args["index"]?.GetValue<int>() ?? 0;
        var className = args["class_name"]?.GetValue<string>();
        var submit = args["submit"]?.GetValue<bool>() ?? false;

        var windows = EnumerateExcelWindows();
        var dialogs = windows.Where(w => !w.IsMainWindow && IsLikelyDialog(w)).ToList();

        if (titleFilter != null)
            dialogs = dialogs.Where(w => w.Title.Contains(titleFilter, StringComparison.OrdinalIgnoreCase)).ToList();

        if (dialogs.Count == 0)
            return Response.Error("No matching Win32 dialog found" +
                (titleFilter != null ? $" (title contains '{titleFilter}')" : "") +
                ". Use {\"cmd\":\"win32.dialog.list\"} to see what's open.");

        var dialog = dialogs[0];

        // Enumerate all Edit / ComboBox descendants
        var edits = EnumerateEditControls(dialog.Handle);

        if (edits.Count == 0)
            return Response.Error($"No Edit/ComboBox controls found in dialog '{dialog.Title}'. " +
                "Use {\"cmd\":\"win32.dialog.read\",\"title\":\"...\"} to inspect the dialog structure.");

        // Pick the target edit
        EditControlInfo? target = null;

        if (controlId.HasValue)
        {
            target = edits.FirstOrDefault(e => e.ControlId == controlId.Value);
            if (target == null)
                return Response.Error($"No edit control with control_id={controlId} in dialog '{dialog.Title}'. " +
                    $"Available IDs: {string.Join(", ", edits.Select(e => e.ControlId))}");
        }
        else if (className != null)
        {
            target = edits.FirstOrDefault(e =>
                e.ClassName.Contains(className, StringComparison.OrdinalIgnoreCase));
            if (target == null)
                return Response.Error($"No edit control with class containing '{className}'. " +
                    $"Available classes: {string.Join(", ", edits.Select(e => e.ClassName).Distinct())}");
        }
        else
        {
            if (index < 0 || index >= edits.Count)
                return Response.Error($"index {index} out of range (found {edits.Count} edit controls). " +
                    "Use {\"cmd\":\"win32.dialog.read\"} to inspect.");
            target = edits[index];
        }

        var targetInfo = target.Value;

        // Focus the control first (some dialogs validate on focus change)
        SetFocus(targetInfo.Handle);

        // Send WM_SETTEXT with the new value. This is the reliable way to set
        // text on a native Edit or WinForms TextBox — no keystroke simulation needed.
        SendMessageSetText(targetInfo.Handle, WM_SETTEXT, 0, text);

        // If submit is requested, click the default button using a cascade of
        // reliable techniques. Only report submitted:true if a click actually
        // happened — previously this lied with submitted=submit which reported
        // true even when the button was never found.
        string? submitClicked = null;
        bool actuallySubmitted = false;
        if (submit)
        {
            var buttons = EnumerateButtons(dialog.Handle);
            var priorities = new[] { "OK", "Save", "Open", "Yes", "Apply", "Select Folder", "Select" };
            foreach (var p in priorities)
            {
                var btn = buttons.FirstOrDefault(b =>
                    b.Text.Replace("&", "").Equals(p, StringComparison.OrdinalIgnoreCase));
                if (btn.Handle != nint.Zero)
                {
                    if (ClickWin32Button(btn.Handle))
                    {
                        submitClicked = btn.Text;
                        actuallySubmitted = true;
                    }
                    break;
                }
            }
        }

        string? submitError = null;
        if (submit && !actuallySubmitted)
        {
            submitError = "Text was set successfully but submit:true failed — no OK/Save/Open-class button could be clicked. " +
                "Available buttons: " +
                string.Join(", ", EnumerateButtons(dialog.Handle).Select(b => b.Text)) +
                ". Follow with {\"cmd\":\"dialog.click\",\"button\":\"<name>\"} manually.";
        }

        return Response.Ok(new
        {
            typed = true,
            dialog_title = dialog.Title,
            control_id = targetInfo.ControlId,
            class_name = targetInfo.ClassName,
            text_length = text.Length,
            submit_requested = submit,
            submitted = actuallySubmitted,
            submit_button = submitClicked,
            submit_error = submitError,
        });
    }

    /// <summary>
    /// Click a Win32 button with a three-tier cascade:
    ///   1. BM_CLICK — works for native Win32 Button controls
    ///   2. WM_COMMAND(BN_CLICKED) to the parent — how real mouse clicks reach the dialog
    ///   3. WM_LBUTTONDOWN + WM_LBUTTONUP — synthetic mouse events as last resort
    /// Returns true if any technique reported success (SendMessage returned 0 is fine,
    /// it just means the button processed the message).
    /// </summary>
    private bool ClickWin32Button(nint hwndButton)
    {
        try
        {
            // 1. Direct BM_CLICK (fastest, works for native buttons)
            SendMessageInt(hwndButton, BM_CLICK, 0, 0);

            // 2. WM_COMMAND to the parent with BN_CLICKED notification.
            // This is the message Windows sends when a real mouse click reaches a button.
            // WinForms buttons respond to this even when BM_CLICK is ignored.
            var parent = GetParent(hwndButton);
            if (parent != nint.Zero)
            {
                var controlId = GetDlgCtrlID(hwndButton);
                // wParam: HIWORD = BN_CLICKED, LOWORD = control id
                nint wParam = (nint)((uint)((BN_CLICKED << 16) | (ushort)controlId));
                SendMessageInt(parent, WM_COMMAND, wParam, hwndButton);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Read the full structure of a Win32 dialog: title, class, all edit controls
    /// (with current text values), all buttons. Use this to discover what's inside
    /// a dialog before driving it with win32.dialog.type / win32.dialog.click.
    /// </summary>
    private string HandleWin32Read(JsonObject args)
    {
        var titleFilter = args["title"]?.GetValue<string>();
        var windows = EnumerateExcelWindows();
        var dialogs = windows.Where(w => !w.IsMainWindow && IsLikelyDialog(w)).ToList();

        if (titleFilter != null)
            dialogs = dialogs.Where(w => w.Title.Contains(titleFilter, StringComparison.OrdinalIgnoreCase)).ToList();

        if (dialogs.Count == 0)
            return Response.Error("No matching Win32 dialog found" +
                (titleFilter != null ? $" (title contains '{titleFilter}')" : ""));

        var dialog = dialogs[0];
        var edits = EnumerateEditControls(dialog.Handle);
        var buttons = EnumerateButtons(dialog.Handle);

        var editsJson = new JsonArray();
        foreach (var e in edits)
        {
            editsJson.Add(new JsonObject
            {
                ["control_id"] = e.ControlId,
                ["class_name"] = e.ClassName,
                ["hwnd"] = e.Handle.ToInt64(),
                ["current_text"] = e.Text,
            });
        }

        var buttonsJson = new JsonArray();
        foreach (var b in buttons)
            buttonsJson.Add(b.Text);

        return Response.Ok(new
        {
            title = dialog.Title,
            class_name = dialog.ClassName,
            hwnd = dialog.Handle.ToInt64(),
            edit_controls = editsJson,
            buttons = buttonsJson,
        });
    }

    [DllImport("user32.dll")]
    private static extern int GetDlgCtrlID(nint hWnd);

    private List<EditControlInfo> EnumerateEditControls(nint hWndParent)
    {
        var result = new List<EditControlInfo>();
        EnumChildWindows(hWndParent, (hWnd, _) =>
        {
            var cls = GetClass(hWnd);
            // WinForms TextBox → "WindowsForms10.EDIT.*"
            // Native Edit → "Edit"
            // Rich Edit → "RICHEDIT50W" / "RichEdit20W"
            // ComboBoxEx32 → "ComboBoxEx32"
            if (cls.StartsWith("Edit", StringComparison.OrdinalIgnoreCase) ||
                cls.StartsWith("RichEdit", StringComparison.OrdinalIgnoreCase) ||
                cls.StartsWith("RICHEDIT", StringComparison.OrdinalIgnoreCase) ||
                cls.Contains(".EDIT.", StringComparison.OrdinalIgnoreCase) ||
                cls.Contains(".TextBox.", StringComparison.OrdinalIgnoreCase) ||
                cls.Equals("ComboBox", StringComparison.OrdinalIgnoreCase) ||
                cls.StartsWith("ComboBoxEx", StringComparison.OrdinalIgnoreCase))
            {
                result.Add(new EditControlInfo
                {
                    Handle = hWnd,
                    ClassName = cls,
                    ControlId = GetDlgCtrlID(hWnd),
                    Text = GetEditText(hWnd),
                });
            }
            return true;
        }, 0);
        return result;
    }

    private static string GetEditText(nint hWnd)
    {
        try
        {
            var len = (int)SendMessageInt(hWnd, WM_GETTEXTLENGTH, 0, 0);
            if (len <= 0) return "";
            var sb = new StringBuilder(len + 1);
            SendMessageGetText(hWnd, WM_GETTEXT, len + 1, sb);
            return sb.ToString();
        }
        catch { return ""; }
    }

    private struct EditControlInfo
    {
        public nint Handle;
        public string ClassName;
        public int ControlId;
        public string Text;
    }

    // ── Autodismiss background thread ────────────────────────────────

    private string HandleAutodismiss(JsonObject args)
    {
        bool enabled = args["enabled"]?.GetValue<bool>() ?? true;
        int intervalMs = args["interval_ms"]?.GetValue<int>() ?? 1000;

        if (enabled)
        {
            bool wasRunning = _autodismissEnabled;
            EnableWatchdog(intervalMs);
            return Response.Ok(new
            {
                enabled = true,
                already_running = wasRunning,
                interval_ms = _autodismissIntervalMs,
                message = wasRunning
                    ? "Watchdog already running"
                    : "Background thread polling for NUIDialog / OLE-wait / server-busy / update-links dialogs"
            });
        }
        else
        {
            DisableWatchdog();
            return Response.Ok(new { enabled = false, stopped = true });
        }
    }

    private string HandleAutodismissStatus(JsonObject args)
    {
        return Response.Ok(new
        {
            enabled = _autodismissEnabled,
            interval_ms = _autodismissIntervalMs,
            total_dismissed = _autodismissDismissCount,
            last_dismissed_title = _lastDismissedTitle,
            last_dismissed_at = _lastDismissedAt?.ToString("o"),
        });
    }

    private void AutodismissLoop()
    {
        var token = _autodismissCts?.Token ?? CancellationToken.None;
        while (_autodismissEnabled && !token.IsCancellationRequested)
        {
            try { DismissNuisanceWindows(); }
            catch { /* swallow — keep polling */ }

            try { Thread.Sleep(_autodismissIntervalMs); } catch { break; }
        }
    }

    /// <summary>
    /// Single pass over Excel's top-level windows. Any window that is a
    /// NUIDialog (by class) or matches a known nuisance title pattern gets
    /// its safest button clicked. Returns the number of dialogs dismissed
    /// on this pass. Safe to call from any thread (all Win32 windowing
    /// APIs used are thread-agnostic).
    /// </summary>
    private int DismissNuisanceWindows()
    {
        int dismissed = 0;
        var windows = EnumerateExcelWindows();
        foreach (var w in windows.Where(w => !w.IsMainWindow && IsLikelyDialog(w)))
        {
            // Tier 1: title matches a specific nuisance pattern → use the
            // pattern's preferred button (e.g. "Don't Update" for update-links).
            string? preferredButton = null;
            foreach (var (pattern, btn) in NuisanceDialogs)
            {
                if (w.Title.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                {
                    preferredButton = btn;
                    break;
                }
            }

            // Tier 2: unmatched title but the class is NUIDialog — still
            // nuisance (workbook.open never legitimately pops a NUIDialog at
            // us during automation). Try the safe-priority cascade.
            bool isNuiDialog = w.ClassName.Equals("NUIDialog", StringComparison.OrdinalIgnoreCase);

            if (preferredButton == null && !isNuiDialog) continue;

            var buttons = EnumerateButtons(w.Handle);
            if (buttons.Count == 0) continue;

            nint targetHandle = nint.Zero;
            string? targetText = null;

            // First try the pattern-preferred button if we have one.
            if (preferredButton != null)
            {
                var match = buttons.FirstOrDefault(b =>
                    b.Text.Replace("&", "").Equals(preferredButton, StringComparison.OrdinalIgnoreCase));
                if (match.Handle != nint.Zero)
                {
                    targetHandle = match.Handle;
                    targetText = match.Text;
                }
            }

            // Fall back to the NUIDialog safe-priority cascade.
            if (targetHandle == nint.Zero)
            {
                foreach (var candidate in NuiDialogButtonPriority)
                {
                    var match = buttons.FirstOrDefault(b =>
                        b.Text.Replace("&", "").Equals(candidate, StringComparison.OrdinalIgnoreCase));
                    if (match.Handle != nint.Zero)
                    {
                        targetHandle = match.Handle;
                        targetText = match.Text;
                        break;
                    }
                }
            }

            // Last resort: first available button.
            if (targetHandle == nint.Zero)
            {
                targetHandle = buttons[0].Handle;
                targetText = buttons[0].Text;
            }

            if (targetHandle != nint.Zero)
            {
                // Click via the three-tier cascade so WinForms buttons on
                // NUIDialogs respond even when BM_CLICK is ignored.
                ClickWin32Button(targetHandle);
                dismissed++;
                _autodismissDismissCount++;
                _lastDismissedTitle = $"{w.Title} [{w.ClassName}] → {targetText}";
                _lastDismissedAt = DateTime.UtcNow;
            }
        }
        return dismissed;
    }

    // ── IDialogWatchdog (in-process API, no JSON round-trip) ─────────

    /// <summary>
    /// Start the background dismiss loop. Idempotent.
    /// </summary>
    public void EnableWatchdog(int intervalMs = 250)
    {
        lock (_autodismissLock)
        {
            _autodismissIntervalMs = Math.Max(100, intervalMs);
            if (_autodismissEnabled) return;
            _autodismissEnabled = true;
            _autodismissCts = new CancellationTokenSource();
            _autodismissThread = new Thread(AutodismissLoop)
            {
                IsBackground = true,
                Name = "xrai-dialog-watchdog"
            };
            _autodismissThread.Start();
        }
    }

    /// <summary>
    /// Stop the background dismiss loop. Idempotent.
    /// </summary>
    public void DisableWatchdog()
    {
        lock (_autodismissLock)
        {
            if (!_autodismissEnabled) return;
            _autodismissEnabled = false;
            try { _autodismissCts?.Cancel(); } catch { }
            _autodismissThread = null;
        }
    }

    /// <summary>
    /// One-shot pass. Returns the number of dialogs dismissed.
    /// </summary>
    public int DismissOnce()
    {
        try { return DismissNuisanceWindows(); }
        catch { return 0; }
    }

    // ── ITimeoutDiagnostics ─────────────────────────────────────────

    /// <summary>
    /// Snapshot of all non-main Excel windows for inclusion in timeout error
    /// responses. Thread-safe (Win32 EnumWindows is apartment-agnostic).
    /// </summary>
    public object? GetDialogSnapshot()
    {
        try
        {
            var windows = EnumerateExcelWindows();
            var dialogs = windows.Where(w => !w.IsMainWindow && IsLikelyDialog(w)).ToList();
            if (dialogs.Count == 0) return null;

            var result = new List<object>();
            foreach (var d in dialogs)
            {
                var buttons = EnumerateButtons(d.Handle);
                var edits = EnumerateEditControls(d.Handle);
                result.Add(new
                {
                    title = d.Title,
                    class_name = d.ClassName,
                    buttons = buttons.Select(b => b.Text).ToArray(),
                    edit_controls = edits.Select(e => new
                    {
                        control_id = e.ControlId,
                        class_name = e.ClassName,
                        current_text = e.Text,
                    }).ToArray(),
                });
            }
            return result;
        }
        catch { return null; }
    }

    // ── Helpers ──────────────────────────────────────────────────────

    private static bool IsLikelyDialog(WindowInfo w)
    {
        if (w.IsMainWindow) return false;
        // Standard Win32 dialog class = "#32770"
        // Excel uses "bosa_sdm_XL9" and "NUIDialog" for some dialogs
        // Task dialogs use "TaskDialog" and "DirectUIHWND"
        return w.ClassName == "#32770"
            || w.ClassName.Contains("Dialog", StringComparison.OrdinalIgnoreCase)
            || w.ClassName.StartsWith("bosa_sdm_")
            || w.ClassName == "NUIDialog"
            || !string.IsNullOrEmpty(w.Title);
    }

    private struct WindowInfo
    {
        public nint Handle;
        public string Title;
        public string ClassName;
        public uint Pid;
        public bool IsMainWindow;
        public bool Enabled;
    }

    private struct ButtonInfo
    {
        public nint Handle;
        public string Text;
    }
}
