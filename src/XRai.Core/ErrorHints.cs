namespace XRai.Core;

/// <summary>
/// Auto-attaches a short, actionable hint to error responses based on
/// pattern matching the error message or exception type. The hint is
/// a single line telling the user (or AI agent) the most likely cause
/// and the fastest recovery.
///
/// Every error response that flows through Response.Error / ErrorFromException
/// gets checked against this catalog.
/// </summary>
public static class ErrorHints
{
    // Pattern -> hint. First match wins. Ordered by specificity — most specific first.
    private static readonly (string Pattern, string Hint)[] Hints =
    [
        // ─── STA / threading / daemon ──────────────────────────────
        ("timed out after", "STA worker may be stuck. Try {\"cmd\":\"sta.reset\"} to recycle the STA thread, or restart the daemon with 'XRai.Tool.exe daemon-stop'."),
        ("handler is still running on the STA thread", "The previous command is still blocked inside Excel. Run {\"cmd\":\"sta.reset\"} to force-recycle the STA thread."),
        ("IsTainted", "The command router is tainted from a prior timeout. Restart the process or run 'XRai.Tool.exe daemon-stop'."),

        // ─── COM / attach / Excel availability ─────────────────────
        ("RPC_E_SERVERCALL_RETRYLATER", "Excel is busy (recalc, dialog, ribbon animation). Wait 1-2 seconds and retry, or ensure IOleMessageFilter is registered (check status.filter_registered)."),
        ("RPC server is unavailable", "Excel crashed or is not running. Run {\"cmd\":\"connect\"} which will relaunch if needed."),
        ("MK_E_UNAVAILABLE", "Excel is not yet ready in the Running Object Table. Run {\"cmd\":\"wait\"} then {\"cmd\":\"connect\"}."),
        ("GetActiveObject failed", "Excel is not running. Launch Excel first: 'start excel' then {\"cmd\":\"connect\"}."),
        ("No active workbook", "Excel is on the start screen. Use {\"cmd\":\"connect\"} which auto-creates a workbook, or {\"cmd\":\"ensure.workbook\"}."),
        ("No Excel process found", "Excel is not running. Launch it: 'start excel' then {\"cmd\":\"connect\"}."),
        ("Not attached", "XRai is not attached to Excel. Run {\"cmd\":\"connect\"} first."),
        ("Already attached", "This is safe to ignore in most cases — the wait command is now inert inside batches."),

        // ─── Hooks pipe ────────────────────────────────────────────
        ("Hooks pipe not connected", "The add-in's XRai.Hooks NuGet may not be loaded. Verify: (1) XRai.Hooks is referenced in the add-in's .csproj (use Version=\"1.0.0-*\" to track pre-release builds), (2) Pilot.Start() is called in IExcelAddIn.AutoOpen(), (3) the .xll is loaded in Excel. Check {\"cmd\":\"log.read\",\"source\":\"startup\"} for Pilot.Start failures, and status.hooks_pipe for the expected pipe name."),
        ("hooks: false", "The add-in is loaded but XRai.Hooks is not. Check {\"cmd\":\"log.read\",\"source\":\"startup\"} for Pilot.Start() failures, then ensure Pilot.Start() is called in AutoOpen() and XRai.Hooks NuGet is referenced (Version=\"1.0.0-*\" for pre-release builds)."),
        ("pipe server died", "The add-in process crashed or the hooks pipe was closed. Run {\"cmd\":\"rebuild\"} to rebuild and reconnect."),

        // ─── Control discovery (pane.*) ────────────────────────────
        ("Control not found", "The named control does not exist in the exposed visual tree. Check: (1) the XAML has x:Name=\"...\" on this control, (2) Pilot.Expose(pane) is called AFTER the pane is constructed, (3) run {\"cmd\":\"pane\"} to list all discovered control names."),
        ("empty controls array", "Pilot.Expose was never called, or was called before the visual tree was built. Call Pilot.Expose(taskPane) inside the pane's Loaded event or after InitializeComponent()."),
        ("Silent no-op", "The button has a Command binding but CanExecute returned false. Check: (1) the bound Command's CanExecute conditions, (2) required ViewModel state, (3) focus — some Commands require the control to be focused first."),

        // ─── Commands / routing ────────────────────────────────────
        ("Unknown command", "This command is not registered. Run {\"cmd\":\"commands\"} to see all available commands, or {\"cmd\":\"help\"} for details."),
        ("JSON parse error", "The JSON command is malformed. Common cause: unescaped backslashes in Windows paths. Use double backslashes \"C:\\\\Temp\" or forward slashes \"C:/Temp\"."),

        // ─── Build / rebuild ───────────────────────────────────────
        ("Build failed", "The .csproj failed to compile. Check the stderr output above for specific C#/XAML errors. Fix them in the source, then re-run {\"cmd\":\"rebuild\"}."),
        ("Project not found", "The csproj path doesn't exist. Pass an absolute path: {\"cmd\":\"rebuild\",\"project\":\"D:\\\\Code\\\\MyAddin\\\\MyAddin.csproj\"}."),

        // ─── Dialogs ────────────────────────────────────────────────
        ("No dialog is open", "No modal dialog found by UIA or Win32 enumeration. If a dialog IS visible, run {\"cmd\":\"win32.dialog.list\"} to see what XRai sees. The dialog may have been dismissed already."),
        ("No folder picker dialog found", "The folder picker is not open yet. Trigger the action that opens it first (e.g. click the Browse button), then run {\"cmd\":\"folder.dialog.set_path\"}."),
        ("No matching Win32 dialog found", "The dialog title or class doesn't match. Run {\"cmd\":\"win32.dialog.list\"} to see all open dialogs and their titles."),

        // ─── Build version / staleness ────────────────────────────
        ("hooks_stale", "The add-in is running old XRai.Hooks code. Run {\"cmd\":\"rebuild\"} to pull the latest XRai.Hooks from the skill's NuGet feed."),
        ("Daemon build", "The running XRai daemon is older than the CLI binary. Run 'XRai.Tool.exe daemon-stop' and restart if needed."),
    ];

    /// <summary>
    /// Return the canonical docs URL for a stable error code. Used by
    /// Response.Error / ErrorFromException so every error response points
    /// users at a stable landing page they can bookmark and share.
    ///
    /// Returns null when the code is null, empty, or doesn't look like an
    /// XRai code (so we don't linkrot on ad-hoc strings).
    /// </summary>
    public static string? GetDocsUrl(string? code)
    {
        if (string.IsNullOrEmpty(code)) return null;
        if (!code.StartsWith("XRAI_", StringComparison.Ordinal)) return null;
        return $"https://xrai.dev/errors/{code}";
    }

    /// <summary>
    /// Look up a hint for an error message. Returns null if no pattern matches.
    /// </summary>
    public static string? GetHint(string errorMessage, Exception? exception = null)
    {
        if (string.IsNullOrEmpty(errorMessage)) return null;

        // Check exception type first for COM-specific errors
        if (exception is System.Runtime.InteropServices.COMException comEx)
        {
            var hex = (uint)comEx.HResult;
            if (hex == 0x80010108) return "Excel closed while a command was in flight. Run {\"cmd\":\"connect\"} to reattach.";
            if (hex == 0x800706BA) return "Excel is not responding (RPC server unavailable). Run {\"cmd\":\"connect\"} to reattach.";
            if (hex == 0x800401E3) return "Excel not in ROT yet. Wait a moment and try {\"cmd\":\"connect\"}.";
        }

        foreach (var (pattern, hint) in Hints)
        {
            if (errorMessage.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                return hint;
        }

        return null;
    }
}
