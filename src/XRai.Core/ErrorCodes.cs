namespace XRai.Core;

/// <summary>
/// Stable error code taxonomy for XRai responses. Codes are stable contracts —
/// their string values must not change once shipped. New error categories get
/// new codes, not renamed codes.
///
/// Format: XRAI_{CATEGORY}_{SPECIFIC}
/// Example: XRAI_PANE_CONTROL_NOT_FOUND, XRAI_STA_STUCK, XRAI_COM_NOT_ATTACHED
///
/// Every code has an entry in ErrorHints with a hint + doc URL.
/// </summary>
public static class ErrorCodes
{
    // === Session / attach / lifecycle ===
    public const string ComNotAttached = "XRAI_COM_NOT_ATTACHED";
    public const string ComAttachFailed = "XRAI_COM_ATTACH_FAILED";
    public const string ComServerBusy = "XRAI_COM_SERVER_BUSY";         // RPC_E_SERVERCALL_RETRYLATER
    public const string ComServerUnavailable = "XRAI_COM_SERVER_UNAVAILABLE"; // 0x800706BA
    public const string ExcelNotRunning = "XRAI_EXCEL_NOT_RUNNING";
    public const string NoActiveWorkbook = "XRAI_NO_ACTIVE_WORKBOOK";
    public const string NoActiveSheet = "XRAI_NO_ACTIVE_SHEET";

    // === STA worker ===
    public const string StaStuck = "XRAI_STA_STUCK";
    public const string StaTimeout = "XRAI_STA_TIMEOUT";
    public const string StaRecycleFailed = "XRAI_STA_RECYCLE_FAILED";

    // === Hooks pipe ===
    public const string HooksNotConnected = "XRAI_HOOKS_NOT_CONNECTED";
    public const string HooksPipeNotFound = "XRAI_HOOKS_PIPE_NOT_FOUND";
    public const string HooksAuthFailed = "XRAI_HOOKS_AUTH_FAILED";
    public const string HooksDied = "XRAI_HOOKS_DIED";
    public const string HooksStale = "XRAI_HOOKS_STALE";

    // === Pane / hooks controls ===
    public const string PaneControlNotFound = "XRAI_PANE_CONTROL_NOT_FOUND";
    public const string PaneNoRoot = "XRAI_PANE_NO_ROOT";
    public const string PaneExposeNotCalled = "XRAI_PANE_EXPOSE_NOT_CALLED";
    public const string PaneNotLoaded = "XRAI_PANE_NOT_LOADED";
    public const string PaneClickSilentNoOp = "XRAI_PANE_CLICK_SILENT_NO_OP";
    public const string PaneClickHandlerThrew = "XRAI_PANE_CLICK_HANDLER_THREW";

    // === Commands / routing ===
    public const string UnknownCommand = "XRAI_UNKNOWN_COMMAND";
    public const string MissingArgument = "XRAI_MISSING_ARGUMENT";
    public const string InvalidJson = "XRAI_INVALID_JSON";
    public const string InvalidArgument = "XRAI_INVALID_ARGUMENT";

    // === Cells / ranges ===
    public const string InvalidRange = "XRAI_INVALID_RANGE";
    public const string SheetNotFound = "XRAI_SHEET_NOT_FOUND";
    public const string NamedRangeNotFound = "XRAI_NAMED_RANGE_NOT_FOUND";

    // === Build / rebuild ===
    public const string BuildFailed = "XRAI_BUILD_FAILED";
    public const string ProjectNotFound = "XRAI_PROJECT_NOT_FOUND";
    public const string RebuildTimeout = "XRAI_REBUILD_TIMEOUT";

    // === Dialogs ===
    public const string DialogNotFound = "XRAI_DIALOG_NOT_FOUND";
    public const string FolderPickerNotFound = "XRAI_FOLDER_PICKER_NOT_FOUND";

    // === VBA ===
    public const string VbaAccessDenied = "XRAI_VBA_ACCESS_DENIED";
    public const string VbaNotSupported = "XRAI_VBA_NOT_SUPPORTED";

    // === Power Query / DAX ===
    public const string PowerQueryNotAvailable = "XRAI_POWER_QUERY_NOT_AVAILABLE";
    public const string DaxNotAvailable = "XRAI_DAX_NOT_AVAILABLE";

    // === Version / capability ===
    public const string NotSupportedInExcelVersion = "XRAI_NOT_SUPPORTED_IN_EXCEL_VERSION";
    public const string VersionMismatch = "XRAI_VERSION_MISMATCH";

    // === Generic ===
    public const string InternalError = "XRAI_INTERNAL_ERROR";
    public const string NotImplemented = "XRAI_NOT_IMPLEMENTED";
    public const string Timeout = "XRAI_TIMEOUT";
}
