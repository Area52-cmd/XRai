namespace XRai.Core;

/// <summary>
/// Provides diagnostic context when a command times out. Implemented by the
/// Win32DialogDriver so the CommandRouter can include inline dialog info
/// in the timeout error response — saves agents a round-trip on every
/// modal-opening button click that hits the STA timeout.
/// </summary>
public interface ITimeoutDiagnostics
{
    /// <summary>
    /// Returns a JSON-serializable object describing any dialogs currently
    /// visible in the Excel process, or null if none. Called from the
    /// CommandRouter's timeout handler on the caller thread (NOT the STA
    /// worker — the STA is still stuck). Must be safe to call from any
    /// thread (Win32 EnumWindows is thread-agnostic).
    /// </summary>
    object? GetDialogSnapshot();
}
