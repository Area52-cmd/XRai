namespace XRai.Core;

/// <summary>
/// Contract for a background watchdog that polls for and auto-dismisses
/// nuisance Excel dialogs (NUIDialog, OLE-wait, server-busy, update-links,
/// protected view, file-recovery, etc.) while a long-running COM call is
/// in flight.
///
/// Implemented by Win32DialogDriver in XRai.UI. Consumed by WorkbookOps
/// (and other COM-heavy handlers that trigger dialogs) to bracket COM calls
/// with a watchdog so the STA thread never deadlocks on a modal.
/// </summary>
public interface IDialogWatchdog
{
    /// <summary>
    /// Start the background dismiss loop. Idempotent — calling when already
    /// running is a no-op. intervalMs is clamped to a sane minimum.
    /// </summary>
    void EnableWatchdog(int intervalMs = 250);

    /// <summary>
    /// Stop the background dismiss loop. Idempotent.
    /// </summary>
    void DisableWatchdog();

    /// <summary>
    /// One-shot: scan all Excel windows once and dismiss any that match
    /// the nuisance patterns. Safe to call from any thread. Returns the
    /// number of dialogs dismissed.
    /// </summary>
    int DismissOnce();
}
