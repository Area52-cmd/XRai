using System.Runtime.InteropServices;

namespace XRai.Com;

/// <summary>
/// Standard Office-automation IOleMessageFilter implementation.
///
/// Every documented Microsoft Office automation sample registers one of these.
/// Without it, when Excel is busy (processing a UI callback, protected view
/// prompt, external link refresh, etc.) and our COM call arrives, Excel returns
/// RPC_E_SERVERCALL_RETRYLATER to the COM runtime. COM's DEFAULT behavior is:
///
///   1. Surface the "Excel is waiting for another application to complete an
///      OLE action" modal dialog to the user
///   2. Block indefinitely until the user clicks Retry or Switch To
///   3. Propagate RPC_E_CALL_REJECTED to the caller on cancel
///
/// With a message filter installed, COM instead calls RetryRejectedCall on our
/// filter, which tells it "retry silently every 100ms for up to 10 seconds".
/// The caller never sees the dialog and never gets the rejection error —
/// the call just takes slightly longer to complete.
///
/// CRITICAL: CoRegisterMessageFilter works ONLY from a single-threaded apartment
/// (STA) thread. Calling it from an MTA thread returns CO_E_NOT_SUPPORTED. This
/// is why XRai Round 10 introduces StaComWorker — a dedicated STA thread that
/// owns the COM session and registers this filter.
///
/// Reference: https://learn.microsoft.com/en-us/previous-versions/ms228772(v=vs.140)
/// (MSDN's canonical "How to: Fix 'Application is Busy' and 'Call Was Rejected
/// by Callee' errors" walkthrough.)
/// </summary>
[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int dwCallType, nint hTaskCaller, int dwTickCount, nint lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(nint hTaskCallee, int dwTickCount, int dwRejectType);

    [PreserveSig]
    int MessagePending(nint hTaskCallee, int dwTickCount, int dwPendingType);
}

public class OleMessageFilter : IDisposable
{
    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    // Reject codes (from objidl.h)
    private const int SERVERCALL_ISHANDLED = 0;
    private const int SERVERCALL_REJECTED = 1;
    private const int SERVERCALL_RETRYLATER = 2;

    // Retry behavior from IOleMessageFilter::RetryRejectedCall:
    //   < 0      → cancel the call, caller gets RPC_E_CALL_REJECTED
    //   0..99    → retry immediately
    //   100+     → wait that many milliseconds, then retry
    private const int RETRY_CANCEL = -1;
    private const int RETRY_INTERVAL_MS = 100;

    // Give up after 30 seconds of retrying. Most Excel busy states resolve in
    // <1s (UI callback), but protected view prompts and external link refresh
    // can take 10-20s legitimately.
    private const int MAX_RETRY_MS = 30_000;

    // Pending type from IOleMessageFilter::MessagePending (from objidl.h)
    private const int PENDINGMSG_CANCELCALL = 0;
    private const int PENDINGMSG_WAITNOPROCESS = 1;
    private const int PENDINGMSG_WAITDEFPROCESS = 2;

    private sealed class Impl : IOleMessageFilter
    {
        public int HandleInComingCall(int dwCallType, nint hTaskCaller, int dwTickCount, nint lpInterfaceInfo)
        {
            // We're a client, not a server — we don't serve incoming calls.
            // Return SERVERCALL_ISHANDLED unconditionally (treat anything we
            // do receive as acceptable; our app isn't expected to process
            // reentrant COM calls from Excel).
            return SERVERCALL_ISHANDLED;
        }

        public int RetryRejectedCall(nint hTaskCallee, int dwTickCount, int dwRejectType)
        {
            // Only retry on SERVERCALL_RETRYLATER. SERVERCALL_REJECTED means
            // the call is permanently rejected (Excel closed, security denied);
            // retrying wouldn't help.
            if (dwRejectType != SERVERCALL_RETRYLATER)
                return RETRY_CANCEL;

            // dwTickCount is the total time (ms) we've been retrying this call.
            if (dwTickCount >= MAX_RETRY_MS)
                return RETRY_CANCEL;

            return RETRY_INTERVAL_MS;
        }

        public int MessagePending(nint hTaskCallee, int dwTickCount, int dwPendingType)
        {
            // A Windows message arrived while we were waiting for a COM call
            // to return. Tell COM to keep waiting and let DefWindowProc handle
            // the message (vs cancelling the call or swallowing the message).
            return PENDINGMSG_WAITDEFPROCESS;
        }
    }

    private Impl? _impl;
    private IOleMessageFilter? _oldFilter;
    private bool _registered;

    /// <summary>
    /// Register this filter on the current thread's apartment. MUST be called
    /// from an STA thread — returns false on MTA (CoRegisterMessageFilter is
    /// not supported on MTA).
    /// </summary>
    public bool Register()
    {
        if (_registered) return true;

        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            return false;

        _impl = new Impl();
        int hr = CoRegisterMessageFilter(_impl, out _oldFilter);
        if (hr != 0)
        {
            _impl = null;
            return false;
        }
        _registered = true;
        return true;
    }

    /// <summary>
    /// Restore the previous message filter and release our own.
    /// </summary>
    public void Revoke()
    {
        if (!_registered) return;
        try { CoRegisterMessageFilter(_oldFilter, out _); } catch { }
        _impl = null;
        _oldFilter = null;
        _registered = false;
    }

    public void Dispose() => Revoke();
}
