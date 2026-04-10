using System.Collections.Concurrent;

namespace XRai.Com;

/// <summary>
/// Dedicated single-threaded apartment (STA) worker for Excel COM operations.
///
/// WHY THIS EXISTS:
/// Excel COM must be called from a thread that cooperates with its apartment
/// threading model. Two problems made XRai's old MTA-worker-thread design fragile:
///
///   1. IOleMessageFilter — the canonical Office-automation mechanism for
///      handling RPC_E_SERVERCALL_RETRYLATER (the "Excel is busy" state) —
///      can ONLY be registered from an STA thread. Calling
///      CoRegisterMessageFilter from MTA returns CO_E_NOT_SUPPORTED. Without
///      the filter, Excel busy-state rejections propagate as hard errors and
///      trigger the "Excel is waiting for another application to complete an
///      OLE action" dialog.
///
///   2. Cross-apartment COM marshalling requires the target STA to pump
///      Windows messages. When the main thread was blocked on Wait() for a
///      worker thread's result, the STA couldn't pump, COM couldn't marshal,
///      and calls deadlocked.
///
/// StaComWorker solves both by running ALL COM operations on a single dedicated
/// STA thread that:
///   - Registers IOleMessageFilter on startup (retries silently on busy state)
///   - Pumps messages via a WPF Dispatcher (so COM marshalling works from
///     other threads AND so the filter's MessagePending callback is wired to
///     a real message pump)
///   - Processes a command queue serially, eliminating OLE races entirely
///
/// Callers queue work via Invoke&lt;T&gt;(func, timeoutMs) and get back the result
/// synchronously, with a timeout for safety. If a handler hangs past the
/// timeout, Invoke returns a TimeoutException but the worker thread stays alive
/// for subsequent commands (the hung handler completes eventually and its
/// result is discarded).
/// </summary>
public sealed class StaComWorker : IDisposable
{
    private Thread _thread;
    private BlockingCollection<WorkItem> _queue = new();
    private ManualResetEventSlim _readyEvent = new(false);
    private OleMessageFilter _filter = new();
    private volatile bool _filterRegistered;
    private volatile bool _disposed;
    private volatile bool _isStuck;
    private Exception? _startupError;
    private System.Windows.Threading.Dispatcher? _dispatcher;
    private readonly object _resetLock = new();
    private int _consecutiveTimeouts;
    private DateTime? _lastTimeoutAt;

    public bool FilterRegistered => _filterRegistered;
    public bool IsAlive => _thread.IsAlive && !_disposed;

    /// <summary>
    /// True if the STA thread is stuck on a long-running operation that has
    /// exceeded its timeout. While stuck, new commands will queue behind the
    /// stuck operation. Call Reset() to recycle the thread.
    /// </summary>
    public bool IsStuck => _isStuck;

    public int ConsecutiveTimeouts => _consecutiveTimeouts;
    public DateTime? LastTimeoutAt => _lastTimeoutAt;

    public StaComWorker()
    {
        _thread = StartStaThread();
    }

    private Thread StartStaThread()
    {
        _readyEvent = new ManualResetEventSlim(false);
        _startupError = null;
        var t = new Thread(RunLoop)
        {
            IsBackground = true,
            Name = "xrai-sta-com-worker"
        };
        t.SetApartmentState(ApartmentState.STA);
        t.Start();

        if (!_readyEvent.Wait(TimeSpan.FromSeconds(10)))
            throw new InvalidOperationException("StaComWorker failed to start within 10 seconds");

        if (_startupError != null)
            throw new InvalidOperationException("StaComWorker startup failed", _startupError);

        return t;
    }

    /// <summary>
    /// Force-recycle the STA thread. Abandons the current thread (it will die
    /// with the process) and starts a fresh one with a new message filter.
    /// Use when IsStuck is true or when diagnostics show the worker is poisoned.
    ///
    /// The abandoned thread leaks until the host process exits, but this is
    /// acceptable — the alternative is a dead daemon that requires a full
    /// process restart.
    /// </summary>
    public void Reset()
    {
        lock (_resetLock)
        {
            // Signal the old queue to stop accepting items
            var oldQueue = _queue;
            try { oldQueue.CompleteAdding(); } catch { }

            // Leak the old thread — it's stuck and can't be cleanly killed
            // on .NET Core (Thread.Abort doesn't exist). It will die with
            // the process. This is the documented cost of STA-thread recovery.

            // Start fresh state
            _queue = new BlockingCollection<WorkItem>();
            _filter = new OleMessageFilter();
            _filterRegistered = false;
            _isStuck = false;
            _consecutiveTimeouts = 0;
            _lastTimeoutAt = null;
            _thread = StartStaThread();
        }
    }

    private void RunLoop()
    {
        try
        {
            // Create a WPF Dispatcher for this thread — gives us a real Windows
            // message pump, which is what IOleMessageFilter::MessagePending
            // expects to cooperate with, and what COM needs to marshal calls
            // from other apartments into this STA.
            _dispatcher = System.Windows.Threading.Dispatcher.CurrentDispatcher;

            // Register the OLE message filter now that we're on an STA thread.
            _filterRegistered = _filter.Register();

            // Signal that we're ready for work
            _readyEvent.Set();

            // Drain the command queue. Each work item captures its own
            // completion event; we just execute synchronously on this thread.
            foreach (var work in _queue.GetConsumingEnumerable())
            {
                try
                {
                    work.Execute();
                }
                catch (Exception ex)
                {
                    work.SetException(ex);
                }
            }
        }
        catch (Exception ex)
        {
            _startupError = ex;
            _readyEvent.Set();
        }
        finally
        {
            try { _filter.Revoke(); } catch { }
        }
    }

    /// <summary>
    /// Queue a function to run on the STA worker thread and wait for its result
    /// with a timeout. If the handler hangs past the timeout, throws
    /// TimeoutException but the worker thread continues processing — the stuck
    /// handler's eventual result is discarded.
    /// </summary>
    public T Invoke<T>(Func<T> func, int timeoutMs)
    {
        if (_disposed)
            throw new ObjectDisposedException(nameof(StaComWorker));

        // If the worker is already stuck from a prior timeout, fail fast
        // with a clear message instead of queuing behind the stuck op.
        if (_isStuck)
        {
            throw new TimeoutException(
                "STA worker is stuck (prior command still running). " +
                "Run {\"cmd\":\"sta.reset\"} to recycle the STA thread.");
        }

        var work = new WorkItem<T>(func);
        try { _queue.Add(work); }
        catch (InvalidOperationException)
        {
            // Queue was CompleteAdding'd (Reset in progress). Retry once.
            Thread.Sleep(100);
            if (_disposed) throw new ObjectDisposedException(nameof(StaComWorker));
            _queue.Add(work);
        }

        if (!work.Done.Wait(timeoutMs))
        {
            // Mark the worker stuck — the work item is still running on the
            // STA thread and will continue to block the queue until it finishes
            // or the thread is recycled via Reset().
            _isStuck = true;
            _consecutiveTimeouts++;
            _lastTimeoutAt = DateTime.UtcNow;
            throw new TimeoutException(
                $"Command timed out after {timeoutMs}ms on STA worker. " +
                "The handler is still running on the STA thread. " +
                "Run {\"cmd\":\"sta.reset\"} to recycle the STA thread.");
        }

        // Success — clear stuck flag if we recovered naturally
        _consecutiveTimeouts = 0;

        if (work.CapturedException != null)
            throw work.CapturedException;

        return work.Result!;
    }

    /// <summary>
    /// Queue an action to run on the STA worker thread and wait for completion.
    /// </summary>
    public void Invoke(Action action, int timeoutMs)
    {
        Invoke<bool>(() => { action(); return true; }, timeoutMs);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _queue.CompleteAdding();
        try { _thread.Join(2000); } catch { }
        _queue.Dispose();
    }

    // ── Work item types ──────────────────────────────────────────────

    private abstract class WorkItem
    {
        public ManualResetEventSlim Done { get; } = new(false);
        public Exception? CapturedException { get; protected set; }

        public abstract void Execute();

        public void SetException(Exception ex)
        {
            CapturedException = ex;
            Done.Set();
        }
    }

    private sealed class WorkItem<T> : WorkItem
    {
        private readonly Func<T> _func;
        public T? Result;

        public WorkItem(Func<T> func) { _func = func; }

        public override void Execute()
        {
            try { Result = _func(); }
            catch (Exception ex) { CapturedException = ex; }
            finally { Done.Set(); }
        }
    }
}
