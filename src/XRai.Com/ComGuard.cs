using System.Runtime.InteropServices;

namespace XRai.Com;

public sealed class ComGuard : IDisposable
{
    private readonly Stack<object> _tracked = new();
    private bool _disposed;

    public T Track<T>(T comObject) where T : class
    {
        if (comObject == null)
            throw new ArgumentNullException(nameof(comObject));
        _tracked.Push(comObject);
        return comObject;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        while (_tracked.Count > 0)
        {
            var obj = _tracked.Pop();
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch
            {
                // Swallow — object may already be released
            }
        }
    }

    /// <summary>
    /// Full GC cleanup. Call once when the entire Excel session is closing.
    /// </summary>
    public static void FinalCleanup()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
