using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class ExcelSession : IDisposable
{
    private Excel.Application? _app;
    private bool _disposed;

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
        nint pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    public Excel.Application App => _app ?? throw new InvalidOperationException("Not attached to Excel");
    public bool IsAttached => _app != null;
    public string? ExcelVersion => _app?.Version;

    /// <summary>
    /// Attach to the running Excel instance. Polls COM readiness for up to
    /// <paramref name="comReadyTimeoutMs"/> ms, retrying on transient errors:
    ///   - 0x800706BE (RPC call failed — Excel busy/starting)
    ///   - 0x800706BA (RPC server unavailable — Excel not ready)
    ///   - 0x800401E3 (MK_E_UNAVAILABLE — ROT not populated yet)
    /// This prevents the "attach failed instantly on cold start" failure mode.
    /// </summary>
    public void Attach(int? pid = null, int comReadyTimeoutMs = 10000)
    {
        if (_app != null)
            throw new InvalidOperationException("Already attached. Call Detach() first.");

        if (pid.HasValue)
        {
            try { Process.GetProcessById(pid.Value); }
            catch { throw new InvalidOperationException($"No process with PID {pid.Value}"); }
        }

        var clsid = Type.GetTypeFromProgID("Excel.Application", true)!.GUID;

        var sw = Stopwatch.StartNew();
        Exception? lastError = null;
        while (sw.ElapsedMilliseconds < comReadyTimeoutMs)
        {
            try
            {
                GetActiveObject(clsid, 0, out object obj);
                _app = (Excel.Application)obj;
                return;
            }
            catch (COMException ex) when (
                (uint)ex.HResult == 0x800706BE || // RPC call failed
                (uint)ex.HResult == 0x800706BA || // RPC server unavailable
                (uint)ex.HResult == 0x800401E3)   // MK_E_UNAVAILABLE
            {
                lastError = ex;
                Thread.Sleep(250);
            }
        }

        throw new InvalidOperationException(
            $"COM readiness timeout after {comReadyTimeoutMs}ms. " +
            $"Excel may not be running, may be on the start screen, or may be stuck. " +
            $"Last error: {lastError?.Message ?? "unknown"}");
    }

    public void WaitAndAttach(int timeoutMs = 30000, CancellationToken ct = default)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            ct.ThrowIfCancellationRequested();

            try
            {
                Attach();
                return;
            }
            catch (COMException)
            {
                Thread.Sleep(500);
            }
        }

        throw new TimeoutException($"Excel did not appear within {timeoutMs}ms");
    }

    public void Detach()
    {
        if (_app != null)
        {
            try { Marshal.ReleaseComObject(_app); } catch { }
            _app = null;
        }
    }

    public Excel.Workbook GetActiveWorkbook()
    {
        return App.ActiveWorkbook
            ?? throw new InvalidOperationException("No active workbook");
    }

    /// <summary>
    /// Idempotent: returns the active workbook, creating a fresh Book1 if Excel
    /// is on the start screen with nothing open. Safe to call multiple times.
    /// Never touches an existing workbook.
    /// </summary>
    public Excel.Workbook EnsureWorkbook()
    {
        var active = App.ActiveWorkbook;
        if (active != null) return active;

        var workbooks = App.Workbooks;
        try
        {
            var wb = workbooks.Add();
            return wb;
        }
        finally
        {
            Marshal.ReleaseComObject(workbooks);
        }
    }

    public Excel.Worksheet GetActiveSheet()
    {
        return (Excel.Worksheet)(App.ActiveSheet
            ?? throw new InvalidOperationException("No active sheet"));
    }

    /// <summary>
    /// Resolve a reference string to an Excel.Range, accepting both:
    ///   - Bare refs: "A1", "A1:D10", "B2"
    ///   - Sheet-qualified refs: "Sheet1!A1", "Sheet1!A1:D10", "'My Sheet'!A1:D10"
    ///   - Defined names: "MyRange"
    /// Returns a Range object that the caller must release via ComGuard.
    /// This is the single source of truth for address normalization —
    /// every Ops class should call this instead of sheet.Range[addr] directly.
    /// </summary>
    public Excel.Range ResolveRange(string refStr)
    {
        if (string.IsNullOrWhiteSpace(refStr))
            throw new ArgumentException("Range reference is empty");

        // Sheet-qualified: "Sheet1!A1:D10" or "'Sheet with spaces'!A1:D10"
        int bangIndex = refStr.IndexOf('!');
        if (bangIndex > 0)
        {
            string sheetPart = refStr.Substring(0, bangIndex);
            string rangePart = refStr.Substring(bangIndex + 1);

            // Strip quotes from sheet name if present
            if (sheetPart.StartsWith("'") && sheetPart.EndsWith("'") && sheetPart.Length >= 2)
                sheetPart = sheetPart.Substring(1, sheetPart.Length - 2);

            // Find the worksheet by name.
            // Leak-audited: 2026-04-10. Previously the matched worksheet
            // ('target') was returned via target.Range[...] without releasing
            // the worksheet proxy itself, leaking one COM RCW per cross-sheet
            // range resolution. We now release the worksheet immediately after
            // grabbing the Range — Excel keeps the Range valid via internal
            // back-references so the early release is safe.
            var workbook = GetActiveWorkbook();
            try
            {
                var sheets = workbook.Worksheets;
                try
                {
                    Excel.Worksheet? target = null;
                    foreach (Excel.Worksheet ws in sheets)
                    {
                        if (string.Equals(ws.Name, sheetPart, StringComparison.OrdinalIgnoreCase))
                        {
                            target = ws;
                            break;
                        }
                        Marshal.ReleaseComObject(ws);
                    }
                    if (target == null)
                        throw new ArgumentException($"Sheet not found: '{sheetPart}'. Available: {string.Join(", ", GetSheetNames())}");

                    try
                    {
                        return target.Range[rangePart];
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(target);
                    }
                }
                finally { Marshal.ReleaseComObject(sheets); }
            }
            finally { Marshal.ReleaseComObject(workbook); }
        }

        // Bare ref: use active sheet
        var activeSheet = GetActiveSheet();
        return activeSheet.Range[refStr];
    }

    private string[] GetSheetNames()
    {
        try
        {
            var workbook = GetActiveWorkbook();
            try
            {
                var sheets = workbook.Worksheets;
                try
                {
                    var names = new List<string>();
                    foreach (Excel.Worksheet ws in sheets)
                    {
                        names.Add(ws.Name);
                        Marshal.ReleaseComObject(ws);
                    }
                    return names.ToArray();
                }
                finally { Marshal.ReleaseComObject(sheets); }
            }
            finally { Marshal.ReleaseComObject(workbook); }
        }
        catch { return Array.Empty<string>(); }
    }

    /// <summary>
    /// Returns (hasWorkbook, workbookName). Does not throw.
    /// Use to query state before attempting operations that require a workbook.
    /// </summary>
    public (bool HasWorkbook, string? Name, int Count) ProbeWorkbookState()
    {
        try
        {
            var workbooks = App.Workbooks;
            try
            {
                int count = workbooks.Count;
                var active = App.ActiveWorkbook;
                return (active != null, active?.Name, count);
            }
            finally
            {
                Marshal.ReleaseComObject(workbooks);
            }
        }
        catch
        {
            return (false, null, 0);
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Detach();
        ComGuard.FinalCleanup();
    }
}
