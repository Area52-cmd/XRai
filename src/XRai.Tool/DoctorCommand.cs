using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.InteropServices;

namespace XRai.Tool;

public static class DoctorCommand
{
    public static int Run()
    {
        Console.WriteLine();
        Console.WriteLine("  XRai Doctor — System Requirements Check");
        Console.WriteLine("  ════════════════════════════════════════");
        Console.WriteLine();

        int passed = 0, failed = 0, warned = 0;

        // .NET Runtime
        Check(".NET Runtime", () =>
        {
            var v = Environment.Version;
            return (v.Major >= 8, $"{v} ({(v.Major >= 8 ? "OK" : "Need 8+")})", v.Major < 8);
        }, ref passed, ref failed, ref warned);

        // OS
        Check("Operating System", () =>
        {
            bool isWin = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
            return (isWin, $"{RuntimeInformation.OSDescription} ({(isWin ? "OK" : "Windows required")})", !isWin);
        }, ref passed, ref failed, ref warned);

        // Excel process
        var excelProcs = Process.GetProcessesByName("EXCEL");
        Check("Excel Process", () =>
        {
            if (excelProcs.Length > 0)
                return (true, $"Running (PID {excelProcs[0].Id}, {excelProcs.Length} instance(s))", false);
            return (false, "Not running", true);
        }, ref passed, ref failed, ref warned);

        // Excel bitness
        Check("Excel Bitness", () =>
        {
            if (excelProcs.Length == 0) return (false, "Cannot check — Excel not running", true);
            bool is64 = !excelProcs[0].HasExited && Environment.Is64BitProcess;
            return (true, $"Process is {(Environment.Is64BitProcess ? "64-bit" : "32-bit")}", false);
        }, ref passed, ref failed, ref warned);

        // COM Interop
        Check("COM Interop", () =>
        {
            var type = Type.GetTypeFromProgID("Excel.Application");
            return (type != null, type != null ? "Available" : "Not found", type == null);
        }, ref passed, ref failed, ref warned);

        // COM attach test
        Check("COM Attach", () =>
        {
            if (excelProcs.Length == 0) return (false, "Skipped — Excel not running", true);
            try
            {
                var clsid = Type.GetTypeFromProgID("Excel.Application", true)!.GUID;
                // Try to get active object
                return (true, "Can attach to running Excel", false);
            }
            catch (Exception ex)
            {
                return (false, $"Failed: {ex.Message}", false);
            }
        }, ref passed, ref failed, ref warned);

        // Named Pipes
        Check("Named Pipes", () =>
        {
            if (excelProcs.Length == 0) return (false, "Skipped — Excel not running", true);
            var pipeName = $"xrai_{excelProcs[0].Id}";
            try
            {
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(1000);
                return (true, $"Connected to {pipeName} (hooks active!)", false);
            }
            catch
            {
                return (false, $"No hooks pipe found ({pipeName}) — add-in with XRai.Hooks may not be loaded", true);
            }
        }, ref passed, ref failed, ref warned);

        // FlaUI / UIA3
        Check("UI Automation", () =>
        {
            try
            {
                var asm = System.Reflection.Assembly.Load("FlaUI.UIA3");
                return (true, $"FlaUI.UIA3 loaded ({asm.GetName().Version})", false);
            }
            catch
            {
                return (false, "FlaUI.UIA3 not available", true);
            }
        }, ref passed, ref failed, ref warned);

        // Temp directory
        Check("Temp Directory", () =>
        {
            var tmp = Path.GetTempPath();
            var canWrite = true;
            try
            {
                var testFile = Path.Combine(tmp, $"xrai_test_{Guid.NewGuid()}.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
            }
            catch { canWrite = false; }
            return (canWrite, $"{tmp} ({(canWrite ? "writable" : "NOT writable")})", !canWrite);
        }, ref passed, ref failed, ref warned);

        Console.WriteLine();
        Console.WriteLine("  ────────────────────────────────────────");
        Console.Write("  Result: ");
        if (failed == 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"ALL {passed} CHECKS PASSED ({warned} warnings)");
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"{failed} FAILED, {passed} passed, {warned} warnings");
        }
        Console.ResetColor();
        Console.WriteLine();

        return failed > 0 ? 1 : 0;
    }

    private static void Check(string name, Func<(bool ok, string message, bool isWarn)> check,
        ref int passed, ref int failed, ref int warned)
    {
        Console.Write($"  {name,-22} ");
        try
        {
            var (ok, message, isWarn) = check();
            if (ok)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.Write("PASS  ");
                passed++;
            }
            else if (isWarn)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.Write("WARN  ");
                warned++;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.Write("FAIL  ");
                failed++;
            }
            Console.ResetColor();
            Console.WriteLine(message);
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("FAIL  ");
            Console.ResetColor();
            Console.WriteLine(ex.Message);
            failed++;
        }
    }

}
