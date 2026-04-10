using ExcelDna.Integration;
using System.Diagnostics;

namespace Spike2.HooksAddin;

public class AddInEntry : IExcelAddIn
{
    private PipeServer? _pipeServer;

    public void AutoOpen()
    {
        int pid = Process.GetCurrentProcess().Id;
        string pipeName = $"xrai_{pid}";

        _pipeServer = new PipeServer(pipeName);
        _pipeServer.Start();

        Debug.WriteLine($"XRai hooks started on pipe: {pipeName}");
    }

    public void AutoClose()
    {
        _pipeServer?.Stop();
        Debug.WriteLine("XRai hooks stopped");
    }
}

public static class Functions
{
    [ExcelFunction(Description = "Test function")]
    public static string PilotTest(string input)
    {
        return $"Pilot received: {input}";
    }
}
