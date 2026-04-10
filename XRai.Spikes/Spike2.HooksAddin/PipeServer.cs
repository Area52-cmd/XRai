using System.Diagnostics;
using System.IO.Pipes;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelDna.Integration;
namespace Spike2.HooksAddin;

public class PipeServer
{
    private readonly string _pipeName;
    private CancellationTokenSource? _cts;
    private Thread? _thread;

    // Mutable control state for set_control
    private readonly Dictionary<string, ControlInfo> _controls = new()
    {
        ["SpotInput"] = new("SpotInput", "TextBox", "100.00"),
        ["CalcButton"] = new("CalcButton", "Button", null, true),
        ["ResultLabel"] = new("ResultLabel", "Label", "Price: 42.50"),
    };

    public PipeServer(string pipeName)
    {
        _pipeName = pipeName;
    }

    public void Start()
    {
        _cts = new CancellationTokenSource();
        _thread = new Thread(ServerLoop) { IsBackground = true, Name = "XRai-PipeServer" };
        _thread.Start();
    }

    public void Stop()
    {
        _cts?.Cancel();
    }

    private void ServerLoop()
    {
        while (_cts != null && !_cts.IsCancellationRequested)
        {
            NamedPipeServerStream? pipe = null;
            try
            {
                pipe = new NamedPipeServerStream(_pipeName, PipeDirection.InOut, 1,
                    PipeTransmissionMode.Byte, PipeOptions.Asynchronous);

                Debug.WriteLine($"Pipe server waiting for connection on {_pipeName}...");

                // Wait for connection with cancellation support
                var connectTask = pipe.WaitForConnectionAsync(_cts!.Token);
                connectTask.Wait(_cts.Token);

                Debug.WriteLine("Pipe client connected");

                using var reader = new StreamReader(pipe);
                using var writer = new StreamWriter(pipe) { AutoFlush = true };

                while (pipe.IsConnected && !_cts.IsCancellationRequested)
                {
                    string? line = reader.ReadLine();
                    if (line == null) break; // client disconnected

                    Debug.WriteLine($"Pipe received: {line}");
                    HandleCommand(line, writer);
                }

                Debug.WriteLine("Pipe client disconnected");
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Pipe error: {ex.Message}");
            }
            finally
            {
                pipe?.Dispose();
            }
        }

        Debug.WriteLine("Pipe server loop ended");
    }

    private void HandleCommand(string json, StreamWriter writer)
    {
        try
        {
            var node = JsonNode.Parse(json);
            string? cmd = node?["cmd"]?.GetValue<string>();

            switch (cmd)
            {
                case "ping":
                    HandlePing(writer);
                    break;
                case "info":
                    HandleInfo(writer);
                    break;
                case "controls":
                    HandleControls(writer);
                    break;
                case "set_control":
                    HandleSetControl(node!, writer);
                    break;
                case "click":
                    HandleClick(node!, writer);
                    break;
                case "log_test":
                    HandleLogTest(writer);
                    break;
                case "error_test":
                    HandleErrorTest(writer);
                    break;
                default:
                    writer.WriteLine(JsonSerializer.Serialize(new { ok = false, error = $"Unknown command: {cmd}" }));
                    break;
            }
        }
        catch (Exception ex)
        {
            writer.WriteLine(JsonSerializer.Serialize(new { ok = false, error = ex.Message }));
        }
    }

    private void HandlePing(StreamWriter writer)
    {
        int pid = Process.GetCurrentProcess().Id;
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            ok = true,
            message = "pong",
            pid,
            addin = "Spike2.HooksAddin"
        }));
    }

    private void HandleInfo(StreamWriter writer)
    {
        string excelVersion = "unknown";
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            excelVersion = app.Version;
        }
        catch { }

        writer.WriteLine(JsonSerializer.Serialize(new
        {
            ok = true,
            excel_version = excelVersion,
            addin_name = "Spike2.HooksAddin",
            functions_registered = 1,
            hooks_active = true
        }));
    }

    private void HandleControls(StreamWriter writer)
    {
        var controls = _controls.Values.Select(c => new
        {
            name = c.Name,
            type = c.Type,
            value = c.Value,
            enabled = c.Enabled
        }).ToArray();

        writer.WriteLine(JsonSerializer.Serialize(new { ok = true, controls }));
    }

    private void HandleSetControl(JsonNode node, StreamWriter writer)
    {
        string? name = node["name"]?.GetValue<string>();
        string? value = node["value"]?.GetValue<string>();

        if (name != null && _controls.TryGetValue(name, out var ctrl))
        {
            string? oldValue = ctrl.Value;
            ctrl.Value = value;
            writer.WriteLine(JsonSerializer.Serialize(new
            {
                ok = true,
                name,
                old_value = oldValue,
                new_value = value
            }));
        }
        else
        {
            writer.WriteLine(JsonSerializer.Serialize(new { ok = false, error = $"Control not found: {name}" }));
        }
    }

    private void HandleClick(JsonNode node, StreamWriter writer)
    {
        string? name = node["name"]?.GetValue<string>();
        Debug.WriteLine($"Button clicked: {name}");
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            ok = true,
            name,
            clicked = true
        }));
    }

    private void HandleLogTest(StreamWriter writer)
    {
        writer.WriteLine(JsonSerializer.Serialize(new { ok = true }));
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            @event = "log",
            message = "This is a test log from the add-in",
            source = "Debug",
            timestamp = DateTime.UtcNow.ToString("o")
        }));
    }

    private void HandleErrorTest(StreamWriter writer)
    {
        writer.WriteLine(JsonSerializer.Serialize(new { ok = true }));
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            @event = "error",
            exception = "NullReferenceException",
            message = "Object reference not set to an instance of an object",
            stack = "at PricerPane.OnCalcClick() in PricerPane.xaml.cs:line 42",
            timestamp = DateTime.UtcNow.ToString("o")
        }));
    }

    private class ControlInfo
    {
        public string Name { get; }
        public string Type { get; }
        public string? Value { get; set; }
        public bool Enabled { get; }

        public ControlInfo(string name, string type, string? value, bool enabled = true)
        {
            Name = name;
            Type = type;
            Value = value;
            Enabled = enabled;
        }
    }
}
