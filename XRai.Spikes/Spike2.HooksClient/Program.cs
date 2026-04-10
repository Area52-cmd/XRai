using System.Diagnostics;
using System.IO.Pipes;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace Spike2.HooksClient;

class Program
{
    static int _passed = 0;
    static int _failed = 0;

    static void Main(string[] args)
    {
        try
        {
            // 1. Find Excel process
            var excelProcesses = Process.GetProcessesByName("EXCEL");
            if (excelProcesses.Length == 0)
            {
                Console.WriteLine("ERROR: No Excel process found. Is Excel running with the add-in loaded?");
                return;
            }

            int pid;
            if (args.Length > 0 && int.TryParse(args[0], out int argPid))
                pid = argPid;
            else
                pid = excelProcesses[0].Id;

            Console.WriteLine($"Found Excel process: PID {pid}");

            // 2. Connect to named pipe
            string pipeName = $"xrai_{pid}";
            using var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);

            Console.WriteLine($"Connecting to pipe: {pipeName}...");
            pipe.Connect(5000);
            Console.WriteLine($"Connected to pipe: {pipeName}");

            using var reader = new StreamReader(pipe);
            using var writer = new StreamWriter(pipe) { AutoFlush = true };

            // 3. Test: ping
            {
                var resp = SendCommand(writer, reader, new { cmd = "ping" });
                Assert(resp?["ok"]?.GetValue<bool>() == true, "ping ok");
                Assert(resp?["message"]?.GetValue<string>() == "pong", "ping message");
                Pass("ping");
            }

            // 4. Test: info
            {
                var resp = SendCommand(writer, reader, new { cmd = "info" });
                Assert(resp?["ok"]?.GetValue<bool>() == true, "info ok");
                Assert(resp?["excel_version"] != null, "info has excel_version");
                Assert(resp?["functions_registered"] != null, "info has functions_registered");
                Console.WriteLine($"  Excel version: {resp?["excel_version"]}");
                Pass("info");
            }

            // 5. Test: controls
            {
                var resp = SendCommand(writer, reader, new { cmd = "controls" });
                Assert(resp?["ok"]?.GetValue<bool>() == true, "controls ok");
                var controls = resp?["controls"]?.AsArray();
                Assert(controls?.Count == 3, $"controls count = {controls?.Count}");
                if (controls != null)
                {
                    foreach (var c in controls)
                    {
                        Console.WriteLine($"  Control: {c?["name"]} ({c?["type"]}) = {c?["value"]}");
                    }
                }
                Pass("controls");
            }

            // 6. Test: set_control
            {
                var resp = SendCommand(writer, reader, new { cmd = "set_control", name = "SpotInput", value = "105.00" });
                Assert(resp?["ok"]?.GetValue<bool>() == true, "set_control ok");
                Assert(resp?["old_value"]?.GetValue<string>() == "100.00", "old_value");
                Assert(resp?["new_value"]?.GetValue<string>() == "105.00", "new_value");
                Pass("set_control");
            }

            // 7. Test: controls again (verify state changed)
            {
                var resp = SendCommand(writer, reader, new { cmd = "controls" });
                var controls = resp?["controls"]?.AsArray();
                var spot = controls?.FirstOrDefault(c => c?["name"]?.GetValue<string>() == "SpotInput");
                Assert(spot?["value"]?.GetValue<string>() == "105.00", "SpotInput updated to 105.00");
                Pass("controls (state persisted)");
            }

            // 8. Test: click
            {
                var resp = SendCommand(writer, reader, new { cmd = "click", name = "CalcButton" });
                Assert(resp?["ok"]?.GetValue<bool>() == true, "click ok");
                Assert(resp?["clicked"]?.GetValue<bool>() == true, "clicked true");
                Pass("click");
            }

            // 9. Test: log_test (push event)
            {
                var resp = SendCommand(writer, reader, new { cmd = "log_test" });
                Assert(resp?["ok"]?.GetValue<bool>() == true, "log_test ok");
                // Read the pushed event
                string? eventLine = reader.ReadLine();
                var evt = JsonNode.Parse(eventLine ?? "{}");
                Assert(evt?["event"]?.GetValue<string>() == "log", "event is log");
                Console.WriteLine($"  Log message: {evt?["message"]}");
                Pass("log push event");
            }

            // 10. Test: error_test (push error event)
            {
                var resp = SendCommand(writer, reader, new { cmd = "error_test" });
                Assert(resp?["ok"]?.GetValue<bool>() == true, "error_test ok");
                // Read the pushed error event
                string? eventLine = reader.ReadLine();
                var evt = JsonNode.Parse(eventLine ?? "{}");
                Assert(evt?["event"]?.GetValue<string>() == "error", "event is error");
                Assert(evt?["stack"] != null, "has stack trace");
                Console.WriteLine($"  Error: {evt?["exception"]}: {evt?["message"]}");
                Console.WriteLine($"  Stack: {evt?["stack"]}");
                Pass("error push event");
            }

            // 11. Test: unknown command
            {
                var resp = SendCommand(writer, reader, new { cmd = "nonexistent" });
                Assert(resp?["ok"]?.GetValue<bool>() == false, "unknown cmd ok=false");
                Assert(resp?["error"]?.GetValue<string>()?.Contains("Unknown command") == true, "error msg");
                Pass("unknown command handling");
            }

            // 12. Print summary
            Console.WriteLine();
            Console.WriteLine("═══════════════════════════════════════");
            if (_failed == 0)
            {
                Console.WriteLine($"SPIKE 2 RESULTS: ALL {_passed} TESTS PASSED");
            }
            else
            {
                Console.WriteLine($"SPIKE 2 RESULTS: {_passed} PASSED, {_failed} FAILED");
            }
            Console.WriteLine($"- Pipe connection:     OK");
            Console.WriteLine($"- Ping:                {(_passed >= 1 ? "OK" : "FAIL")}");
            Console.WriteLine($"- Info:                {(_passed >= 2 ? "OK" : "FAIL")}");
            Console.WriteLine($"- Controls read:       {(_passed >= 3 ? "OK" : "FAIL")}");
            Console.WriteLine($"- Control set:         {(_passed >= 4 ? "OK" : "FAIL")}");
            Console.WriteLine($"- State persistence:   {(_passed >= 5 ? "OK" : "FAIL")}");
            Console.WriteLine($"- Button click:        {(_passed >= 6 ? "OK" : "FAIL")}");
            Console.WriteLine($"- Log push event:      {(_passed >= 7 ? "OK" : "FAIL")}");
            Console.WriteLine($"- Error push event:    {(_passed >= 8 ? "OK" : "FAIL")}");
            Console.WriteLine($"- Error handling:      {(_passed >= 9 ? "OK" : "FAIL")}");
            Console.WriteLine("═══════════════════════════════════════");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FATAL ERROR: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    static JsonNode? SendCommand(StreamWriter writer, StreamReader reader, object cmd)
    {
        string json = JsonSerializer.Serialize(cmd);
        Console.WriteLine($"> {json}");
        writer.WriteLine(json);
        string? response = reader.ReadLine();
        Console.WriteLine($"< {response}");
        return JsonNode.Parse(response ?? "{}");
    }

    static void Assert(bool condition, string description)
    {
        if (!condition)
            throw new Exception($"ASSERTION FAILED: {description}");
    }

    static void Pass(string testName)
    {
        _passed++;
        Console.WriteLine($"PASS: {testName}");
        Console.WriteLine();
    }
}
