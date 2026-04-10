using System.Diagnostics;
using System.Text.Json.Nodes;
using XRai.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace XRai.Com;

public class CalcOps
{
    private readonly ExcelSession _session;

    public CalcOps(ExcelSession session)
    {
        _session = session;
    }

    public void Register(CommandRouter router)
    {
        router.Register("calc", HandleCalc);
        router.Register("calc.mode", HandleCalcMode);
        router.Register("wait.calc", HandleWaitCalc);
        router.Register("wait.cell", HandleWaitCell);
        router.Register("time.calc", HandleTimeCalc);
    }

    private string HandleCalc(JsonObject args)
    {
        _session.App.Calculate();
        return Response.Ok(new { calculated = true });
    }

    private string HandleCalcMode(JsonObject args)
    {
        var mode = args["mode"]?.GetValue<string>();
        if (mode != null)
        {
            _session.App.Calculation = mode.ToLowerInvariant() switch
            {
                "auto" or "automatic" => Excel.XlCalculation.xlCalculationAutomatic,
                "manual" => Excel.XlCalculation.xlCalculationManual,
                "semiautomatic" => Excel.XlCalculation.xlCalculationSemiautomatic,
                _ => throw new ArgumentException($"Unknown calc mode: {mode}")
            };
            return Response.Ok(new { mode });
        }

        // Read current mode
        var current = _session.App.Calculation switch
        {
            Excel.XlCalculation.xlCalculationAutomatic => "automatic",
            Excel.XlCalculation.xlCalculationManual => "manual",
            Excel.XlCalculation.xlCalculationSemiautomatic => "semiautomatic",
            _ => "unknown"
        };
        return Response.Ok(new { mode = current });
    }

    private string HandleWaitCalc(JsonObject args)
    {
        int timeoutMs = args["timeout"]?.GetValue<int>() ?? 10000;
        var sw = Stopwatch.StartNew();

        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            try
            {
                var status = _session.App.CalculationState;
                if (status == Excel.XlCalculationState.xlDone)
                    return Response.Ok(new { done = true, elapsed_ms = sw.ElapsedMilliseconds });
            }
            catch
            {
                // Excel may be busy
            }
            Thread.Sleep(50);
        }

        return Response.Error($"Calculation did not finish within {timeoutMs}ms");
    }

    private string HandleWaitCell(JsonObject args)
    {
        var refStr = args["ref"]?.GetValue<string>()
            ?? throw new ArgumentException("wait.cell requires 'ref'");
        var condition = args["condition"]?.GetValue<string>() ?? "not_empty";
        var expected = args["value"]?.GetValue<string>();
        int timeoutMs = args["timeout"]?.GetValue<int>() ?? 10000;

        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            using var guard = new ComGuard();
            var sheet = guard.Track(_session.GetActiveSheet());
            var range = guard.Track(sheet.Range[refStr]);
            var val = range.Value2;

            bool met = condition switch
            {
                "not_empty" => val != null && val.ToString() != "",
                "equals" => val?.ToString() == expected,
                "not_error" => !(val is int && range.Text?.ToString()?.StartsWith("#") == true),
                _ => throw new ArgumentException($"Unknown condition: {condition}")
            };

            if (met)
                return Response.Ok(new { @ref = refStr, condition, met = true, value = val?.ToString(), elapsed_ms = sw.ElapsedMilliseconds });

            Thread.Sleep(100);
        }

        return Response.Error($"Cell {refStr} did not meet condition '{condition}' within {timeoutMs}ms");
    }

    private string HandleTimeCalc(JsonObject args)
    {
        var sw = Stopwatch.StartNew();
        _session.App.Calculate();

        // Wait for calc to finish
        while (_session.App.CalculationState != Excel.XlCalculationState.xlDone)
            Thread.Sleep(10);

        sw.Stop();
        return Response.Ok(new { elapsed_ms = sw.ElapsedMilliseconds, elapsed_s = sw.Elapsed.TotalSeconds });
    }
}
