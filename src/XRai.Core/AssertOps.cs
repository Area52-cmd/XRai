using System.Text.Json;
using System.Text.Json.Nodes;

namespace XRai.Core;

/// <summary>
/// Assertion commands for automated testing. Each assertion reads a value
/// (cell, pane control, or model property) and compares it against an expected
/// value. Results are automatically recorded in the active test session
/// (if test.start was called).
///
/// Commands:
///   assert.cell  — read cell value, compare against expected
///   assert.pane  — read pane control value, compare against expected
///   assert.model — read model property, compare against expected
///   test.assert  — generic assert (delegates to assert.cell/pane/model based on args)
/// </summary>
public class AssertOps
{
    private readonly CommandRouter _router;

    public AssertOps(CommandRouter router)
    {
        _router = router;
    }

    public void Register(CommandRouter router)
    {
        router.Register("assert.cell", HandleAssertCell);
        router.Register("assert.pane", HandleAssertPane);
        router.Register("assert.model", HandleAssertModel);
        router.Register("test.assert", HandleTestAssert);
    }

    // ── assert.cell ─────────────────────────────────────────────────

    /// <summary>
    /// Read a cell and assert its value matches the expected value.
    /// Args: ref (required), value (required), tolerance (optional for numeric)
    /// </summary>
    private string HandleAssertCell(JsonObject args)
    {
        var cellRef = args["ref"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(cellRef))
            return Response.Error("assert.cell requires 'ref'");

        var expectedNode = args["value"];
        if (expectedNode == null)
            return Response.Error("assert.cell requires 'value'");

        var tolerance = args["tolerance"]?.GetValue<double>();
        var expectedStr = expectedNode.ToJsonString().Trim('"');

        // Read the cell value by dispatching a read command
        var readResult = _router.Dispatch(new JsonObject
        {
            ["cmd"] = "read",
            ["ref"] = cellRef,
        }.ToJsonString());

        // Parse the read result
        try
        {
            var doc = JsonDocument.Parse(readResult);
            if (!doc.RootElement.TryGetProperty("ok", out var ok) || !ok.GetBoolean())
            {
                var error = doc.RootElement.TryGetProperty("error", out var errProp) ? errProp.GetString() : "Unknown error";
                RecordResult($"assert.cell {cellRef}", false, $"Read failed: {error}");
                return Response.Error($"Failed to read cell {cellRef}: {error}");
            }

            // Extract the actual value
            string actualStr;
            if (doc.RootElement.TryGetProperty("value", out var valueProp))
            {
                actualStr = valueProp.ValueKind == JsonValueKind.Null ? "" : valueProp.ToString();
            }
            else
            {
                actualStr = "";
            }

            // Compare
            bool passed = CompareValues(actualStr, expectedStr, tolerance);
            var stepName = $"assert.cell {cellRef} == {expectedStr}";
            var message = passed ? null : $"Expected '{expectedStr}', got '{actualStr}'";

            RecordResult(stepName, passed, message);

            return Response.Ok(new
            {
                passed,
                cell = cellRef,
                expected = expectedStr,
                actual = actualStr,
                message,
            });
        }
        catch (Exception ex)
        {
            RecordResult($"assert.cell {cellRef}", false, $"Parse error: {ex.Message}");
            return Response.Error($"Failed to parse read result: {ex.Message}");
        }
    }

    // ── assert.pane ─────────────────────────────────────────────────

    /// <summary>
    /// Read a pane control value and assert it matches the expected value.
    /// Args: control (required), value (required)
    /// </summary>
    private string HandleAssertPane(JsonObject args)
    {
        var control = args["control"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(control))
            return Response.Error("assert.pane requires 'control'");

        var expectedNode = args["value"];
        if (expectedNode == null)
            return Response.Error("assert.pane requires 'value'");

        var expectedStr = expectedNode.ToJsonString().Trim('"');

        // Read the pane control value
        var readResult = _router.Dispatch(new JsonObject
        {
            ["cmd"] = "pane.read",
            ["control"] = control,
        }.ToJsonString());

        try
        {
            var doc = JsonDocument.Parse(readResult);
            if (!doc.RootElement.TryGetProperty("ok", out var ok) || !ok.GetBoolean())
            {
                var error = doc.RootElement.TryGetProperty("error", out var errProp) ? errProp.GetString() : "Unknown error";
                RecordResult($"assert.pane {control}", false, $"Read failed: {error}");
                return Response.Error($"Failed to read pane control '{control}': {error}");
            }

            string actualStr;
            if (doc.RootElement.TryGetProperty("value", out var valueProp))
            {
                actualStr = valueProp.ValueKind == JsonValueKind.Null ? "" : valueProp.ToString();
            }
            else
            {
                actualStr = "";
            }

            bool passed = CompareValues(actualStr, expectedStr, null);
            var stepName = $"assert.pane {control} == {expectedStr}";
            var message = passed ? null : $"Expected '{expectedStr}', got '{actualStr}'";

            RecordResult(stepName, passed, message);

            return Response.Ok(new
            {
                passed,
                control,
                expected = expectedStr,
                actual = actualStr,
                message,
            });
        }
        catch (Exception ex)
        {
            RecordResult($"assert.pane {control}", false, $"Parse error: {ex.Message}");
            return Response.Error($"Failed to parse pane.read result: {ex.Message}");
        }
    }

    // ── assert.model ────────────────────────────────────────────────

    /// <summary>
    /// Read a ViewModel property and assert it matches the expected value.
    /// Args: property (required), value (required)
    /// </summary>
    private string HandleAssertModel(JsonObject args)
    {
        var property = args["property"]?.GetValue<string>();
        if (string.IsNullOrWhiteSpace(property))
            return Response.Error("assert.model requires 'property'");

        var expectedNode = args["value"];
        if (expectedNode == null)
            return Response.Error("assert.model requires 'value'");

        var expectedStr = expectedNode.ToJsonString().Trim('"');

        // Read the model property — dispatch a "model" command and pick out the property
        var readResult = _router.Dispatch(new JsonObject
        {
            ["cmd"] = "model",
        }.ToJsonString());

        try
        {
            var doc = JsonDocument.Parse(readResult);
            if (!doc.RootElement.TryGetProperty("ok", out var ok) || !ok.GetBoolean())
            {
                var error = doc.RootElement.TryGetProperty("error", out var errProp) ? errProp.GetString() : "Unknown error";
                RecordResult($"assert.model {property}", false, $"Model read failed: {error}");
                return Response.Error($"Failed to read model: {error}");
            }

            // Look for the property in the model response (case-insensitive via snake_case)
            string? actualStr = null;
            if (doc.RootElement.TryGetProperty("properties", out var props) &&
                props.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in props.EnumerateObject())
                {
                    if (string.Equals(prop.Name, property, StringComparison.OrdinalIgnoreCase))
                    {
                        actualStr = prop.Value.ValueKind == JsonValueKind.Null ? "" : prop.Value.ToString();
                        break;
                    }
                }
            }

            // Also check top-level properties
            if (actualStr == null)
            {
                foreach (var prop in doc.RootElement.EnumerateObject())
                {
                    if (string.Equals(prop.Name, property, StringComparison.OrdinalIgnoreCase))
                    {
                        actualStr = prop.Value.ValueKind == JsonValueKind.Null ? "" : prop.Value.ToString();
                        break;
                    }
                }
            }

            if (actualStr == null)
            {
                RecordResult($"assert.model {property}", false, $"Property '{property}' not found in model response");
                return Response.Ok(new
                {
                    passed = false,
                    property,
                    expected = expectedStr,
                    actual = (string?)null,
                    message = $"Property '{property}' not found in model response",
                });
            }

            bool passed = CompareValues(actualStr, expectedStr, null);
            var stepName = $"assert.model {property} == {expectedStr}";
            var message = passed ? null : $"Expected '{expectedStr}', got '{actualStr}'";

            RecordResult(stepName, passed, message);

            return Response.Ok(new
            {
                passed,
                property,
                expected = expectedStr,
                actual = actualStr,
                message,
            });
        }
        catch (Exception ex)
        {
            RecordResult($"assert.model {property}", false, $"Parse error: {ex.Message}");
            return Response.Error($"Failed to parse model result: {ex.Message}");
        }
    }

    // ── test.assert ─────────────────────────────────────────────────

    /// <summary>
    /// Generic assert that delegates to assert.cell, assert.pane, or assert.model
    /// based on which arguments are provided.
    /// </summary>
    private string HandleTestAssert(JsonObject args)
    {
        if (args["ref"] != null)
            return HandleAssertCell(args);
        if (args["control"] != null)
            return HandleAssertPane(args);
        if (args["property"] != null)
            return HandleAssertModel(args);

        return Response.Error("test.assert requires 'ref' (cell), 'control' (pane), or 'property' (model)");
    }

    // ── Helpers ──────────────────────────────────────────────────────

    private static bool CompareValues(string actual, string expected, double? tolerance)
    {
        // Try numeric comparison if tolerance is specified
        if (tolerance.HasValue &&
            double.TryParse(actual, out var actualNum) &&
            double.TryParse(expected, out var expectedNum))
        {
            return Math.Abs(actualNum - expectedNum) <= tolerance.Value;
        }

        // String comparison (case-insensitive)
        return string.Equals(actual.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase);
    }

    private static void RecordResult(string stepName, bool passed, string? message)
    {
        TestReporter.RecordAssertResult(stepName, passed, message);
    }
}
