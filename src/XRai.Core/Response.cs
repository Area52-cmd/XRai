using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace XRai.Core;

public static class Response
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
    };

    public static string Ok(object? data = null)
    {
        var node = new JsonObject { ["ok"] = true };
        if (data != null)
            MergeData(node, data);
        return node.ToJsonString(SerializerOptions);
    }

    public static string Error(string message, string? code = null)
    {
        var node = new JsonObject
        {
            ["ok"] = false,
            ["error"] = message,
        };
        if (code != null) node["code"] = code;
        var hint = ErrorHints.GetHint(message);
        if (hint != null) node["hint"] = hint;
        var docs = ErrorHints.GetDocsUrl(code);
        if (docs != null) node["docs"] = docs;
        return node.ToJsonString(SerializerOptions);
    }

    public static string ErrorWithData(string message, object? data, string? code = null)
    {
        var node = new JsonObject
        {
            ["ok"] = false,
            ["error"] = message,
        };
        if (code != null) node["code"] = code;
        if (data != null)
            MergeData(node, data);
        var hint = ErrorHints.GetHint(message);
        if (hint != null && node["hint"] == null) node["hint"] = hint;
        var docs = ErrorHints.GetDocsUrl(code);
        if (docs != null && node["docs"] == null) node["docs"] = docs;
        return node.ToJsonString(SerializerOptions);
    }

    /// <summary>
    /// Build an error response from an exception. Captures type, message,
    /// inner exception, and top stack frame. Includes auto-hint lookup.
    /// Use this in every command handler's catch block.
    /// </summary>
    public static string ErrorFromException(Exception ex, string? contextCommand = null, string? code = null)
    {
        var type = ex.GetType().Name;
        var message = ex.Message;
        var inner = ex.InnerException?.Message;
        var topFrame = GetTopStackFrame(ex);
        var hresult = ex is System.Runtime.InteropServices.COMException comEx
            ? $"0x{comEx.HResult:X8}"
            : null;

        var errorText = contextCommand != null
            ? $"{contextCommand} failed: {type}: {message}"
            : $"{type}: {message}";

        // Infer a code from the exception type if one wasn't supplied.
        var resolvedCode = code ?? InferCodeFromException(ex);

        var node = new JsonObject
        {
            ["ok"] = false,
            ["error"] = errorText,
            ["code"] = resolvedCode,
            ["exception_type"] = type,
            ["exception_message"] = message,
        };
        if (inner != null) node["inner_exception"] = inner;
        if (topFrame != null) node["stack_frame"] = topFrame;
        if (hresult != null) node["hresult"] = hresult;

        var hint = ErrorHints.GetHint(errorText, ex);
        if (hint != null) node["hint"] = hint;

        var docs = ErrorHints.GetDocsUrl(resolvedCode);
        if (docs != null) node["docs"] = docs;

        return node.ToJsonString(SerializerOptions);
    }

    /// <summary>
    /// Map a .NET exception to a stable XRai error code. Used by
    /// ErrorFromException when the caller doesn't pass an explicit code.
    /// </summary>
    private static string InferCodeFromException(Exception ex)
    {
        switch (ex)
        {
            case System.Runtime.InteropServices.COMException comEx:
                var hex = unchecked((uint)comEx.HResult);
                if (hex == 0x800706BA) return ErrorCodes.ComServerUnavailable;
                if (hex == 0x8001010A /* RPC_E_SERVERCALL_RETRYLATER */) return ErrorCodes.ComServerBusy;
                if (hex == 0x80010105 /* RPC_E_SERVERFAULT */) return ErrorCodes.ComServerBusy;
                if (hex == 0x80010108 /* RPC_E_DISCONNECTED */) return ErrorCodes.ComServerUnavailable;
                return ErrorCodes.InternalError;

            case TimeoutException:
                return ex.Message.Contains("STA worker", StringComparison.OrdinalIgnoreCase)
                    ? ErrorCodes.StaTimeout
                    : ErrorCodes.Timeout;

            case ArgumentNullException:
            case ArgumentException:
                return ErrorCodes.InvalidArgument;

            case InvalidOperationException:
                return ErrorCodes.InternalError;

            case NotImplementedException:
                return ErrorCodes.NotImplemented;

            default:
                return ErrorCodes.InternalError;
        }
    }

    private static string? GetTopStackFrame(Exception ex)
    {
        try
        {
            var stack = ex.StackTrace;
            if (string.IsNullOrEmpty(stack)) return null;
            var firstLine = stack.Split('\n')[0].Trim();
            // Strip leading "at " prefix
            if (firstLine.StartsWith("at ")) firstLine = firstLine.Substring(3);
            return firstLine.Length > 200 ? firstLine.Substring(0, 200) : firstLine;
        }
        catch { return null; }
    }

    public static string Event(string eventType, object? data = null)
    {
        var node = new JsonObject { ["event"] = eventType };
        if (data != null)
            MergeData(node, data);
        return node.ToJsonString(SerializerOptions);
    }

    private static void MergeData(JsonObject target, object data)
    {
        var json = JsonSerializer.Serialize(data, SerializerOptions);
        var parsed = JsonNode.Parse(json);
        if (parsed is JsonObject obj)
        {
            foreach (var kvp in obj)
                target[kvp.Key] = kvp.Value?.DeepClone();
        }
    }
}
