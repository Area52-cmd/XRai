namespace XRai.Core;

public class EventStream
{
    private readonly TextWriter _output;
    private readonly object _lock = new();

    public EventStream(TextWriter output)
    {
        _output = output;
    }

    public void Emit(string eventType, object? data = null)
    {
        var line = Response.Event(eventType, data);
        lock (_lock)
        {
            _output.WriteLine(line);
            _output.Flush();
        }
    }

    public void Write(string jsonLine)
    {
        lock (_lock)
        {
            _output.WriteLine(jsonLine);
            _output.Flush();
        }
    }
}
