namespace XRai.Core;

public class Repl
{
    private readonly CommandRouter _router;
    private readonly EventStream _events;
    private readonly TextReader _input;

    public Repl(CommandRouter router, EventStream events, TextReader? input = null)
    {
        _router = router;
        _events = events;
        _input = input ?? Console.In;
    }

    public void Run()
    {
        // Stream JSON documents from stdin via the shared bracket-counting parser.
        // Supports both single-line (one JSON per line) and multi-line pretty-printed
        // JSON (tracked across newlines via depth counting). Same parser is used by
        // DaemonClient so behavior is identical with or without the daemon.
        foreach (var doc in JsonStreamReader.ReadDocuments(_input))
        {
            var response = _router.Dispatch(doc);
            _events.Write(response);
        }
    }
}
