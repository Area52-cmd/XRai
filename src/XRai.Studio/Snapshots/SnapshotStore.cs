using System.Collections.Concurrent;
using System.Text.Json.Nodes;

namespace XRai.Studio.Snapshots;

/// <summary>
/// In-memory circular ring buffer of <see cref="Snapshot"/> records. Provides:
///   - Capture(cause): auto-snapshot based on the latest-known state fragments
///     (caller supplies the fragments via the Current* setters below)
///   - Get(id): retrieve a specific snapshot for scrubbing
///   - All: enumerate metadata for the listing endpoint
///   - Last(n): recent snapshots for the dashboard timeline
///
/// Thread-safe: capture can be called from any thread (bus consumer), lookup
/// from the web request thread. Uses a lock around the ring.
///
/// Memory budget: ~65 KB per snapshot × 500 = ~32 MB steady-state worst case.
/// Older snapshots are evicted on overflow; the caller can subscribe to the
/// <see cref="Evicted"/> event if disk persistence is desired.
/// </summary>
public sealed class SnapshotStore
{
    private readonly int _capacity;
    private readonly Snapshot[] _ring;
    private readonly object _lock = new();
    private int _head;
    private int _count;
    private long _nextId = 1;

    // ── Current-state fragments — the bus updates these in real time,
    //    Capture() reads them atomically to build a full Snapshot.
    private byte[]? _latestJpeg;
    private JsonNode? _latestPaneTree;
    private readonly ConcurrentDictionary<string, object?> _modelProperties = new();
    private readonly LinkedList<string> _logTail = new();
    private readonly LinkedList<JsonObject> _commandTail = new();
    private const int LogTailMax = 200;
    private const int CommandTailMax = 50;

    /// <summary>Raised when a snapshot is evicted from the ring. Subscribers
    /// can spill to disk (recording mode) or just ignore it.</summary>
    public event Action<Snapshot>? Evicted;

    public int Capacity => _capacity;
    public int Count { get { lock (_lock) return _count; } }
    public long NextId { get { lock (_lock) return _nextId; } }

    public SnapshotStore(int capacity = 500)
    {
        if (capacity < 1) throw new ArgumentOutOfRangeException(nameof(capacity));
        _capacity = capacity;
        _ring = new Snapshot[capacity];
    }

    // ── Fragment setters, called by the bus or DaemonServer ─────────────

    public void UpdateFrame(byte[]? jpeg) => _latestJpeg = jpeg;
    public void UpdatePaneTree(JsonNode? pane) => _latestPaneTree = pane?.DeepClone();

    public void UpdateModelProperty(string name, object? value)
    {
        if (string.IsNullOrEmpty(name)) return;
        _modelProperties[name] = value;
    }

    public void ReplaceModel(JsonNode? fullModel)
    {
        _modelProperties.Clear();
        if (fullModel is JsonObject obj)
        {
            foreach (var kvp in obj)
            {
                _modelProperties[kvp.Key] = kvp.Value?.ToString();
            }
        }
    }

    public void AppendLog(string line)
    {
        if (string.IsNullOrEmpty(line)) return;
        lock (_logTail)
        {
            _logTail.AddLast(line);
            while (_logTail.Count > LogTailMax) _logTail.RemoveFirst();
        }
    }

    public void AppendCommand(JsonObject cmdEvent)
    {
        lock (_commandTail)
        {
            _commandTail.AddLast((JsonObject)cmdEvent.DeepClone());
            while (_commandTail.Count > CommandTailMax) _commandTail.RemoveFirst();
        }
    }

    // ── Capture + retrieval ─────────────────────────────────────────────

    /// <summary>
    /// Snapshot the current fragments into a new Snapshot record and insert
    /// it into the ring. Returns the captured snapshot so the caller can
    /// publish a corresponding snapshot.captured event.
    /// </summary>
    public Snapshot Capture(string cause)
    {
        Snapshot snap;
        Snapshot? evicted = null;

        var modelJson = new JsonObject();
        foreach (var kvp in _modelProperties)
        {
            try { modelJson[kvp.Key] = kvp.Value?.ToString(); }
            catch { modelJson[kvp.Key] = null; }
        }

        string[] logsSnapshot;
        lock (_logTail) { logsSnapshot = _logTail.ToArray(); }

        JsonArray cmdTail;
        lock (_commandTail)
        {
            cmdTail = new JsonArray();
            foreach (var c in _commandTail) cmdTail.Add(c.DeepClone());
        }

        lock (_lock)
        {
            snap = new Snapshot
            {
                Id = _nextId++,
                Ts = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(),
                Cause = cause,
                Jpeg = _latestJpeg,
                PaneTree = _latestPaneTree?.DeepClone(),
                Model = modelJson,
                LogsTail = logsSnapshot,
                CommandTail = cmdTail,
            };

            if (_count == _capacity)
            {
                evicted = _ring[_head];
            }
            _ring[_head] = snap;
            _head = (_head + 1) % _capacity;
            if (_count < _capacity) _count++;
        }

        if (evicted != null)
        {
            try { Evicted?.Invoke(evicted); } catch { }
        }

        return snap;
    }

    public Snapshot? Get(long id)
    {
        lock (_lock)
        {
            for (int i = 0; i < _count; i++)
            {
                var idx = ((_head - 1 - i) % _capacity + _capacity) % _capacity;
                var s = _ring[idx];
                if (s?.Id == id) return s;
            }
            return null;
        }
    }

    /// <summary>Metadata for every snapshot, newest-first.</summary>
    public List<JsonObject> AllMetadata()
    {
        var list = new List<JsonObject>(Count);
        lock (_lock)
        {
            for (int i = 0; i < _count; i++)
            {
                var idx = ((_head - 1 - i) % _capacity + _capacity) % _capacity;
                var s = _ring[idx];
                if (s != null) list.Add(s.ToMetadata());
            }
        }
        return list;
    }

    /// <summary>Last N snapshots, newest first. Used for the timeline.</summary>
    public List<Snapshot> Last(int n)
    {
        var list = new List<Snapshot>(Math.Min(n, Count));
        lock (_lock)
        {
            int take = Math.Min(n, _count);
            for (int i = 0; i < take; i++)
            {
                var idx = ((_head - 1 - i) % _capacity + _capacity) % _capacity;
                var s = _ring[idx];
                if (s != null) list.Add(s);
            }
        }
        return list;
    }
}
