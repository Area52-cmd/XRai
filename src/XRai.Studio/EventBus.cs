using System.Collections.Concurrent;
using System.Threading.Channels;

namespace XRai.Studio;

/// <summary>
/// In-process pub/sub event bus for Studio. Single writer-many-readers model:
///   - Sources (PipeEventSource, CaptureLoop, FileWatcherSource, etc.) call
///     Publish(StudioEvent) from any thread.
///   - Subscribers (WebSocket clients) call Subscribe() and get a bounded
///     Channel&lt;StudioEvent&gt; to drain.
///
/// Each subscriber has its own channel with capacity 1024; on overflow the
/// oldest events are dropped (BoundedChannelFullMode.DropOldest) so a slow
/// client never blocks the bus or memory-bloats the server.
///
/// The bus also maintains a small ring buffer of the last 256 events so a
/// newly-connecting client gets the most recent history as a replay before
/// the live stream starts.
/// </summary>
public sealed class EventBus
{
    private readonly ConcurrentDictionary<Guid, Channel<StudioEvent>> _subscribers = new();
    private readonly object _ringLock = new();
    private readonly StudioEvent[] _ring = new StudioEvent[256];
    private int _ringHead; // index of next write slot
    private int _ringCount; // number of valid entries (0..256)

    public int SubscriberCount => _subscribers.Count;

    /// <summary>
    /// Publish an event to all subscribers. Non-blocking — each channel is
    /// DropOldest so this never waits even if a subscriber is slow.
    /// </summary>
    public void Publish(StudioEvent evt)
    {
        // Record in ring buffer first so late subscribers see it.
        lock (_ringLock)
        {
            _ring[_ringHead] = evt;
            _ringHead = (_ringHead + 1) % _ring.Length;
            if (_ringCount < _ring.Length) _ringCount++;
        }

        foreach (var kvp in _subscribers)
        {
            // TryWrite on a BoundedChannel with DropOldest never blocks.
            try { kvp.Value.Writer.TryWrite(evt); }
            catch { /* channel closed — subscriber will be removed on next iteration */ }
        }
    }

    /// <summary>
    /// Subscribe to the live event stream. Returns the subscriber ID (for
    /// Unsubscribe) and the channel reader to drain. The reader yields the
    /// ring-buffer replay FIRST, then the live stream, in arrival order.
    /// </summary>
    public (Guid Id, ChannelReader<StudioEvent> Reader) Subscribe()
    {
        var channel = Channel.CreateBounded<StudioEvent>(new BoundedChannelOptions(1024)
        {
            SingleWriter = false,
            SingleReader = true,
            FullMode = BoundedChannelFullMode.DropOldest,
            AllowSynchronousContinuations = false,
        });

        // Pre-load the ring buffer into the channel so the new subscriber
        // sees recent history before the live stream kicks in.
        lock (_ringLock)
        {
            if (_ringCount > 0)
            {
                int start = _ringCount < _ring.Length ? 0 : _ringHead;
                for (int i = 0; i < _ringCount; i++)
                {
                    var idx = (start + i) % _ring.Length;
                    channel.Writer.TryWrite(_ring[idx]);
                }
            }
        }

        var id = Guid.NewGuid();
        _subscribers[id] = channel;
        return (id, channel.Reader);
    }

    /// <summary>
    /// Remove a subscriber and complete its channel so any pending reader
    /// observes completion and exits cleanly.
    /// </summary>
    public void Unsubscribe(Guid id)
    {
        if (_subscribers.TryRemove(id, out var ch))
        {
            try { ch.Writer.TryComplete(); } catch { }
        }
    }
}
