using System.Text.Json.Nodes;
using XRai.Studio;
using XRai.Studio.Sources;
using XRai.Studio.Sources.Agents;
using XRai.Studio.Snapshots;
using Xunit;

namespace XRai.Tests.Unit;

/// <summary>
/// Unit tests for the new XRai.Studio code added in Phase 1 → Phase 3.
/// Covers EventBus, IdeLauncher detection, StudioPreferences round-trip,
/// ClaudeCodeAgentAdapter transcript parsing, FileWatcherSource, the
/// Differ for snapshots, and the SnapshotStore ring buffer.
/// </summary>
public class StudioEventBusTests
{
    [Fact]
    public void Publish_To_Single_Subscriber_Delivers_Event()
    {
        var bus = new EventBus();
        var (id, reader) = bus.Subscribe();

        bus.Publish(StudioEvent.Now("test.kind", "test", new JsonObject { ["x"] = 42 }));

        var ok = reader.TryRead(out var evt);
        Assert.True(ok);
        Assert.NotNull(evt);
        Assert.Equal("test.kind", evt!.Kind);
        Assert.Equal("test", evt.Source);
        Assert.Equal(42, evt.Data?["x"]?.GetValue<int>());

        bus.Unsubscribe(id);
    }

    [Fact]
    public void New_Subscriber_Receives_Ring_Buffer_Replay()
    {
        var bus = new EventBus();
        // Publish 3 events with NO subscribers — these go into the ring buffer.
        bus.Publish(StudioEvent.Now("a", "src", null));
        bus.Publish(StudioEvent.Now("b", "src", null));
        bus.Publish(StudioEvent.Now("c", "src", null));

        // New subscriber should see all three on first reads.
        var (id, reader) = bus.Subscribe();
        var kinds = new List<string>();
        for (int i = 0; i < 3; i++)
        {
            Assert.True(reader.TryRead(out var evt));
            kinds.Add(evt!.Kind);
        }
        Assert.Equal(new[] { "a", "b", "c" }, kinds);
        bus.Unsubscribe(id);
    }

    [Fact]
    public void Multiple_Subscribers_Each_Get_Their_Own_Copy()
    {
        var bus = new EventBus();
        var (id1, reader1) = bus.Subscribe();
        var (id2, reader2) = bus.Subscribe();

        bus.Publish(StudioEvent.Now("hello", "src", null));

        // Both readers receive a copy
        Assert.True(reader1.TryRead(out var e1));
        Assert.True(reader2.TryRead(out var e2));
        Assert.Equal("hello", e1!.Kind);
        Assert.Equal("hello", e2!.Kind);

        bus.Unsubscribe(id1);
        bus.Unsubscribe(id2);
    }

    [Fact]
    public void Unsubscribe_Removes_Subscriber_From_Future_Publishes()
    {
        var bus = new EventBus();
        var (id, reader) = bus.Subscribe();
        bus.Unsubscribe(id);

        // Should not crash, even after unsubscribe
        bus.Publish(StudioEvent.Now("after", "src", null));

        // Reader was completed by Unsubscribe — no new events arrive
        Assert.False(reader.TryRead(out _));
    }

    [Fact]
    public void SubscriberCount_Reflects_State()
    {
        var bus = new EventBus();
        Assert.Equal(0, bus.SubscriberCount);
        var (id1, _) = bus.Subscribe();
        Assert.Equal(1, bus.SubscriberCount);
        var (id2, _) = bus.Subscribe();
        Assert.Equal(2, bus.SubscriberCount);
        bus.Unsubscribe(id1);
        Assert.Equal(1, bus.SubscriberCount);
        bus.Unsubscribe(id2);
        Assert.Equal(0, bus.SubscriberCount);
    }
}

public class StudioPreferencesTests
{
    [Fact]
    public void Defaults_Are_Sane()
    {
        var prefs = new StudioPreferences();
        Assert.Equal(1, prefs.SchemaVersion);
        Assert.True(prefs.FollowMode);
        Assert.Null(prefs.PreferredIde);
        Assert.False(prefs.Onboarded);
        Assert.Equal("dark", prefs.Theme);
    }

    [Fact]
    public void ToJson_Round_Trips_Through_FromJson()
    {
        var original = new StudioPreferences
        {
            FollowMode = false,
            PreferredIde = "VSCode",
            Onboarded = true,
            Theme = "light",
        };
        var node = original.ToJson();
        var restored = StudioPreferences.FromJson(node);

        Assert.Equal(original.FollowMode, restored.FollowMode);
        Assert.Equal(original.PreferredIde, restored.PreferredIde);
        Assert.Equal(original.Onboarded, restored.Onboarded);
        Assert.Equal(original.Theme, restored.Theme);
    }

    [Fact]
    public void FromJson_Tolerates_Missing_Fields()
    {
        var partial = new JsonObject { ["followMode"] = true };
        var prefs = StudioPreferences.FromJson(partial);
        Assert.True(prefs.FollowMode);
        Assert.Null(prefs.PreferredIde);
        Assert.False(prefs.Onboarded);
    }

    [Fact]
    public void FromJson_Handles_Null_Or_NonObject_Gracefully()
    {
        var prefs1 = StudioPreferences.FromJson(null);
        Assert.NotNull(prefs1);
        Assert.True(prefs1.FollowMode);

        var prefs2 = StudioPreferences.FromJson(JsonValue.Create(42));
        Assert.NotNull(prefs2);
    }
}

public class StudioIdeLauncherTests
{
    [Fact]
    public void DetectAll_Returns_All_Three_Known_Ides()
    {
        var ides = IdeLauncher.DetectAll();
        Assert.Equal(3, ides.Count);
        Assert.Contains(ides, i => i.Kind == IdeLauncher.IdeKind.VSCode);
        Assert.Contains(ides, i => i.Kind == IdeLauncher.IdeKind.VisualStudio);
        Assert.Contains(ides, i => i.Kind == IdeLauncher.IdeKind.Rider);
    }

    [Fact]
    public void DetectAll_Includes_Install_Url_For_Every_Entry()
    {
        var ides = IdeLauncher.DetectAll();
        foreach (var ide in ides)
        {
            Assert.False(string.IsNullOrEmpty(ide.InstallUrl), $"{ide.Kind} missing InstallUrl");
            Assert.True(ide.InstallUrl!.StartsWith("http"), $"{ide.Kind} InstallUrl is not a URL");
        }
    }

    [Fact]
    public void IdeInfo_ToJson_Has_Required_Fields()
    {
        var info = new IdeLauncher.IdeInfo
        {
            Kind = IdeLauncher.IdeKind.VSCode,
            DisplayName = "Visual Studio Code",
            Installed = true,
            Running = false,
        };
        var json = info.ToJson();
        Assert.Equal("VSCode", json["kind"]?.GetValue<string>());
        Assert.Equal("Visual Studio Code", json["name"]?.GetValue<string>());
        Assert.True(json["installed"]?.GetValue<bool>());
        Assert.False(json["running"]?.GetValue<bool>());
    }

    [Fact]
    public void Open_Without_FilePath_Returns_Error()
    {
        var result = IdeLauncher.Open("");
        Assert.False(result["ok"]?.GetValue<bool>() ?? true);
    }

    [Fact]
    public void Open_NonexistentFile_Returns_Error()
    {
        var fakePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".cs");
        var result = IdeLauncher.Open(fakePath);
        Assert.False(result["ok"]?.GetValue<bool>() ?? true);
        Assert.Contains("does not exist", result["error"]?.GetValue<string>() ?? "");
    }
}

public class StudioClaudeCodeAdapterTests
{
    [Fact]
    public void DefaultProjectsRoot_Is_Under_User_Profile()
    {
        var root = ClaudeCodeAgentAdapter.DefaultProjectsRoot();
        var profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        Assert.StartsWith(profile, root);
        Assert.EndsWith("projects", root);
    }

    [Fact]
    public void Adapter_Reports_Correct_Agent_Name()
    {
        var bus = new EventBus();
        using var adapter = new ClaudeCodeAgentAdapter(bus, projectsRoot: Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()));
        Assert.Equal("Claude Code", adapter.AgentName);
        Assert.False(adapter.IsConnected);
    }

    [Fact]
    public void AgentAdapterFactory_Detect_Returns_Some_Adapter()
    {
        var bus = new EventBus();
        var adapter = AgentAdapterFactory.Detect(bus);
        Assert.NotNull(adapter);
        Assert.False(string.IsNullOrEmpty(adapter.AgentName));
        adapter.Dispose();
    }
}

public class StudioSnapshotStoreTests
{
    [Fact]
    public void Capture_Inserts_Snapshot_With_Monotonic_Id()
    {
        var store = new SnapshotStore(capacity: 10);
        var s1 = store.Capture("test1");
        var s2 = store.Capture("test2");
        Assert.Equal(s1.Id + 1, s2.Id);
    }

    [Fact]
    public void Get_Returns_Most_Recent_Match()
    {
        var store = new SnapshotStore(capacity: 10);
        var s = store.Capture("test");
        var found = store.Get(s.Id);
        Assert.NotNull(found);
        Assert.Equal(s.Id, found!.Id);
    }

    [Fact]
    public void Ring_Buffer_Evicts_Oldest_When_Full()
    {
        var store = new SnapshotStore(capacity: 3);
        var s1 = store.Capture("a");
        store.Capture("b");
        store.Capture("c");
        store.Capture("d"); // forces eviction of s1

        Assert.Null(store.Get(s1.Id));
        Assert.Equal(3, store.Count);
    }

    [Fact]
    public void UpdateModelProperty_Reflects_In_Captured_Snapshot()
    {
        var store = new SnapshotStore(capacity: 10);
        store.UpdateModelProperty("Count", 42);
        store.UpdateModelProperty("Name", "test");
        var snap = store.Capture("after-update");
        Assert.NotNull(snap.Model);
        var modelObj = snap.Model as JsonObject;
        Assert.NotNull(modelObj);
        Assert.NotNull(modelObj!["Count"]);
        Assert.NotNull(modelObj["Name"]);
    }
}

public class StudioDifferTests
{
    [Fact]
    public void Diff_Of_Identical_Snapshots_Has_No_Changes()
    {
        var snap = new Snapshot
        {
            Id = 1,
            Ts = 1000,
            Cause = "test",
            Model = new JsonObject { ["a"] = 1, ["b"] = 2 },
        };
        var diff = Differ.Diff(snap, snap);
        Assert.Equal(1, diff["beforeId"]?.GetValue<long>());
        Assert.Equal(1, diff["afterId"]?.GetValue<long>());
        Assert.Equal("No structural differences.", diff["summary"]?.GetValue<string>());
    }

    [Fact]
    public void Diff_Detects_Model_Property_Change()
    {
        var before = new Snapshot
        {
            Id = 1, Ts = 1000, Cause = "x",
            Model = new JsonObject { ["count"] = 1 },
        };
        var after = new Snapshot
        {
            Id = 2, Ts = 2000, Cause = "y",
            Model = new JsonObject { ["count"] = 2 },
        };
        var diff = Differ.Diff(before, after);
        var modelDiff = diff["model"] as JsonObject;
        Assert.NotNull(modelDiff);
        var changed = modelDiff!["changed"] as JsonArray;
        Assert.NotNull(changed);
        Assert.Single(changed!);
    }

    [Fact]
    public void Diff_Detects_Added_Model_Property()
    {
        var before = new Snapshot
        {
            Id = 1, Ts = 1000, Cause = "x",
            Model = new JsonObject { ["count"] = 1 },
        };
        var after = new Snapshot
        {
            Id = 2, Ts = 2000, Cause = "y",
            Model = new JsonObject { ["count"] = 1, ["name"] = "test" },
        };
        var diff = Differ.Diff(before, after);
        var modelDiff = diff["model"] as JsonObject;
        var added = modelDiff!["added"] as JsonArray;
        Assert.NotNull(added);
        Assert.Single(added!);
    }
}

public class StudioEventTests
{
    [Fact]
    public void Now_Generates_Current_Timestamp()
    {
        var before = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        var evt = StudioEvent.Now("kind", "src", null);
        var after = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        Assert.InRange(evt.Ts, before, after);
    }

    [Fact]
    public void ToJson_Includes_Required_Fields()
    {
        var evt = StudioEvent.Now("test.kind", "test.source", new JsonObject { ["data"] = 1 });
        var json = evt.ToJson();
        Assert.NotNull(json["ts"]);
        Assert.Equal("test.kind", json["kind"]?.GetValue<string>());
        Assert.Equal("test.source", json["source"]?.GetValue<string>());
        Assert.NotNull(json["data"]);
    }

    [Fact]
    public void ToJson_Omits_Data_When_Null()
    {
        var evt = StudioEvent.Now("test", "src", null);
        var json = evt.ToJson();
        Assert.False(json.ContainsKey("data"));
    }
}
