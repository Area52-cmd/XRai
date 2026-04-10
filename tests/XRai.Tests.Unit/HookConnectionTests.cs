using Xunit;
using XRai.HooksClient;

namespace XRai.Tests.Unit;

public class HookConnectionTests
{
    [Fact]
    public void IsConnected_DefaultFalse()
    {
        using var conn = new HookConnection();
        Assert.False(conn.IsConnected);
    }

    [Fact]
    public void SendCommand_WhenNotConnected_AndAutoReconnectDisabled_Throws()
    {
        using var conn = new HookConnection { AutoReconnect = false };
        Assert.Throws<InvalidOperationException>(() => conn.SendCommand("ping"));
    }

    [Fact]
    public void AutoReconnect_DefaultsToTrue()
    {
        using var conn = new HookConnection();
        Assert.True(conn.AutoReconnect);
    }

    [Fact]
    public void Disconnect_WhenNotConnected_DoesNotThrow()
    {
        using var conn = new HookConnection();
        conn.Disconnect(); // Should not throw
    }
}
