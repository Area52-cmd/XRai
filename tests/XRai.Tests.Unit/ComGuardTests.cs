using Xunit;
using XRai.Com;

namespace XRai.Tests.Unit;

public class ComGuardTests
{
    [Fact]
    public void Track_ReturnsTheSameObject()
    {
        using var guard = new ComGuard();
        var obj = new object();
        var tracked = guard.Track(obj);
        Assert.Same(obj, tracked);
    }

    [Fact]
    public void Track_NullThrows()
    {
        using var guard = new ComGuard();
        Assert.Throws<ArgumentNullException>(() => guard.Track<object>(null!));
    }

    [Fact]
    public void Dispose_DoesNotThrowForNonComObjects()
    {
        // ComGuard should swallow errors when releasing non-COM objects
        var guard = new ComGuard();
        guard.Track(new object());
        guard.Track(new object());
        guard.Dispose(); // Should not throw
    }
}
