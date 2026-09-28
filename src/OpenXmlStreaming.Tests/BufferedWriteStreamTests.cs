public class BufferedWriteStreamTests
{
    [Test]
    public async Task WriteAsync_SpillsViaTargetWriteAsync_NotSyncWrite()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        // 16-byte buffer with a 48-byte async write forces 3 spills during WriteAsync.
        // Every spill must go through target.WriteAsync — that's what distinguishes
        // the override from the base-Stream sync fallback.
        await using var buffered = new BufferedWriteStream(tracker, bufferSize: 16, leaveOpen: true);

        await buffered.WriteAsync(new byte[48]);

        using (Assert.Multiple())
        {
            await Assert.That(tracker.AsyncWriteCalls).IsGreaterThanOrEqualTo(2).Because("Spills inside WriteAsync should use target.WriteAsync, not sync Write");
            await Assert.That(tracker.SyncWriteCalls).IsZero().Because("WriteAsync must not fall back to target.Write");
        }
    }

    [Test]
    public async Task WriteAsync_AccumulatesUntilDisposeAsync_WhenFitsInBuffer()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using (var buffered = new BufferedWriteStream(tracker, bufferSize: 1024, leaveOpen: true))
        {
            await buffered.WriteAsync(new byte[100]);
            await Assert.That(tracker.TotalBytesWritten).IsZero().Because("Small async writes that fit should accumulate, not reach the target");
        }

        await Assert.That(tracker.TotalBytesWritten).IsEqualTo(100).Because("DisposeAsync should flush the accumulated bytes");
        await Assert.That(tracker.AsyncWriteCalls).IsGreaterThanOrEqualTo(1);
    }

    [Test]
    public async Task Write_SyncSpillsViaTargetWrite()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        using var buffered = new BufferedWriteStream(tracker, bufferSize: 16, leaveOpen: true);
        buffered.Write(new byte[48], 0, 48);
        buffered.Flush();

        using (Assert.Multiple())
        {
            await Assert.That(tracker.SyncWriteCalls).IsGreaterThan(0).Because("Sync Write should spill via target.Write");
            await Assert.That(tracker.AsyncWriteCalls).IsZero();
        }
    }

    [Test]
    public async Task Dispose_LeaveOpen_DoesNotDisposeTarget()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        using (var buffered = new BufferedWriteStream(tracker, bufferSize: 16, leaveOpen: true))
        {
            buffered.Write([1, 2, 3], 0, 3);
        }

        await Assert.That(() => stream.WriteByte(0)).ThrowsNothing();
    }

    [Test]
    public async Task DisposeAsync_LeaveOpen_DoesNotDisposeTarget()
    {
        using var stream = new MemoryStream();
        var tracker = new SyncAsyncTrackingStream(stream);

        await using (var buffered = new BufferedWriteStream(tracker, bufferSize: 16, leaveOpen: true))
        {
            await buffered.WriteAsync(new byte[] { 1, 2, 3 });
        }

        await Assert.That(() => stream.WriteByte(0)).ThrowsNothing();
    }
}
