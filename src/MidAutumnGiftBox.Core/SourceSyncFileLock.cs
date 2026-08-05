namespace MidAutumnGiftBox.Core;

public sealed class SourceSyncFileLock : IDisposable
{
    private FileStream? _stream;

    private SourceSyncFileLock(FileStream stream)
    {
        _stream = stream;
    }

    public static SourceSyncFileLock? TryAcquire(string lockDirectory, string sourceCode)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(lockDirectory);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceCode);
        Directory.CreateDirectory(lockDirectory);
        var safeSourceCode = string.Concat(sourceCode.Select(character =>
            Path.GetInvalidFileNameChars().Contains(character) ? '_' : character));
        var path = Path.Combine(Path.GetFullPath(lockDirectory), $"sync-{safeSourceCode}.lock");

        try
        {
            var stream = new FileStream(
                path,
                FileMode.OpenOrCreate,
                FileAccess.ReadWrite,
                FileShare.None,
                1,
                FileOptions.None);
            return new SourceSyncFileLock(stream);
        }
        catch (IOException)
        {
            return null;
        }
    }

    public static async Task<SourceSyncFileLock> AcquireAsync(
        string lockDirectory,
        string sourceCode,
        TimeSpan timeout,
        CancellationToken cancellationToken = default)
    {
        var deadline = DateTimeOffset.UtcNow + timeout;
        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            var acquired = TryAcquire(lockDirectory, sourceCode);
            if (acquired is not null)
            {
                return acquired;
            }

            await Task.Delay(50, cancellationToken);
        }
        while (DateTimeOffset.UtcNow < deadline);

        throw new IOException("本機資料檔正由另一個程序更新，請稍後再試。");
    }

    public void Dispose()
    {
        Interlocked.Exchange(ref _stream, null)?.Dispose();
    }
}
