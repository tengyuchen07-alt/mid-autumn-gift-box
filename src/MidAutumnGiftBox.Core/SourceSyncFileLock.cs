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

    public void Dispose()
    {
        Interlocked.Exchange(ref _stream, null)?.Dispose();
    }
}
