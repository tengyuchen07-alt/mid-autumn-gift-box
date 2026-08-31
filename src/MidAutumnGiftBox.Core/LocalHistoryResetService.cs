using System.Collections.Concurrent;

namespace MidAutumnGiftBox.Core;

public sealed record LocalHistoryResetResult(
    string BackupDirectory,
    IReadOnlyList<string> BackedUpFileNames);

public sealed class LocalHistoryRefreshException : InvalidOperationException
{
    public LocalHistoryRefreshException(
        string message,
        string backupDirectory,
        bool historyRestored,
        Exception innerException)
        : base(message, innerException)
    {
        BackupDirectory = backupDirectory;
        HistoryRestored = historyRestored;
    }

    public string BackupDirectory { get; }
    public bool HistoryRestored { get; }
}

public sealed class LocalHistoryResetService
{
    private static readonly string[] HistoryFileNames =
    [
        "sync-status.json",
        "order-snapshot.json",
        "order-snapshot.staged.json",
        "order-change-ledger.json"
    ];

    private static readonly ConcurrentDictionary<string, SemaphoreSlim> DirectoryLocks =
        new(StringComparer.OrdinalIgnoreCase);

    private readonly string _dataDirectory;
    private readonly string[] _sourceCodes;
    private readonly SemaphoreSlim _directoryLock;

    public LocalHistoryResetService(string dataDirectory, IReadOnlyList<string> sourceCodes)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(dataDirectory);
        ArgumentNullException.ThrowIfNull(sourceCodes);
        if (sourceCodes.Count == 0 || sourceCodes.Any(string.IsNullOrWhiteSpace))
        {
            throw new ArgumentException("歷史重設至少需要一個有效資料來源。", nameof(sourceCodes));
        }

        _dataDirectory = Path.GetFullPath(dataDirectory);
        _sourceCodes = sourceCodes.Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(code => code, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        _directoryLock = DirectoryLocks.GetOrAdd(_dataDirectory, _ => new SemaphoreSlim(1, 1));
    }

    public async Task<LocalHistoryResetResult> BackupAndResetAsync(
        DateTimeOffset resetAt,
        CancellationToken cancellationToken = default)
    {
        await _directoryLock.WaitAsync(cancellationToken);
        try
        {
            using var sourceLocks = await AcquireSourceLocksAsync(cancellationToken);
            Directory.CreateDirectory(_dataDirectory);
            var backupDirectory = Path.Combine(
                _dataDirectory,
                "history-backups",
                $"{resetAt:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}");
            Directory.CreateDirectory(backupDirectory);
            var moved = new List<string>();
            try
            {
                foreach (var fileName in HistoryFileNames)
                {
                    var activePath = Path.Combine(_dataDirectory, fileName);
                    if (!File.Exists(activePath))
                    {
                        continue;
                    }

                    File.Move(activePath, Path.Combine(backupDirectory, fileName));
                    moved.Add(fileName);
                }

                return new LocalHistoryResetResult(backupDirectory, moved);
            }
            catch
            {
                foreach (var fileName in moved.AsEnumerable().Reverse())
                {
                    var backupPath = Path.Combine(backupDirectory, fileName);
                    var activePath = Path.Combine(_dataDirectory, fileName);
                    if (File.Exists(backupPath) && !File.Exists(activePath))
                    {
                        File.Move(backupPath, activePath);
                    }
                }

                throw;
            }
        }
        finally
        {
            _directoryLock.Release();
        }
    }

    public async Task<LocalHistoryResetResult> ResetAndRefreshAsync(
        DateTimeOffset resetAt,
        Func<CancellationToken, Task> refresh,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(refresh);
        using var workflowLock = await AcquireWorkflowLockAsync(cancellationToken);
        var reset = await BackupAndResetAsync(resetAt, cancellationToken);
        try
        {
            await refresh(cancellationToken);
            return reset;
        }
        catch (Exception refreshException)
        {
            try
            {
                await RestoreAsync(reset, CancellationToken.None);
            }
            catch (Exception restoreException)
            {
                throw new LocalHistoryRefreshException(
                    $"全量重抓失敗，且無法自動復原。舊資料備份仍位於：{reset.BackupDirectory}。",
                    reset.BackupDirectory,
                    historyRestored: false,
                    new AggregateException(refreshException, restoreException));
            }

            throw new LocalHistoryRefreshException(
                $"全量重抓失敗，已自動復原原本的本機歷史。原因：{refreshException.Message}",
                reset.BackupDirectory,
                historyRestored: true,
                refreshException);
        }
    }

    public async Task<IDisposable> AcquireWorkflowLockAsync(
        CancellationToken cancellationToken = default) =>
        await SourceSyncFileLock.AcquireAsync(
            Path.Combine(_dataDirectory, "locks"),
            "history-refresh",
            TimeSpan.FromSeconds(10),
            cancellationToken);

    public async Task RestoreAsync(
        LocalHistoryResetResult reset,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(reset);
        var backupDirectory = Path.GetFullPath(reset.BackupDirectory);
        var backupRoot = Path.GetFullPath(Path.Combine(_dataDirectory, "history-backups")) +
                         Path.DirectorySeparatorChar;
        if (!backupDirectory.StartsWith(backupRoot, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("歷史備份位置不屬於目前資料目錄，拒絕復原。");
        }

        await _directoryLock.WaitAsync(cancellationToken);
        try
        {
            using var sourceLocks = await AcquireSourceLocksAsync(cancellationToken);
            var failedRefreshDirectory = Path.Combine(backupDirectory, "failed-refresh");
            var movedRefreshFiles = new List<string>();
            var restoredFiles = new List<string>();
            try
            {
                foreach (var fileName in HistoryFileNames)
                {
                    var activePath = Path.Combine(_dataDirectory, fileName);
                    if (File.Exists(activePath))
                    {
                        Directory.CreateDirectory(failedRefreshDirectory);
                        File.Move(activePath, Path.Combine(failedRefreshDirectory, fileName));
                        movedRefreshFiles.Add(fileName);
                    }
                }

                foreach (var fileName in HistoryFileNames)
                {
                    var backupPath = Path.Combine(backupDirectory, fileName);
                    if (!File.Exists(backupPath))
                    {
                        continue;
                    }

                    File.Move(backupPath, Path.Combine(_dataDirectory, fileName));
                    restoredFiles.Add(fileName);
                }
            }
            catch
            {
                foreach (var fileName in restoredFiles.AsEnumerable().Reverse())
                {
                    var activePath = Path.Combine(_dataDirectory, fileName);
                    var backupPath = Path.Combine(backupDirectory, fileName);
                    if (File.Exists(activePath) && !File.Exists(backupPath))
                    {
                        File.Move(activePath, backupPath);
                    }
                }

                foreach (var fileName in movedRefreshFiles.AsEnumerable().Reverse())
                {
                    var failedRefreshPath = Path.Combine(failedRefreshDirectory, fileName);
                    var activePath = Path.Combine(_dataDirectory, fileName);
                    if (File.Exists(failedRefreshPath) && !File.Exists(activePath))
                    {
                        File.Move(failedRefreshPath, activePath);
                    }
                }

                throw;
            }
        }
        finally
        {
            _directoryLock.Release();
        }
    }

    private async Task<IDisposable> AcquireSourceLocksAsync(CancellationToken cancellationToken)
    {
        var locks = new List<SourceSyncFileLock>();
        try
        {
            var lockDirectory = Path.Combine(_dataDirectory, "locks");
            foreach (var sourceCode in _sourceCodes)
            {
                locks.Add(await SourceSyncFileLock.AcquireAsync(
                    lockDirectory,
                    sourceCode,
                    TimeSpan.FromSeconds(10),
                    cancellationToken));
            }

            return new CompositeDisposable(locks);
        }
        catch
        {
            foreach (var sourceLock in locks)
            {
                sourceLock.Dispose();
            }

            throw;
        }
    }

    private sealed class CompositeDisposable(IReadOnlyList<SourceSyncFileLock> locks) : IDisposable
    {
        public void Dispose()
        {
            foreach (var sourceLock in locks.Reverse())
            {
                sourceLock.Dispose();
            }
        }
    }
}
