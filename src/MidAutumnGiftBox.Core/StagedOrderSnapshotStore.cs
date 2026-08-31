using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text.Json;

namespace MidAutumnGiftBox.Core;

public static class StagedSnapshotStates
{
    public const string Staged = "staged";
    public const string ExportPrepared = "export_prepared";
    public const string ExcelCommitted = "excel_committed";
}

public sealed record StagedOrderSnapshot(
    string TransactionId,
    DateTimeOffset StagedAt,
    string State,
    string? TargetWorkbookPath,
    IReadOnlyList<OrderLineSnapshot> Lines,
    ManualOverrideState? PendingManualState = null,
    string? PreparedWorkbookPath = null,
    string? PreparedWorkbookSha256 = null);

public sealed record SnapshotSelection(
    bool IsStaged,
    string? TransactionId,
    DateTimeOffset? StagedAt,
    string? State,
    string? TargetWorkbookPath,
    IReadOnlyList<OrderLineSnapshot> Lines);

public sealed class StagedOrderSnapshotStore
{
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> FileLocks =
        new(StringComparer.OrdinalIgnoreCase);
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path;
    private readonly SemaphoreSlim _fileLock;

    public StagedOrderSnapshotStore(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("暫存訂單快照檔案路徑不可空白。", nameof(path));
        }

        _path = Path.GetFullPath(path);
        _fileLock = FileLocks.GetOrAdd(_path, _ => new SemaphoreSlim(1, 1));
    }

    public async Task<StagedOrderSnapshot> StageAsync(
        IReadOnlyList<OrderLineSnapshot> lines,
        DateTimeOffset stagedAt,
        CancellationToken cancellationToken = default)
    {
        var ordered = OrderSnapshotStore.NormalizeReplacement(lines);
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            using var dataLock = await AcquireDataFileLockAsync(cancellationToken);
            var existing = await LoadUnlockedAsync(cancellationToken);
            if (existing is not null && existing.State != StagedSnapshotStates.Staged)
            {
                throw new InvalidOperationException(
                    "前一次正式 Excel 交易尚未提交，請先完成復原後再重新載入。");
            }

            var staged = new StagedOrderSnapshot(
                Guid.NewGuid().ToString("N"),
                stagedAt,
                StagedSnapshotStates.Staged,
                null,
                ordered);
            await SaveUnlockedAsync(staged, cancellationToken);
            return staged;
        }
        finally
        {
            _fileLock.Release();
        }
    }

    public async Task<StagedOrderSnapshot?> GetAsync(CancellationToken cancellationToken = default)
    {
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            return await LoadUnlockedAsync(cancellationToken);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    public async Task<SnapshotSelection> GetPreferredAsync(
        OrderSnapshotStore committedStore,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(committedStore);
        var staged = await GetAsync(cancellationToken);
        if (staged is not null)
        {
            return new SnapshotSelection(
                true,
                staged.TransactionId,
                staged.StagedAt,
                staged.State,
                staged.TargetWorkbookPath,
                staged.Lines);
        }

        return new SnapshotSelection(
            false,
            null,
            null,
            null,
            null,
            await committedStore.GetAllAsync(cancellationToken));
    }

    public async Task<SnapshotSelection> EnsureForFormalExportAsync(
        SnapshotSelection selection,
        DateTimeOffset stagedAt,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(selection);
        if (selection.IsStaged)
        {
            if (!string.Equals(selection.State, StagedSnapshotStates.Staged, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "前一次正式 Excel 交易尚未完成復原，不能開始另一筆正式匯出。");
            }

            return selection;
        }

        var transaction = await StageAsync(selection.Lines, stagedAt, cancellationToken);
        return new SnapshotSelection(
            true,
            transaction.TransactionId,
            transaction.StagedAt,
            transaction.State,
            transaction.TargetWorkbookPath,
            transaction.Lines);
    }

    public async Task MarkExcelCommittedAsync(
        string transactionId,
        string targetWorkbookPath,
        CancellationToken cancellationToken = default)
        => await MarkExcelCommittedAsync(
            transactionId,
            targetWorkbookPath,
            pendingManualState: null,
            cancellationToken);

    public async Task MarkExcelCommittedAsync(
        string transactionId,
        string targetWorkbookPath,
        ManualOverrideState? pendingManualState,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(transactionId);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetWorkbookPath);
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            using var dataLock = await AcquireDataFileLockAsync(cancellationToken);
            var staged = await LoadRequiredUnlockedAsync(transactionId, cancellationToken);
            await SaveUnlockedAsync(
                staged with
                {
                    State = StagedSnapshotStates.ExcelCommitted,
                    TargetWorkbookPath = Path.GetFullPath(targetWorkbookPath),
                    PendingManualState = pendingManualState ?? staged.PendingManualState,
                    PreparedWorkbookPath = null
                },
                cancellationToken);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    public async Task MarkExportPreparedAsync(
        string transactionId,
        string targetWorkbookPath,
        string preparedWorkbookPath,
        ManualOverrideState pendingManualState,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(transactionId);
        ArgumentException.ThrowIfNullOrWhiteSpace(targetWorkbookPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(preparedWorkbookPath);
        ArgumentNullException.ThrowIfNull(pendingManualState);
        var preparedFullPath = Path.GetFullPath(preparedWorkbookPath);
        if (!File.Exists(preparedFullPath))
        {
            throw new FileNotFoundException("找不到已備妥的正式 Excel 暫存檔。", preparedFullPath);
        }
        var preparedSha256 = ComputeFileSha256(preparedFullPath);

        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            using var dataLock = await AcquireDataFileLockAsync(cancellationToken);
            var staged = await LoadRequiredUnlockedAsync(transactionId, cancellationToken);
            if (staged.State != StagedSnapshotStates.Staged)
            {
                throw new InvalidOperationException("暫存訂單已進入正式匯出交易，不能重複備妥。 ");
            }

            await SaveUnlockedAsync(
                staged with
                {
                    State = StagedSnapshotStates.ExportPrepared,
                    TargetWorkbookPath = Path.GetFullPath(targetWorkbookPath),
                    PendingManualState = pendingManualState,
                    PreparedWorkbookPath = preparedFullPath,
                    PreparedWorkbookSha256 = preparedSha256
                },
                cancellationToken);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    public static bool MatchesPreparedWorkbook(StagedOrderSnapshot snapshot, string path)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return !string.IsNullOrWhiteSpace(snapshot.PreparedWorkbookSha256) &&
               File.Exists(path) &&
               snapshot.PreparedWorkbookSha256.Equals(
                   ComputeFileSha256(path),
                   StringComparison.OrdinalIgnoreCase);
    }

    public async Task CommitAsync(
        string transactionId,
        OrderSnapshotStore committedStore,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(transactionId);
        ArgumentNullException.ThrowIfNull(committedStore);
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            using var dataLock = await AcquireDataFileLockAsync(cancellationToken);
            var staged = await LoadRequiredUnlockedAsync(transactionId, cancellationToken);
            if (staged.State != StagedSnapshotStates.ExcelCommitted)
            {
                throw new InvalidOperationException("暫存訂單尚未完成正式 Excel，不能提交為正式快照。");
            }

            await committedStore.ReplaceAllAsync(staged.Lines, cancellationToken);
            File.Delete(_path);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    private async Task<StagedOrderSnapshot> LoadRequiredUnlockedAsync(
        string transactionId,
        CancellationToken cancellationToken)
    {
        var staged = await LoadUnlockedAsync(cancellationToken)
                     ?? throw new InvalidOperationException("找不到尚未正式匯出的訂單快照。");
        if (!staged.TransactionId.Equals(transactionId, StringComparison.Ordinal))
        {
            throw new InvalidOperationException("暫存訂單已被較新的完整載入取代，請重新開始匯出。");
        }

        return staged;
    }

    private async Task<StagedOrderSnapshot?> LoadUnlockedAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_path))
        {
            return null;
        }

        await using var stream = new FileStream(
            _path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
        return await JsonSerializer.DeserializeAsync<StagedOrderSnapshot>(
            stream, JsonOptions, cancellationToken);
    }

    private async Task SaveUnlockedAsync(
        StagedOrderSnapshot snapshot,
        CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(_path)
                        ?? throw new InvalidOperationException("暫存訂單快照檔案缺少目錄。");
        Directory.CreateDirectory(directory);
        var temporaryPath = _path + $".{Guid.NewGuid():N}.tmp";
        try
        {
            await using (var stream = new FileStream(
                             temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            {
                await JsonSerializer.SerializeAsync(stream, snapshot, JsonOptions, cancellationToken);
            }

            File.Move(temporaryPath, _path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }

    private Task<SourceSyncFileLock> AcquireDataFileLockAsync(CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(_path)
                        ?? throw new InvalidOperationException("暫存訂單快照檔案缺少目錄。");
        return SourceSyncFileLock.AcquireAsync(
            directory, "order-snapshot-staged-data", TimeSpan.FromSeconds(10), cancellationToken);
    }

    private static string ComputeFileSha256(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
