using System.Collections.Concurrent;
using System.Text.Json;

namespace MidAutumnGiftBox.Core;

public sealed record SyncRunSummary(
    string SourceCode,
    string SourceName,
    DateOnly QueryFrom,
    DateOnly QueryThrough,
    DateTimeOffset StartedAt,
    DateTimeOffset FinishedAt,
    int PageCount,
    int RecordCount);

public sealed record SyncRunFailure(
    string SourceCode,
    string SourceName,
    DateOnly QueryFrom,
    DateOnly QueryThrough,
    DateTimeOffset StartedAt,
    DateTimeOffset FinishedAt,
    string ErrorSummary);

public sealed record SyncSourceStatus(
    string SourceCode,
    string SourceName,
    DateOnly QueryFrom,
    DateOnly QueryThrough,
    string Status,
    DateTimeOffset StartedAt,
    DateTimeOffset FinishedAt,
    int PageCount,
    int RecordCount,
    DateTimeOffset? LastSuccessAt,
    string ErrorSummary);

public sealed class SyncStatusStore
{
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> FileLocks =
        new(StringComparer.OrdinalIgnoreCase);
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path;
    private readonly SemaphoreSlim _fileLock;

    public SyncStatusStore(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("同步狀態檔案路徑不可空白。", nameof(path));
        }

        _path = Path.GetFullPath(path);
        _fileLock = FileLocks.GetOrAdd(_path, _ => new SemaphoreSlim(1, 1));
    }

    public async Task<SyncSourceStatus?> GetAsync(
        string sourceCode,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceCode);
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            var directory = Path.GetDirectoryName(_path)
                ?? throw new InvalidOperationException("同步狀態檔案缺少目錄。");
            using var dataLock = await SourceSyncFileLock.AcquireAsync(
                directory, "sync-status-data", TimeSpan.FromSeconds(10), cancellationToken);
            var statuses = await LoadAsync(cancellationToken);
            return statuses.TryGetValue(sourceCode, out var status) ? status : null;
        }
        finally
        {
            _fileLock.Release();
        }
    }

    public async Task RecordSuccessAsync(
        SyncRunSummary run,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(run);
        ValidateRun(run.SourceCode, run.QueryFrom, run.QueryThrough, run.StartedAt, run.FinishedAt);
        if (run.PageCount < 0 || run.RecordCount < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(run), "分頁數與筆數不可小於零。");
        }

        await UpdateAsync(run.SourceCode, previous => new SyncSourceStatus(
            run.SourceCode,
            run.SourceName,
            run.QueryFrom,
            run.QueryThrough,
            "success",
            run.StartedAt,
            run.FinishedAt,
            run.PageCount,
            run.RecordCount,
            run.FinishedAt,
            string.Empty), cancellationToken);
    }

    public async Task RecordFailureAsync(
        SyncRunFailure run,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(run);
        ValidateRun(run.SourceCode, run.QueryFrom, run.QueryThrough, run.StartedAt, run.FinishedAt);

        await UpdateAsync(run.SourceCode, previous => new SyncSourceStatus(
            run.SourceCode,
            run.SourceName,
            run.QueryFrom,
            run.QueryThrough,
            "failed",
            run.StartedAt,
            run.FinishedAt,
            0,
            0,
            previous?.LastSuccessAt,
            Limit(run.ErrorSummary, 300)), cancellationToken);
    }

    private async Task UpdateAsync(
        string sourceCode,
        Func<SyncSourceStatus?, SyncSourceStatus> update,
        CancellationToken cancellationToken)
    {
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            var directory = Path.GetDirectoryName(_path)
                ?? throw new InvalidOperationException("同步狀態檔案缺少目錄。");
            using var dataLock = await SourceSyncFileLock.AcquireAsync(
                directory, "sync-status-data", TimeSpan.FromSeconds(10), cancellationToken);
            var statuses = await LoadAsync(cancellationToken);
            statuses.TryGetValue(sourceCode, out var previous);
            statuses[sourceCode] = update(previous);
            await SaveAsync(statuses, cancellationToken);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    private async Task<Dictionary<string, SyncSourceStatus>> LoadAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_path))
        {
            return new Dictionary<string, SyncSourceStatus>(StringComparer.OrdinalIgnoreCase);
        }

        await using var stream = new FileStream(
            _path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
        var stored = await JsonSerializer.DeserializeAsync<Dictionary<string, SyncSourceStatus>>(
            stream, JsonOptions, cancellationToken);
        return stored is null
            ? new Dictionary<string, SyncSourceStatus>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, SyncSourceStatus>(stored, StringComparer.OrdinalIgnoreCase);
    }

    private async Task SaveAsync(
        IReadOnlyDictionary<string, SyncSourceStatus> statuses,
        CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(_path)
            ?? throw new InvalidOperationException("同步狀態檔案缺少目錄。");
        Directory.CreateDirectory(directory);
        var temporaryPath = _path + $".{Guid.NewGuid():N}.tmp";
        try
        {
            await using (var stream = new FileStream(
                             temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            {
                await JsonSerializer.SerializeAsync(stream, statuses, JsonOptions, cancellationToken);
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

    private static void ValidateRun(
        string sourceCode,
        DateOnly queryFrom,
        DateOnly queryThrough,
        DateTimeOffset startedAt,
        DateTimeOffset finishedAt)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceCode);
        if (queryThrough < queryFrom)
        {
            throw new ArgumentException("同步查詢迄日不可早於起日。");
        }

        if (finishedAt < startedAt)
        {
            throw new ArgumentException("同步完成時間不可早於開始時間。");
        }
    }

    private static string Limit(string? value, int maximumLength)
    {
        var normalized = string.IsNullOrWhiteSpace(value) ? "同步失敗" : value.Trim();
        return normalized.Length <= maximumLength ? normalized : normalized[..maximumLength];
    }
}
