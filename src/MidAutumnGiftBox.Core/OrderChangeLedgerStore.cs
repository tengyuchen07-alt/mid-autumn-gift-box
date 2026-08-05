using System.Collections.Concurrent;
using System.Text.Json;

namespace MidAutumnGiftBox.Core;

public sealed class OrderChangeLedgerStore
{
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> FileLocks =
        new(StringComparer.OrdinalIgnoreCase);
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path;
    private readonly SemaphoreSlim _fileLock;

    public OrderChangeLedgerStore(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        _path = Path.GetFullPath(path);
        _fileLock = FileLocks.GetOrAdd(_path, _ => new SemaphoreSlim(1, 1));
    }

    public async Task<IReadOnlyList<OrderChangeEntry>> GetAllAsync(
        CancellationToken cancellationToken = default)
    {
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            return await LoadAsync(cancellationToken);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    public async Task AppendAsync(
        IReadOnlyList<OrderChangeEntry> entries,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(entries);
        if (entries.Count == 0)
        {
            return;
        }

        if (entries.Any(entry => string.IsNullOrWhiteSpace(entry.EntryId)))
        {
            throw new InvalidDataException("異動紀錄缺少識別碼，無法安全追加。");
        }

        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            var directory = Path.GetDirectoryName(_path)
                ?? throw new InvalidOperationException("無法判斷異動紀錄的資料夾。");
            using var dataLock = await SourceSyncFileLock.AcquireAsync(
                directory, "order-change-ledger-data", TimeSpan.FromSeconds(10), cancellationToken);
            var existing = await LoadAsync(cancellationToken);
            var known = existing
                .Select(entry => entry.EntryId)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var combined = existing.ToList();
            foreach (var entry in entries)
            {
                if (known.Add(entry.EntryId))
                {
                    combined.Add(entry with { Confirmation = string.Empty });
                }
            }

            await SaveAsync(combined, cancellationToken);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    private async Task<List<OrderChangeEntry>> LoadAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_path))
        {
            return [];
        }

        await using var stream = new FileStream(
            _path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
        return await JsonSerializer.DeserializeAsync<List<OrderChangeEntry>>(
                   stream, JsonOptions, cancellationToken) ?? [];
    }

    private async Task SaveAsync(
        IReadOnlyList<OrderChangeEntry> entries,
        CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(_path)
            ?? throw new InvalidOperationException("無法判斷異動紀錄的資料夾。");
        Directory.CreateDirectory(directory);
        var temporaryPath = _path + $".{Guid.NewGuid():N}.tmp";
        try
        {
            await using (var stream = new FileStream(
                             temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            {
                await JsonSerializer.SerializeAsync(stream, entries, JsonOptions, cancellationToken);
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
}
