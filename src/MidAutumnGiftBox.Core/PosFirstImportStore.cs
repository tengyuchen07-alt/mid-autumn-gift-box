using System.Collections.Concurrent;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace MidAutumnGiftBox.Core;

public sealed record PosFirstImportRecord(
    string ExternalOrderNo,
    string ExternalLineKey,
    string FirstImportedOn);

public sealed record PosFirstImportState(IReadOnlyList<PosFirstImportRecord> Records)
{
    public static PosFirstImportState Empty { get; } = new([]);
}

public sealed class PosFirstImportStore
{
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> FileLocks =
        new(StringComparer.OrdinalIgnoreCase);
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path;
    private readonly SemaphoreSlim _fileLock;

    public PosFirstImportStore(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        _path = Path.GetFullPath(path);
        _fileLock = FileLocks.GetOrAdd(_path, _ => new SemaphoreSlim(1, 1));
    }

    public async Task<ImportedSpreadsheetBatch> ImportAsync(
        string path,
        DateTimeOffset importedAt,
        CancellationToken cancellationToken = default)
    {
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            using var dataLock = await AcquireDataFileLockAsync(cancellationToken).ConfigureAwait(false);
            var state = await LoadUnlockedAsync(cancellationToken);
            var records = state.Records.ToDictionary(RecordKey, StringComparer.OrdinalIgnoreCase);
            var changed = false;
            DateOnly Resolve(string externalOrderNo, string externalLineKey)
            {
                var key = RecordKey(externalOrderNo, externalLineKey);
                if (records.TryGetValue(key, out var existing) &&
                    DateOnly.TryParseExact(existing.FirstImportedOn, "yyyy-MM-dd",
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out var firstImportedOn))
                {
                    return firstImportedOn;
                }

                var firstDate = DateOnly.FromDateTime(importedAt.Date);
                records[key] = new PosFirstImportRecord(
                    externalOrderNo.Trim(),
                    externalLineKey,
                    firstDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                changed = true;
                return firstDate;
            }

            var batch = await Task.Run(
                () => SpreadsheetOrderImporter.ReadPos(path, importedAt, Resolve),
                cancellationToken).ConfigureAwait(false);
            if (changed)
            {
                await SaveUnlockedAsync(
                    new PosFirstImportState(records.Values
                        .OrderBy(record => record.ExternalOrderNo, StringComparer.OrdinalIgnoreCase)
                        .ThenBy(record => record.ExternalLineKey, StringComparer.OrdinalIgnoreCase)
                        .ToArray()),
                    cancellationToken);
            }

            return batch;
        }
        finally
        {
            _fileLock.Release();
        }
    }

    private async Task<PosFirstImportState> LoadUnlockedAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_path)) return PosFirstImportState.Empty;
        await using var stream = new FileStream(
            _path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
        return await JsonSerializer.DeserializeAsync<PosFirstImportState>(
                   stream, JsonOptions, cancellationToken) ?? PosFirstImportState.Empty;
    }

    private async Task SaveUnlockedAsync(PosFirstImportState state, CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(_path)
            ?? throw new InvalidOperationException("POS 首次載入狀態路徑無效。");
        Directory.CreateDirectory(directory);
        var temporaryPath = _path + $".{Guid.NewGuid():N}.tmp";
        try
        {
            await using (var stream = new FileStream(
                             temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None,
                             4096, useAsync: true))
            {
                await JsonSerializer.SerializeAsync(stream, state, JsonOptions, cancellationToken);
            }
            File.Move(temporaryPath, _path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    private static string RecordKey(PosFirstImportRecord record) =>
        RecordKey(record.ExternalOrderNo, record.ExternalLineKey);

    private Task<SourceSyncFileLock> AcquireDataFileLockAsync(CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(_path)
            ?? throw new InvalidOperationException("POS 首次載入狀態檔案缺少目錄。");
        var pathHash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(_path)))[..24];
        return SourceSyncFileLock.AcquireAsync(
            directory, $"pos-first-import-{pathHash}", TimeSpan.FromSeconds(10), cancellationToken);
    }

    private static string RecordKey(string externalOrderNo, string externalLineKey) =>
        $"{externalOrderNo.Trim()}\u001f{externalLineKey}";
}

public static class PosSpreadsheetImportWorkflow
{
    public static Task<ImportedSpreadsheetBatch> ImportAsync(
        string path,
        DateTimeOffset importedAt,
        PosFirstImportStore store,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(store);
        return store.ImportAsync(path, importedAt, cancellationToken);
    }
}
