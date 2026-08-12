using System.Collections.Concurrent;
using System.Text.Json;

namespace MidAutumnGiftBox.Core;

public sealed record ManualOverrideState(
    string? LastWorkbookPath,
    IReadOnlyList<WorkbookEditableRow> LastExportedRows,
    IReadOnlyList<OrderRowManualOverride> Overrides,
    IReadOnlyList<WorkbookRowAutomaticReset>? PendingAutomaticResets = null)
{
    public static ManualOverrideState Empty { get; } = new(null, [], []);
}

public sealed class ManualOverrideStore
{
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> FileLocks =
        new(StringComparer.OrdinalIgnoreCase);
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path;
    private readonly SemaphoreSlim _fileLock;

    public ManualOverrideStore(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        _path = Path.GetFullPath(path);
        _fileLock = FileLocks.GetOrAdd(_path, _ => new SemaphoreSlim(1, 1));
    }

    public async Task<ManualOverrideState> LoadAsync(CancellationToken cancellationToken = default)
    {
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            if (!File.Exists(_path))
            {
                return ManualOverrideState.Empty;
            }

            await using var stream = new FileStream(
                _path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
            return await JsonSerializer.DeserializeAsync<ManualOverrideState>(
                       stream, JsonOptions, cancellationToken) ?? ManualOverrideState.Empty;
        }
        finally
        {
            _fileLock.Release();
        }
    }

    public async Task SaveAsync(
        ManualOverrideState state,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(state);
        var normalized = state with
        {
            LastWorkbookPath = string.IsNullOrWhiteSpace(state.LastWorkbookPath)
                ? null
                : Path.GetFullPath(state.LastWorkbookPath),
            LastExportedRows = state.LastExportedRows.ToArray(),
            Overrides = state.Overrides.ToArray(),
            PendingAutomaticResets = state.PendingAutomaticResets?.ToArray() ?? []
        };

        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            var directory = Path.GetDirectoryName(_path)
                ?? throw new InvalidOperationException("人工覆寫資料路徑缺少資料夾。");
            Directory.CreateDirectory(directory);
            var temporaryPath = _path + $".{Guid.NewGuid():N}.tmp";
            try
            {
                await using (var stream = new FileStream(
                                 temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None,
                                 4096, useAsync: true))
                {
                    await JsonSerializer.SerializeAsync(stream, normalized, JsonOptions, cancellationToken);
                }

                File.Move(temporaryPath, _path, overwrite: true);
            }
            finally
            {
                if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
            }
        }
        finally
        {
            _fileLock.Release();
        }
    }
}
