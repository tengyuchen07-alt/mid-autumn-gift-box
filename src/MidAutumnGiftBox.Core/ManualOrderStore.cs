using System.Collections.Concurrent;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace MidAutumnGiftBox.Core;

public sealed record ManualOrderInput(
    string ChannelName,
    string Sku,
    int Quantity,
    DateOnly OriginalArrivalDate,
    bool IsRegistered = false,
    string? Note = null,
    string? ExternalOrderNo = null);

public sealed record ManualOrderActor(string UserName, string ComputerName)
{
    public static ManualOrderActor Capture() => new(
        string.IsNullOrWhiteSpace(Environment.UserName) ? "unknown" : Environment.UserName,
        string.IsNullOrWhiteSpace(Environment.MachineName) ? "unknown" : Environment.MachineName);
}

public sealed record ManualOrder(
    string Id,
    string SourceCode,
    string ChannelName,
    string Sku,
    string ProductName,
    int Quantity,
    DateOnly OriginalArrivalDate,
    bool IsRegistered,
    string? Note,
    string? ExternalOrderNo,
    string CreatedBy,
    string CreatedOnComputer,
    DateTimeOffset CreatedAt,
    string ModifiedBy,
    string ModifiedOnComputer,
    DateTimeOffset ModifiedAt,
    bool IsVoided = false);

public sealed record ManualOrderAuditEntry(
    string AuditId,
    string ManualOrderId,
    string Action,
    bool? BeforeRegistered,
    bool? AfterRegistered,
    string Actor,
    string ComputerName,
    DateTimeOffset OccurredAt);

public static class ManualGiftBoxCatalog
{
    private static readonly IReadOnlyDictionary<string, string> Products =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["4710964232435"] = "蛋黃酥3入禮盒",
            ["4710964232411"] = "蛋黃酥6入禮盒",
            ["4710964232565"] = "蛋黃酥9入禮盒"
        };

    public static IReadOnlyList<(string Sku, string ProductName)> Options =>
        Products.Select(product => (product.Key, product.Value)).ToArray();

    public static bool TryGetProductName(string sku, out string productName) =>
        Products.TryGetValue(sku, out productName!);
}

public sealed class ManualOrderStore
{
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> FileLocks =
        new(StringComparer.OrdinalIgnoreCase);
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _path;
    private readonly SemaphoreSlim _fileLock;

    public ManualOrderStore(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        _path = Path.GetFullPath(path);
        _fileLock = FileLocks.GetOrAdd(_path, _ => new SemaphoreSlim(1, 1));
    }

    public async Task<ManualOrder> AddAsync(
        ManualOrderInput input,
        ManualOrderActor actor,
        DateTimeOffset occurredAt,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input);
        ArgumentNullException.ThrowIfNull(actor);
        var normalized = Validate(input);
        var normalizedActor = NormalizeActor(actor);

        return await UpdateAsync(document =>
        {
            var order = new ManualOrder(
                Guid.NewGuid().ToString("N"),
                "manual",
                normalized.ChannelName,
                normalized.Sku,
                normalized.ProductName,
                normalized.Quantity,
                normalized.OriginalArrivalDate,
                normalized.IsRegistered,
                normalized.Note,
                normalized.ExternalOrderNo,
                normalizedActor.UserName,
                normalizedActor.ComputerName,
                occurredAt,
                normalizedActor.UserName,
                normalizedActor.ComputerName,
                occurredAt);
            document.Orders.Add(order);
            document.Audit.Add(new ManualOrderAuditEntry(
                Guid.NewGuid().ToString("N"),
                order.Id,
                "created",
                null,
                order.IsRegistered,
                normalizedActor.UserName,
                normalizedActor.ComputerName,
                occurredAt));
            return order;
        }, cancellationToken);
    }

    public async Task<ManualOrder> SetRegistrationAsync(
        string id,
        bool isRegistered,
        ManualOrderActor actor,
        DateTimeOffset occurredAt,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(id);
        ArgumentNullException.ThrowIfNull(actor);
        var normalizedActor = NormalizeActor(actor);
        return await UpdateAsync(document =>
        {
            var index = document.Orders.FindIndex(order =>
                order.Id.Equals(id, StringComparison.OrdinalIgnoreCase));
            if (index < 0 || document.Orders[index].IsVoided)
            {
                throw new KeyNotFoundException("找不到可更新的手打單。");
            }

            var previous = document.Orders[index];
            if (previous.IsRegistered == isRegistered)
            {
                return previous;
            }

            var updated = previous with
            {
                IsRegistered = isRegistered,
                ModifiedBy = normalizedActor.UserName,
                ModifiedOnComputer = normalizedActor.ComputerName,
                ModifiedAt = occurredAt
            };
            document.Orders[index] = updated;
            document.Audit.Add(new ManualOrderAuditEntry(
                Guid.NewGuid().ToString("N"),
                updated.Id,
                "registration_changed",
                previous.IsRegistered,
                updated.IsRegistered,
                normalizedActor.UserName,
                normalizedActor.ComputerName,
                occurredAt));
            return updated;
        }, cancellationToken);
    }

    public async Task<IReadOnlyList<ManualOrder>> GetActiveAsync(
        CancellationToken cancellationToken = default)
    {
        var document = await ReadAsync(cancellationToken);
        return document.Orders
            .Where(order => !order.IsVoided)
            .OrderBy(order => order.CreatedAt)
            .ToArray();
    }

    public async Task<IReadOnlyList<ManualOrderAuditEntry>> GetAuditAsync(
        CancellationToken cancellationToken = default)
    {
        var document = await ReadAsync(cancellationToken);
        return document.Audit.OrderBy(entry => entry.OccurredAt).ToArray();
    }

    public async Task<IReadOnlyList<OrderChangeEntry>> GetExportEntriesAsync(
        CancellationToken cancellationToken = default)
    {
        var orders = await GetActiveAsync(cancellationToken);
        return orders.Select(ToExportEntry).ToArray();
    }

    private static OrderChangeEntry ToExportEntry(ManualOrder order)
    {
        var identity = $"manual\u001f{order.Id}";
        var entryId = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)));
        return new OrderChangeEntry(
            entryId,
            order.CreatedAt,
            "manual",
            "手打單",
            string.IsNullOrWhiteSpace(order.ExternalOrderNo) ? $"手打-{order.Id[..8]}" : order.ExternalOrderNo,
            $"manual:{order.Id}",
            order.ChannelName,
            order.CreatedAt.ToString("yyyy/MM/dd HH:mm:ss", CultureInfo.InvariantCulture),
            ShippingLeadTimePolicy.Adjust(order.OriginalArrivalDate)
                .ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            order.Sku,
            order.ProductName,
            0m,
            order.Quantity,
            order.Quantity,
            "手打單",
            order.IsRegistered ? "MANUAL_REGISTERED" : "MANUAL_PENDING",
            order.IsRegistered ? "已登記" : "未登記",
            false,
            string.Empty,
            order.OriginalArrivalDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
    }

    private async Task<ManualOrderDocument> ReadAsync(CancellationToken cancellationToken)
    {
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            var directory = Path.GetDirectoryName(_path)
                ?? throw new InvalidOperationException("手打單資料檔缺少目錄。");
            using var dataLock = await SourceSyncFileLock.AcquireAsync(
                directory, "manual-orders-data", TimeSpan.FromSeconds(10), cancellationToken);
            return await LoadAsync(cancellationToken);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    private async Task<T> UpdateAsync<T>(
        Func<ManualOrderDocument, T> update,
        CancellationToken cancellationToken)
    {
        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            var directory = Path.GetDirectoryName(_path)
                ?? throw new InvalidOperationException("手打單資料檔缺少目錄。");
            using var dataLock = await SourceSyncFileLock.AcquireAsync(
                directory, "manual-orders-data", TimeSpan.FromSeconds(10), cancellationToken);
            var document = await LoadAsync(cancellationToken);
            var result = update(document);
            await SaveAsync(document, cancellationToken);
            return result;
        }
        finally
        {
            _fileLock.Release();
        }
    }

    private async Task<ManualOrderDocument> LoadAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_path))
        {
            return new ManualOrderDocument();
        }

        await using var stream = new FileStream(
            _path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
        return await JsonSerializer.DeserializeAsync<ManualOrderDocument>(
                   stream, JsonOptions, cancellationToken) ?? new ManualOrderDocument();
    }

    private async Task SaveAsync(ManualOrderDocument document, CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(_path)
            ?? throw new InvalidOperationException("手打單資料檔缺少目錄。");
        Directory.CreateDirectory(directory);
        var temporaryPath = _path + $".{Guid.NewGuid():N}.tmp";
        try
        {
            await using (var stream = new FileStream(
                             temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            {
                await JsonSerializer.SerializeAsync(stream, document, JsonOptions, cancellationToken);
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

    private static NormalizedManualOrderInput Validate(ManualOrderInput input)
    {
        var channel = input.ChannelName?.Trim() ?? string.Empty;
        if (channel.Length is 0 or > 100)
        {
            throw new ArgumentException("通路別必填且不可超過 100 字。", nameof(input));
        }

        var sku = input.Sku?.Trim() ?? string.Empty;
        if (!ManualGiftBoxCatalog.TryGetProductName(sku, out var productName))
        {
            throw new ArgumentException("手打單禮盒須選擇三入、六入或九入。", nameof(input));
        }

        if (input.Quantity <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(input), "手打單數量須為正整數。");
        }

        var note = string.IsNullOrWhiteSpace(input.Note) ? null : input.Note.Trim();
        if (note?.Length > 500)
        {
            throw new ArgumentException("手打單備註不可超過 500 字。", nameof(input));
        }

        var externalOrderNo = string.IsNullOrWhiteSpace(input.ExternalOrderNo)
            ? null
            : input.ExternalOrderNo.Trim();
        if (externalOrderNo?.Length > 100)
        {
            throw new ArgumentException("外部訂單編號不可超過 100 字。", nameof(input));
        }

        return new NormalizedManualOrderInput(
            channel,
            sku,
            productName,
            input.Quantity,
            input.OriginalArrivalDate,
            input.IsRegistered,
            note,
            externalOrderNo);
    }

    private static ManualOrderActor NormalizeActor(ManualOrderActor actor) => new(
        string.IsNullOrWhiteSpace(actor.UserName) ? "unknown" : actor.UserName.Trim(),
        string.IsNullOrWhiteSpace(actor.ComputerName) ? "unknown" : actor.ComputerName.Trim());

    private sealed record NormalizedManualOrderInput(
        string ChannelName,
        string Sku,
        string ProductName,
        int Quantity,
        DateOnly OriginalArrivalDate,
        bool IsRegistered,
        string? Note,
        string? ExternalOrderNo);

    private sealed class ManualOrderDocument
    {
        public List<ManualOrder> Orders { get; init; } = [];
        public List<ManualOrderAuditEntry> Audit { get; init; } = [];
    }
}
