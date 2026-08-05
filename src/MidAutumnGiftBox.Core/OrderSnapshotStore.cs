using System.Collections.Concurrent;
using System.Globalization;
using System.Text.Json;

namespace MidAutumnGiftBox.Core;

public sealed record OrderLineSnapshot(
    string SourceCode,
    string ExternalOrderNo,
    string ExternalLineKey,
    string LineLevel,
    string? ParentLineKey,
    string? ChannelCode,
    string? ShopName,
    string? OrderDate,
    string? ArrivalDate,
    string? DerivedShippingDate,
    string? Sku,
    string? ItemNo,
    string? ProductName,
    string? Spec,
    decimal Quantity,
    decimal ShippedQuantity,
    string? StatusCode,
    DateTimeOffset SynchronizedAt,
    string? ChannelName = null,
    string? ShipWindowStart = null,
    string? ShipWindowEnd = null,
    string? ShippingDateSource = null,
    string? ShippingDateStatus = null,
    string? ProductType = null,
    string? StatusName = null,
    string? TotalPrice = null);

public sealed class OrderSnapshotStore
{
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> FileLocks =
        new(StringComparer.OrdinalIgnoreCase);
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };
    private readonly string _path;
    private readonly SemaphoreSlim _fileLock;

    public OrderSnapshotStore(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("訂單快照檔案路徑不可空白。", nameof(path));
        }

        _path = Path.GetFullPath(path);
        _fileLock = FileLocks.GetOrAdd(_path, _ => new SemaphoreSlim(1, 1));
    }

    public async Task<IReadOnlyList<OrderLineSnapshot>> GetAllAsync(
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

    public async Task ReplaceSourceAsync(
        string sourceCode,
        IReadOnlyList<IReadOnlyDictionary<string, string?>> rows,
        DateTimeOffset synchronizedAt,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceCode);
        ArgumentNullException.ThrowIfNull(rows);
        var replacement = BuildSourceSnapshot(sourceCode, rows, synchronizedAt);
        EnsureUniqueKeys(replacement);

        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            using var dataLock = await AcquireDataFileLockAsync(cancellationToken);
            var existing = await LoadAsync(cancellationToken);
            var combined = existing
                .Where(line => !line.SourceCode.Equals(sourceCode, StringComparison.OrdinalIgnoreCase))
                .Concat(replacement)
                .OrderBy(line => line.SourceCode, StringComparer.OrdinalIgnoreCase)
                .ThenBy(line => line.ExternalOrderNo, StringComparer.OrdinalIgnoreCase)
                .ThenBy(line => line.ExternalLineKey, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            await SaveAsync(combined, cancellationToken);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    public async Task ReplaceSourceRangeAsync(
        string sourceCode,
        DateOnly fromDate,
        DateOnly throughDate,
        IReadOnlyList<IReadOnlyDictionary<string, string?>> rows,
        DateTimeOffset synchronizedAt,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceCode);
        ArgumentNullException.ThrowIfNull(rows);
        if (throughDate < fromDate)
        {
            throw new ArgumentException("增量快照查詢迄日不可早於起日。");
        }

        var replacement = BuildSourceSnapshot(sourceCode, rows, synchronizedAt);
        EnsureUniqueKeys(replacement);
        foreach (var line in replacement)
        {
            var orderDate = ParseRequiredOrderDate(line);
            if (orderDate < fromDate || orderDate > throughDate)
            {
                throw new InvalidDataException(
                    $"訂單 {line.ExternalOrderNo} 的成立日不在本次同步範圍內，未提交本次同步。");
            }
        }

        await _fileLock.WaitAsync(cancellationToken);
        try
        {
            using var dataLock = await AcquireDataFileLockAsync(cancellationToken);
            var existing = await LoadAsync(cancellationToken);
            var retained = new List<OrderLineSnapshot>(existing.Count);
            foreach (var line in existing)
            {
                if (!line.SourceCode.Equals(sourceCode, StringComparison.OrdinalIgnoreCase))
                {
                    retained.Add(line);
                    continue;
                }

                var orderDate = ParseRequiredOrderDate(line);
                if (orderDate < fromDate || orderDate > throughDate)
                {
                    retained.Add(line);
                }
            }

            var combined = retained
                .Concat(replacement)
                .OrderBy(line => line.SourceCode, StringComparer.OrdinalIgnoreCase)
                .ThenBy(line => line.ExternalOrderNo, StringComparer.OrdinalIgnoreCase)
                .ThenBy(line => line.ExternalLineKey, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            EnsureUniqueKeys(combined);
            await SaveAsync(combined, cancellationToken);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    public static IReadOnlyList<OrderLineSnapshot> BuildSourceSnapshot(
        string sourceCode,
        IReadOnlyList<IReadOnlyDictionary<string, string?>> rows,
        DateTimeOffset synchronizedAt)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceCode);
        ArgumentNullException.ThrowIfNull(rows);
        var result = new List<OrderLineSnapshot>();
        foreach (var row in rows)
        {
            var orderNo = GetRowValue(row, "order_no")?.Trim();
            if (string.IsNullOrWhiteSpace(orderNo))
            {
                throw new InvalidDataException("訂單快照包含缺少訂單編號的資料，未提交本次同步。");
            }

            var productsJson = GetRowValue(row, "products");
            if (string.IsNullOrWhiteSpace(productsJson))
            {
                throw new InvalidDataException($"訂單 {orderNo} 缺少商品明細，未提交本次同步。");
            }

            JsonDocument document;
            try
            {
                document = JsonDocument.Parse(productsJson);
            }
            catch (JsonException exception)
            {
                throw new InvalidDataException($"訂單 {orderNo} 的商品明細格式無法解析，未提交本次同步。", exception);
            }

            using (document)
            {
                if (document.RootElement.ValueKind != JsonValueKind.Array)
                {
                    throw new InvalidDataException($"訂單 {orderNo} 的商品明細不是陣列，未提交本次同步。");
                }

                var parentOccurrences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var parentIndex = 0;
                foreach (var product in document.RootElement.EnumerateArray())
                {
                    parentIndex++;
                    if (product.ValueKind != JsonValueKind.Object)
                    {
                        throw new InvalidDataException($"訂單 {orderNo} 包含無效商品列，未提交本次同步。");
                    }

                    var parentIdentity = GetIdentity(product, parentIndex);
                    var parentOccurrence = NextOccurrence(parentOccurrences, parentIdentity.Key);
                    var parentLineKey = $"parent:{parentIdentity.Key}:{parentOccurrence}";
                    result.Add(CreateLine(
                        sourceCode, orderNo, parentLineKey, "parent", null, row, product, product, synchronizedAt));

                    if (!TryGetProperty(product, "items", out var items) ||
                        items.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
                    {
                        continue;
                    }

                    if (items.ValueKind != JsonValueKind.Array)
                    {
                        throw new InvalidDataException($"訂單 {orderNo} 的子商品明細不是陣列，未提交本次同步。");
                    }

                    var itemOccurrences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    var itemIndex = 0;
                    foreach (var item in items.EnumerateArray())
                    {
                        itemIndex++;
                        if (item.ValueKind != JsonValueKind.Object)
                        {
                            throw new InvalidDataException($"訂單 {orderNo} 包含無效子商品列，未提交本次同步。");
                        }

                        var itemIdentity = GetIdentity(item, itemIndex);
                        var itemOccurrence = NextOccurrence(itemOccurrences, itemIdentity.Key);
                        var itemLineKey = $"item:{parentLineKey}:{itemIdentity.Key}:{itemOccurrence}";
                        result.Add(CreateLine(
                            sourceCode, orderNo, itemLineKey, "item", parentLineKey, row, item, product, synchronizedAt));
                    }
                }
            }
        }

        return result;
    }

    private static void EnsureUniqueKeys(IReadOnlyList<OrderLineSnapshot> lines)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var line in lines)
        {
            var key = $"{line.SourceCode}\u001f{line.ExternalOrderNo}\u001f{line.ExternalLineKey}";
            if (!keys.Add(key))
            {
                throw new InvalidDataException(
                    $"訂單 {line.ExternalOrderNo} 包含重複商品列鍵，未提交本次同步。");
            }
        }
    }

    private static DateOnly ParseRequiredOrderDate(OrderLineSnapshot line)
    {
        if (!string.IsNullOrWhiteSpace(line.OrderDate) &&
            DateTime.TryParse(
                line.OrderDate,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces,
                out var parsed))
        {
            return DateOnly.FromDateTime(parsed);
        }

        throw new InvalidDataException(
            $"訂單 {line.ExternalOrderNo} 缺少有效成立日，未提交本次同步。");
    }

    private static OrderLineSnapshot CreateLine(
        string sourceCode,
        string orderNo,
        string lineKey,
        string lineLevel,
        string? parentLineKey,
        IReadOnlyDictionary<string, string?> row,
        JsonElement product,
        JsonElement shippingProduct,
        DateTimeOffset synchronizedAt) => new(
            sourceCode,
            orderNo,
            lineKey,
            lineLevel,
            parentLineKey,
            GetRowValue(row, "source_key"),
            GetRowValue(row, "shop_name"),
            GetRowValue(row, "order_date"),
            GetRowValue(row, "arrival_date"),
            GetProductOrRowValue(shippingProduct, row, "derived_shipping_date"),
            GetScalarText(product, "sku"),
            GetScalarText(product, "item_no"),
            GetScalarText(product, "name"),
            GetScalarText(product, "spec"),
            GetDecimal(product, "qty", "商品數量"),
            GetDecimal(product, "shipp_qty", "已出貨數量"),
            GetRowValue(row, "status_code"),
            synchronizedAt,
            GetRowValue(row, "source"),
            GetProductOrRowValue(shippingProduct, row, "ship_window_start"),
            GetProductOrRowValue(shippingProduct, row, "ship_window_end"),
            GetProductOrRowValue(shippingProduct, row, "shipping_date_source"),
            GetProductOrRowValue(shippingProduct, row, "shipping_date_status"),
            GetScalarText(product, "type"),
            GetRowValue(row, "status_name"),
            GetRowValue(row, "total_price"));

    private async Task<List<OrderLineSnapshot>> LoadAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(_path))
        {
            return [];
        }

        await using var stream = new FileStream(
            _path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
        return await JsonSerializer.DeserializeAsync<List<OrderLineSnapshot>>(
                   stream, JsonOptions, cancellationToken) ?? [];
    }

    private async Task SaveAsync(
        IReadOnlyList<OrderLineSnapshot> lines,
        CancellationToken cancellationToken)
    {
        var directory = Path.GetDirectoryName(_path)
            ?? throw new InvalidOperationException("訂單快照檔案缺少目錄。");
        Directory.CreateDirectory(directory);
        var temporaryPath = _path + $".{Guid.NewGuid():N}.tmp";
        try
        {
            await using (var stream = new FileStream(
                             temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true))
            {
                await JsonSerializer.SerializeAsync(stream, lines, JsonOptions, cancellationToken);
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
            ?? throw new InvalidOperationException("訂單快照檔案缺少目錄。");
        return SourceSyncFileLock.AcquireAsync(
            directory, "order-snapshot-data", TimeSpan.FromSeconds(10), cancellationToken);
    }

    private static (string Key, string? Value) GetIdentity(JsonElement product, int index)
    {
        var itemNo = GetScalarText(product, "item_no")?.Trim();
        if (!string.IsNullOrWhiteSpace(itemNo))
        {
            return ($"item_no:{itemNo}", itemNo);
        }

        var sku = GetScalarText(product, "sku")?.Trim();
        return !string.IsNullOrWhiteSpace(sku)
            ? ($"sku:{sku}", sku)
            : ($"index:{index}", null);
    }

    private static int NextOccurrence(IDictionary<string, int> occurrences, string identity)
    {
        occurrences.TryGetValue(identity, out var count);
        count++;
        occurrences[identity] = count;
        return count;
    }

    private static string? GetRowValue(IReadOnlyDictionary<string, string?> row, string name) =>
        row.FirstOrDefault(pair => pair.Key.Equals(name, StringComparison.OrdinalIgnoreCase)).Value;

    private static string? GetProductOrRowValue(
        JsonElement product,
        IReadOnlyDictionary<string, string?> row,
        string name)
    {
        var productValue = GetScalarText(product, name);
        return string.IsNullOrWhiteSpace(productValue) ? GetRowValue(row, name) : productValue;
    }

    private static string? GetScalarText(JsonElement element, string name)
    {
        if (!TryGetProperty(element, name, out var value) ||
            value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return null;
        }

        return value.ValueKind == JsonValueKind.String ? value.GetString() : value.ToString();
    }

    private static decimal GetDecimal(JsonElement element, string name, string displayName)
    {
        if (!TryGetProperty(element, name, out var value) ||
            value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return 0m;
        }

        if (value.ValueKind == JsonValueKind.Number && value.TryGetDecimal(out var number))
        {
            return number;
        }

        if (value.ValueKind == JsonValueKind.String &&
            decimal.TryParse(value.GetString(), NumberStyles.Number, CultureInfo.InvariantCulture, out number))
        {
            return number;
        }

        throw new InvalidDataException($"{displayName}格式無效，未提交本次同步。");
    }

    private static bool TryGetProperty(JsonElement element, string name, out JsonElement value)
    {
        foreach (var property in element.EnumerateObject())
        {
            if (property.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                value = property.Value;
                return true;
            }
        }

        value = default;
        return false;
    }
}
