using System.Security.Cryptography;
using System.Text;

namespace MidAutumnGiftBox.Core;

public sealed record OrderStatusResolution(
    string ExternalOrderNo,
    string? StatusCode,
    string? StatusName);

public sealed record OrderChangeEntry(
    string EntryId,
    DateTimeOffset ObservedAt,
    string SourceCode,
    string SourceName,
    string ExternalOrderNo,
    string ExternalLineKey,
    string? ChannelName,
    string? OrderDate,
    string? DeliveryDate,
    string? Sku,
    string? ProductName,
    decimal PreviousQuantity,
    decimal NewQuantity,
    decimal QuantityChange,
    string ChangeType,
    string? StatusCode,
    string? StatusName,
    bool NeedsReview,
    string Confirmation = "",
    string? OriginalDeliveryDate = null,
    int? GiftBoxSize = null,
    string? Note = null,
    string? SourceLocation = null,
    bool HasManualLocation = false,
    bool HasManualConfirmation = false,
    bool HasManualOriginalDeliveryDate = false,
    bool HasManualOrderDate = false,
    bool IgnoreExistingLocation = false,
    bool IgnoreExistingConfirmation = false);

public static class OrderChangePlanner
{
    public static IReadOnlyList<OrderChangeEntry> Plan(
        string sourceName,
        IReadOnlyList<OrderLineSnapshot> previous,
        IReadOnlyList<OrderLineSnapshot> current,
        IReadOnlyDictionary<string, OrderStatusResolution> departedOrderStatuses,
        DateTimeOffset observedAt)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceName);
        ArgumentNullException.ThrowIfNull(previous);
        ArgumentNullException.ThrowIfNull(current);
        ArgumentNullException.ThrowIfNull(departedOrderStatuses);

        var oldParents = previous
            .Where(IsParent)
            .ToDictionary(LineKey, StringComparer.OrdinalIgnoreCase);
        var newParents = current
            .Where(IsParent)
            .ToDictionary(LineKey, StringComparer.OrdinalIgnoreCase);
        var currentOrders = current
            .Select(line => line.ExternalOrderNo)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var entries = new List<OrderChangeEntry>();

        foreach (var currentLine in newParents.Values)
        {
            oldParents.TryGetValue(LineKey(currentLine), out var previousLine);
            var oldQuantity = previousLine?.Quantity ?? 0m;
            var change = currentLine.Quantity - oldQuantity;
            if (change == 0m)
            {
                continue;
            }

            entries.Add(CreateEntry(
                sourceName,
                currentLine,
                oldQuantity,
                currentLine.Quantity,
                change,
                previousLine is null ? "新增" : change > 0m ? "數量增加" : "數量減少",
                currentLine.StatusCode,
                currentLine.StatusName,
                needsReview: false,
                observedAt,
                current));
        }

        foreach (var previousLine in oldParents.Values.Where(line => !newParents.ContainsKey(LineKey(line))))
        {
            if (currentOrders.Contains(previousLine.ExternalOrderNo))
            {
                entries.Add(CreateEntry(
                    sourceName,
                    previousLine,
                    previousLine.Quantity,
                    0m,
                    -previousLine.Quantity,
                    "品項移除",
                    "F",
                    "待處理",
                    needsReview: false,
                    observedAt,
                    previous));
                continue;
            }

            departedOrderStatuses.TryGetValue(previousLine.ExternalOrderNo, out var resolution);
            var statusCode = resolution?.StatusCode?.Trim().ToUpperInvariant();
            switch (statusCode)
            {
                case "N":
                    entries.Add(CreateEntry(
                        sourceName, previousLine, previousLine.Quantity, 0m, -previousLine.Quantity,
                        "取消", statusCode, resolution?.StatusName, needsReview: false, observedAt, previous));
                    break;
                case "R":
                    entries.Add(CreateEntry(
                        sourceName, previousLine, previousLine.Quantity, 0m, -previousLine.Quantity,
                        "退貨", statusCode, resolution?.StatusName, needsReview: false, observedAt, previous));
                    break;
                case "D":
                    entries.Add(CreateEntry(
                        sourceName, previousLine, previousLine.Quantity, previousLine.Quantity, 0m,
                        "刪除或併單待確認", statusCode, resolution?.StatusName, needsReview: true, observedAt, previous));
                    break;
                case "F":
                    entries.Add(CreateEntry(
                        sourceName, previousLine, previousLine.Quantity, 0m, -previousLine.Quantity,
                        "品項移除", statusCode, resolution?.StatusName, needsReview: false, observedAt, previous));
                    break;
                case null or "":
                    entries.Add(CreateEntry(
                        sourceName, previousLine, previousLine.Quantity, previousLine.Quantity, 0m,
                        "狀態待確認", statusCode, resolution?.StatusName, needsReview: true, observedAt, previous));
                    break;
                // W/P/C/S/T/A and other known workflow states do not reverse factory demand.
            }
        }

        return entries
            .OrderBy(entry => entry.SourceCode, StringComparer.OrdinalIgnoreCase)
            .ThenBy(entry => entry.ExternalOrderNo, StringComparer.OrdinalIgnoreCase)
            .ThenBy(entry => entry.ExternalLineKey, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static bool IsParent(OrderLineSnapshot line) =>
        line.LineLevel.Equals("parent", StringComparison.OrdinalIgnoreCase);

    private static string LineKey(OrderLineSnapshot line) =>
        $"{line.SourceCode}\u001f{line.ExternalOrderNo}\u001f{line.ExternalLineKey}";

    private static OrderChangeEntry CreateEntry(
        string sourceName,
        OrderLineSnapshot line,
        decimal previousQuantity,
        decimal newQuantity,
        decimal quantityChange,
        string changeType,
        string? statusCode,
        string? statusName,
        bool needsReview,
        DateTimeOffset observedAt,
        IReadOnlyList<OrderLineSnapshot> sourceSnapshot)
    {
        var identity = string.Join("\u001f",
            line.SourceCode,
            line.ExternalOrderNo,
            line.ExternalLineKey,
            line.SynchronizedAt.ToUniversalTime().ToString("O"),
            previousQuantity,
            newQuantity,
            changeType,
            statusCode);
        var entryId = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)));
        return new OrderChangeEntry(
            entryId,
            observedAt,
            line.SourceCode,
            sourceName,
            line.ExternalOrderNo,
            line.ExternalLineKey,
            line.ChannelName ?? line.ChannelCode,
            line.OrderDate,
            ResolveDeliveryDate(line),
            line.Sku,
            line.ProductName,
            previousQuantity,
            newQuantity,
            quantityChange,
            changeType,
            statusCode,
            statusName,
            needsReview,
            string.Empty,
            ResolveOriginalDeliveryDate(line),
            ResolveGiftBoxSize(line, sourceSnapshot),
            line.Note);
    }

    private static int? ResolveGiftBoxSize(
        OrderLineSnapshot parent,
        IReadOnlyList<OrderLineSnapshot> sourceSnapshot) =>
        GiftBoxSizePolicy.ResolveParent(parent, sourceSnapshot);

    private static string? ResolveDeliveryDate(OrderLineSnapshot line) =>
        WmsDepartmentStoreDatePolicy.ResolveOrderDate(line);

    private static string? ResolveOriginalDeliveryDate(OrderLineSnapshot line) =>
        WmsDepartmentStoreDatePolicy.ResolveDeliveryDate(line);
}
