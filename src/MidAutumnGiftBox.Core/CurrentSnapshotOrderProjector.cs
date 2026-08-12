using System.Globalization;

namespace MidAutumnGiftBox.Core;

public static class CurrentSnapshotOrderProjector
{
    public static IReadOnlyList<OrderChangeEntry> Project(
        IReadOnlyList<OrderLineSnapshot> snapshots,
        IReadOnlyDictionary<string, string> sourceNames,
        DateTimeOffset observedAt)
    {
        ArgumentNullException.ThrowIfNull(snapshots);
        ArgumentNullException.ThrowIfNull(sourceNames);

        var entries = new List<OrderChangeEntry>();
        foreach (var sourceGroup in snapshots
                     .GroupBy(snapshot => snapshot.SourceCode, StringComparer.OrdinalIgnoreCase)
                     .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            if (!sourceNames.TryGetValue(sourceGroup.Key, out var sourceName) ||
                string.IsNullOrWhiteSpace(sourceName))
            {
                throw new KeyNotFoundException($"找不到資料來源 {sourceGroup.Key} 的顯示名稱。");
            }

            var sourceLines = sourceGroup.ToArray();
            var parentEntries = OrderChangePlanner.Plan(
                    sourceName,
                    [],
                    sourceLines,
                    new Dictionary<string, OrderStatusResolution>(StringComparer.OrdinalIgnoreCase),
                    observedAt)
                .Select(entry => entry with { ChangeType = "目前訂單" })
                .ToArray();
            entries.AddRange(ApplyDependentSingleRules(ProjectMainItems(parentEntries, sourceLines)));
        }

        return entries
            .OrderBy(entry => entry.SourceCode, StringComparer.OrdinalIgnoreCase)
            .ThenBy(entry => entry.ExternalOrderNo, StringComparer.OrdinalIgnoreCase)
            .ThenBy(entry => entry.ExternalLineKey, StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static IReadOnlyList<OrderChangeEntry> ProjectMainItems(
        IReadOnlyList<OrderChangeEntry> parentEntries,
        IReadOnlyList<OrderLineSnapshot> snapshots)
    {
        var parents = snapshots
            .Where(line => line.LineLevel.Equals("parent", StringComparison.OrdinalIgnoreCase))
            .ToDictionary(ParentIdentity, StringComparer.OrdinalIgnoreCase);
        var childrenByParent = snapshots
            .Where(line => line.LineLevel.Equals("item", StringComparison.OrdinalIgnoreCase) &&
                           !string.IsNullOrWhiteSpace(line.ParentLineKey))
            .GroupBy(line => ParentIdentity(line.SourceCode, line.ExternalOrderNo, line.ParentLineKey!),
                StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.OrdinalIgnoreCase);
        var result = new List<OrderChangeEntry>();

        foreach (var entry in parentEntries)
        {
            var identity = ParentIdentity(entry.SourceCode, entry.ExternalOrderNo, entry.ExternalLineKey);
            if (!parents.TryGetValue(identity, out var parent))
            {
                continue;
            }

            childrenByParent.TryGetValue(identity, out var children);
            var itemGroups = (children ?? [])
                .Where(child => GiftBoxItemPolicy.TryResolve(child.ItemNo, out _))
                .GroupBy(child => child.ItemNo!.Trim(), StringComparer.Ordinal)
                .OrderBy(group => group.Key, StringComparer.Ordinal)
                .ToArray();
            if (itemGroups.Length > 0)
            {
                var legacyItemNo = itemGroups.Length == 1 ? itemGroups[0].Key : null;
                for (var index = 0; index < itemGroups.Length; index++)
                {
                    var group = itemGroups[index];
                    GiftBoxItemPolicy.TryResolve(group.Key, out var definition);
                    var quantity = group.Sum(child => child.Quantity);
                    var usesLegacyKey = group.Key.Equals(legacyItemNo, StringComparison.Ordinal);
                    result.Add(entry with
                    {
                        EntryId = usesLegacyKey ? entry.EntryId : $"{entry.EntryId}:{group.Key}",
                        ExternalLineKey = usesLegacyKey
                            ? entry.ExternalLineKey
                            : $"{entry.ExternalLineKey}:item_no:{group.Key}",
                        Sku = group.Key,
                        ProductName = group.First().ProductName,
                        PreviousQuantity = 0m,
                        NewQuantity = quantity,
                        QuantityChange = quantity,
                        GiftBoxSize = definition.PieceCount
                    });
                }
                continue;
            }

            if (GiftBoxItemPolicy.TryResolve(parent.ItemNo, out var parentDefinition))
            {
                result.Add(entry with
                {
                    Sku = parentDefinition.ItemNo,
                    GiftBoxSize = parentDefinition.PieceCount
                });
            }
        }

        return result;
    }

    private static IReadOnlyList<OrderChangeEntry> ApplyDependentSingleRules(
        IReadOnlyList<OrderChangeEntry> entries)
    {
        var result = new List<OrderChangeEntry>();
        foreach (var order in entries.GroupBy(
                     entry => entry.ExternalOrderNo,
                     StringComparer.OrdinalIgnoreCase))
        {
            var orderEntries = order.ToArray();
            var anchors = orderEntries
                .Where(entry => HasRole(entry, GiftBoxItemRole.GiftDateAnchor))
                .ToArray();
            if (anchors.Length == 0)
            {
                result.AddRange(orderEntries.Where(entry =>
                    !HasRole(entry, GiftBoxItemRole.DependentSingle) || HasValidDate(entry)));
                continue;
            }

            var earliestAnchor = anchors
                .Select(entry => new
                {
                    Entry = entry,
                    Date = ParseDate(entry.OriginalDeliveryDate) ?? ParseDate(entry.DeliveryDate)
                })
                .Where(item => item.Date is not null)
                .OrderBy(item => item.Date)
                .Select(item => item.Entry)
                .FirstOrDefault();

            foreach (var entry in orderEntries)
            {
                if (!HasRole(entry, GiftBoxItemRole.DependentSingle) ||
                    ParseDate(entry.OriginalDeliveryDate) is not null ||
                    ParseDate(entry.DeliveryDate) is not null ||
                    earliestAnchor is null)
                {
                    result.Add(entry);
                    continue;
                }

                result.Add(entry with
                {
                    OriginalDeliveryDate = earliestAnchor.OriginalDeliveryDate,
                    DeliveryDate = earliestAnchor.DeliveryDate
                });
            }
        }

        return result;
    }

    private static bool HasRole(OrderChangeEntry entry, GiftBoxItemRole role) =>
        GiftBoxItemPolicy.TryResolve(entry.Sku, out var definition) && definition.Role == role;

    private static bool HasValidDate(OrderChangeEntry entry) =>
        ParseDate(entry.OriginalDeliveryDate) is not null || ParseDate(entry.DeliveryDate) is not null;

    private static DateOnly? ParseDate(string? value) =>
        DateOnly.TryParseExact(
            value?.Trim(),
            ["yyyy-M-d", "yyyy-MM-dd", "yyyy/M/d", "yyyy/MM/dd"],
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out var date)
            ? date
            : null;

    private static string ParentIdentity(OrderLineSnapshot line) =>
        ParentIdentity(line.SourceCode, line.ExternalOrderNo, line.ExternalLineKey);

    private static string ParentIdentity(string sourceCode, string externalOrderNo, string lineKey) =>
        $"{sourceCode}\u001f{externalOrderNo}\u001f{lineKey}";
}
