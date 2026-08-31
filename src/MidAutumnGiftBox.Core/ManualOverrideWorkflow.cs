using System.Security.Cryptography;
using System.Globalization;
using System.Text;

namespace MidAutumnGiftBox.Core;

public sealed record WorkbookEditableRow(
    string RowKey,
    string Sheet,
    IReadOnlyList<string> SourceLineFingerprints,
    string OriginalDeliveryDate,
    string OrderDate,
    string Location,
    string Confirmation,
    string ChannelName,
    string ExternalOrderNo,
    string ProductName,
    string AggregateId = "");

public sealed record ManualFieldOverride(bool IsOverridden, string Value)
{
    public static ManualFieldOverride None { get; } = new(false, string.Empty);
}

public sealed record OrderRowManualOverride(
    string RowKey,
    IReadOnlyList<string> SourceLineFingerprints,
    ManualFieldOverride OriginalDeliveryDate,
    ManualFieldOverride OrderDate,
    ManualFieldOverride Location,
    ManualFieldOverride Confirmation,
    DateTimeOffset UpdatedAt);

public sealed record WorkbookRowAutomaticReset(
    string RowKey,
    IReadOnlyList<string> SourceLineFingerprints);

public static class ManualOverrideWorkflow
{
    public static ManualOverrideState ContinueWithNewWorkbook(ManualOverrideState state)
    {
        ArgumentNullException.ThrowIfNull(state);
        return state with
        {
            LastWorkbookPath = null,
            LastExportedRows = []
        };
    }

    public static string BuildRowKey(IReadOnlyList<string> sourceLineFingerprints)
    {
        ArgumentNullException.ThrowIfNull(sourceLineFingerprints);
        var identity = string.Join(
            "\u001f",
            sourceLineFingerprints
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim().ToUpperInvariant())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase));
        if (identity.Length == 0)
        {
            throw new ArgumentException("訂單明細缺少穩定識別值，無法保存人工修改。", nameof(sourceLineFingerprints));
        }

        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)));
    }

    public static IReadOnlyList<OrderRowManualOverride> CaptureChanges(
        IReadOnlyList<WorkbookEditableRow> previousExport,
        IReadOnlyList<WorkbookEditableRow> currentWorkbook,
        IReadOnlyList<OrderRowManualOverride> existingOverrides,
        DateTimeOffset observedAt)
    {
        ArgumentNullException.ThrowIfNull(previousExport);
        ArgumentNullException.ThrowIfNull(currentWorkbook);
        ArgumentNullException.ThrowIfNull(existingOverrides);

        var result = existingOverrides.ToDictionary(item => item.RowKey, StringComparer.OrdinalIgnoreCase);
        var previousByKey = previousExport.ToDictionary(item => item.RowKey, StringComparer.OrdinalIgnoreCase);
        var previousByAggregateId = previousExport
            .Where(item => !string.IsNullOrWhiteSpace(item.AggregateId))
            .GroupBy(item => item.AggregateId, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        var previousManualRows = previousExport
            .Select(item => (Item: item, Identity: BuildManualRowMigrationIdentity(item)))
            .Where(item => item.Identity is not null)
            .GroupBy(item => item.Identity!, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single().Item, StringComparer.OrdinalIgnoreCase);
        foreach (var current in currentWorkbook)
        {
            if (!previousByKey.TryGetValue(current.RowKey, out var previous) &&
                (string.IsNullOrWhiteSpace(current.AggregateId) ||
                 !previousByAggregateId.TryGetValue(current.AggregateId, out previous)) &&
                (BuildManualRowMigrationIdentity(current) is not { } manualIdentity ||
                 !previousManualRows.TryGetValue(manualIdentity, out previous)))
            {
                continue;
            }

            result.TryGetValue(previous.RowKey, out var existing);
            var originalDeliveryDate = Changed(
                previous.OriginalDeliveryDate,
                current.OriginalDeliveryDate,
                existing?.OriginalDeliveryDate);
            var orderDate = Changed(previous.OrderDate, current.OrderDate, existing?.OrderDate);
            if (current.Sheet.Equals("pending", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(previous.OriginalDeliveryDate, current.OriginalDeliveryDate,
                    StringComparison.Ordinal) &&
                !string.IsNullOrWhiteSpace(current.OriginalDeliveryDate))
            {
                var specifiedDate = ParsePendingSpecifiedDate(current.OriginalDeliveryDate);
                originalDeliveryDate = new ManualFieldOverride(
                    true, specifiedDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                orderDate = new ManualFieldOverride(
                    true, specifiedDate.AddDays(-7).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
            }

            var updated = new OrderRowManualOverride(
                previous.RowKey,
                previous.SourceLineFingerprints.ToArray(),
                originalDeliveryDate,
                orderDate,
                Changed(previous.Location, current.Location, existing?.Location),
                Changed(previous.Confirmation, current.Confirmation, existing?.Confirmation),
                observedAt);
            if (HasAnyOverride(updated))
            {
                result[previous.RowKey] = updated;
            }
            else
            {
                result.Remove(previous.RowKey);
            }
        }

        return result.Values.OrderBy(item => item.RowKey, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static string? BuildManualRowMigrationIdentity(WorkbookEditableRow row)
    {
        const string prefix = "手打檔-";
        if (!row.ExternalOrderNo.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var rowSeparator = row.ExternalOrderNo.LastIndexOf('-');
        if (rowSeparator < prefix.Length ||
            !int.TryParse(row.ExternalOrderNo[(rowSeparator + 1)..], out var sourceRowNumber))
        {
            return null;
        }

        return string.Join(
            "\u001f",
            row.Sheet.Trim().ToUpperInvariant(),
            sourceRowNumber,
            row.ChannelName.Trim().ToUpperInvariant(),
            row.ProductName.Trim().ToUpperInvariant());
    }

    public static IReadOnlyList<OrderChangeEntry> Apply(
        IReadOnlyList<OrderChangeEntry> entries,
        IReadOnlyList<OrderRowManualOverride> overrides)
    {
        ArgumentNullException.ThrowIfNull(entries);
        ArgumentNullException.ThrowIfNull(overrides);
        var byFingerprint = new Dictionary<string, OrderRowManualOverride>(StringComparer.OrdinalIgnoreCase);
        foreach (var manual in overrides.OrderBy(item => item.UpdatedAt))
        {
            foreach (var fingerprint in manual.SourceLineFingerprints)
            {
                byFingerprint[fingerprint] = manual;
            }
        }

        return entries.Select(entry =>
        {
            if (!byFingerprint.TryGetValue(BuildEntryFingerprint(entry), out var manual))
            {
                return entry;
            }

            return entry with
            {
                OriginalDeliveryDate = manual.OriginalDeliveryDate.IsOverridden
                    ? NullIfBlank(manual.OriginalDeliveryDate.Value)
                    : entry.OriginalDeliveryDate,
                DeliveryDate = manual.OrderDate.IsOverridden
                    ? NullIfBlank(manual.OrderDate.Value)
                    : entry.DeliveryDate,
                SourceLocation = manual.Location.IsOverridden
                    ? manual.Location.Value
                    : entry.SourceLocation,
                Confirmation = manual.Confirmation.IsOverridden
                    ? manual.Confirmation.Value
                    : entry.Confirmation,
                HasManualLocation = manual.Location.IsOverridden || entry.HasManualLocation,
                HasManualConfirmation = manual.Confirmation.IsOverridden || entry.HasManualConfirmation,
                HasManualOriginalDeliveryDate = manual.OriginalDeliveryDate.IsOverridden ||
                                                entry.HasManualOriginalDeliveryDate,
                HasManualOrderDate = manual.OrderDate.IsOverridden || entry.HasManualOrderDate
            };
        }).ToArray();
    }

    public static IReadOnlyList<OrderRowManualOverride> Clear(
        IReadOnlyList<OrderRowManualOverride> overrides,
        string rowKey)
    {
        ArgumentNullException.ThrowIfNull(overrides);
        ArgumentException.ThrowIfNullOrWhiteSpace(rowKey);
        return overrides
            .Where(item => !item.RowKey.Equals(rowKey, StringComparison.OrdinalIgnoreCase))
            .ToArray();
    }

    public static IReadOnlyList<OrderChangeEntry> ApplyAutomaticResets(
        IReadOnlyList<OrderChangeEntry> entries,
        IReadOnlyList<WorkbookRowAutomaticReset> resets)
    {
        ArgumentNullException.ThrowIfNull(entries);
        ArgumentNullException.ThrowIfNull(resets);
        var fingerprints = resets
            .SelectMany(reset => reset.SourceLineFingerprints)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return entries.Select(entry =>
            fingerprints.Contains(BuildEntryFingerprint(entry))
                ? entry with
                {
                    IgnoreExistingLocation = true,
                    IgnoreExistingConfirmation = true
                }
                : entry).ToArray();
    }

    public static string BuildEntryFingerprint(OrderChangeEntry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);
        var identity = string.Join(
            "\u001f",
            entry.SourceCode.ToUpperInvariant(),
            entry.ExternalOrderNo.ToUpperInvariant(),
            entry.ExternalLineKey.ToUpperInvariant());
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)))[..24];
    }

    private static ManualFieldOverride Changed(
        string previous,
        string current,
        ManualFieldOverride? existing) =>
        !string.Equals(previous, current, StringComparison.Ordinal)
            ? new ManualFieldOverride(true, current)
            : existing ?? ManualFieldOverride.None;

    private static DateOnly ParsePendingSpecifiedDate(string value)
    {
        if (!DateOnly.TryParseExact(
                value.Trim(),
                ["yyyy-M-d", "yyyy-MM-dd", "yyyy/M/d", "yyyy/MM/dd"],
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var date))
        {
            throw new InvalidDataException("日期待確認的「指定到貨日」必須是有效的 Excel 日期固定值。");
        }

        if (date.DayOfWeek is not (DayOfWeek.Tuesday or DayOfWeek.Thursday))
        {
            throw new InvalidDataException("日期待確認的「指定到貨日」必須是星期二或星期四。");
        }

        return date;
    }

    private static bool HasAnyOverride(OrderRowManualOverride item) =>
        item.OriginalDeliveryDate.IsOverridden || item.OrderDate.IsOverridden ||
        item.Location.IsOverridden || item.Confirmation.IsOverridden;

    private static string? NullIfBlank(string value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();
}
