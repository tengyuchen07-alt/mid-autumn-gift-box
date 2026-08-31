using System.Globalization;

namespace MidAutumnGiftBox.Core;

public static class ShippingLeadTimePolicy
{
    private static readonly string[] DateFormats = ["yyyy/M/d", "yyyy/MM/dd", "yyyy-M-d", "yyyy-MM-dd"];

    public static DateOnly Adjust(DateOnly originalDate)
    {
        return AdjustDeliveryDate(originalDate).AddDays(-7);
    }

    public static DateOnly AdjustDeliveryDate(DateOnly originalDate)
    {
        var candidate = originalDate;
        while (candidate.DayOfWeek is not (DayOfWeek.Tuesday or DayOfWeek.Thursday))
        {
            candidate = candidate.AddDays(-1);
        }

        return candidate;
    }

    public static string? Resolve(
        string? derivedShippingDate,
        string? arrivalDate,
        string? shippingDateStatus)
    {
        var status = shippingDateStatus?.Trim();
        if (status is not null &&
            (status.Equals("conflict", StringComparison.OrdinalIgnoreCase) ||
             status.Equals("needs_review", StringComparison.OrdinalIgnoreCase)))
        {
            return null;
        }

        if (!string.IsNullOrWhiteSpace(derivedShippingDate))
        {
            return derivedShippingDate.Trim();
        }

        return DateOnly.TryParseExact(
            arrivalDate?.Trim(),
            DateFormats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out var originalDate)
            ? Adjust(originalDate).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : null;
    }

    public static string? ResolveOriginal(
        string? shipWindowStart,
        string? shipWindowEnd,
        string? arrivalDate,
        string? shippingDateStatus)
    {
        var sourceDate = ResolveSourceDeliveryDate(
            shipWindowStart,
            shipWindowEnd,
            arrivalDate,
            shippingDateStatus);
        return sourceDate is not null
            ? AdjustDeliveryDate(sourceDate.Value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : null;
    }

    public static string? ResolveOneDayEarlier(
        string? shipWindowStart,
        string? shipWindowEnd,
        string? arrivalDate,
        string? shippingDateStatus)
    {
        var deliveryDate = ResolveOneDayEarlierDeliveryDate(
            shipWindowStart,
            shipWindowEnd,
            arrivalDate,
            shippingDateStatus);
        return deliveryDate?.AddDays(-7).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }

    public static string? ResolveOriginalOneDayEarlier(
        string? shipWindowStart,
        string? shipWindowEnd,
        string? arrivalDate,
        string? shippingDateStatus) =>
        ResolveOneDayEarlierDeliveryDate(
                shipWindowStart,
                shipWindowEnd,
                arrivalDate,
                shippingDateStatus)
            ?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

    private static DateOnly? ResolveOneDayEarlierDeliveryDate(
        string? shipWindowStart,
        string? shipWindowEnd,
        string? arrivalDate,
        string? shippingDateStatus)
    {
        var sourceDate = ResolveSourceDeliveryDate(
            shipWindowStart,
            shipWindowEnd,
            arrivalDate,
            shippingDateStatus);
        return sourceDate is not null
            ? AdjustDeliveryDate(sourceDate.Value.AddDays(-1))
            : null;
    }

    private static DateOnly? ResolveSourceDeliveryDate(
        string? shipWindowStart,
        string? shipWindowEnd,
        string? arrivalDate,
        string? shippingDateStatus)
    {
        var status = shippingDateStatus?.Trim();
        if (status is not null &&
            (status.Equals("conflict", StringComparison.OrdinalIgnoreCase) ||
             status.Equals("needs_review", StringComparison.OrdinalIgnoreCase)))
        {
            return null;
        }

        if (TryParseDate(shipWindowStart, out var start) && TryParseDate(shipWindowEnd, out var end) && start <= end)
        {
            var candidate = start;
            while (candidate <= end && candidate.DayOfWeek != DayOfWeek.Monday)
            {
                candidate = candidate.AddDays(1);
            }

            if (candidate > end)
            {
                candidate = start;
                while (candidate <= end &&
                       candidate.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday)
                {
                    candidate = candidate.AddDays(1);
                }
            }

            return candidate <= end ? candidate : null;
        }

        return TryParseDate(arrivalDate, out var arrival) ? arrival : null;
    }

    private static bool TryParseDate(string? value, out DateOnly date) =>
        DateOnly.TryParseExact(
            value?.Trim(),
            DateFormats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out date);
}
