using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace MidAutumnGiftBox.Core;

public static partial class OrderPreviewEnricher
{
    private static readonly string[] ApplicationOwnedShippingFields =
    [
        "ship_window_start",
        "ship_window_end",
        "derived_shipping_date",
        "shipping_date_source",
        "shipping_date_status"
    ];

    public static void Enrich(JsonObject order, JsonArray products)
    {
        ApplyChannelDisplayName(order);
        ApplyDerivedShippingWindow(
            order,
            products,
            WmsDepartmentStoreDatePolicy.IsDepartmentStoreChannel(GetString(order, "source")));
    }

    private static void ApplyChannelDisplayName(JsonObject order)
    {
        var sourceKey = GetString(order, "source_key");
        var displayName = sourceKey?.ToLowerInvariant() switch
        {
            "shopee" => "蝦皮賣場",
            _ => null
        };
        if (displayName is null)
        {
            return;
        }

        var sourceProperty = FindKey(order, "source");
        var originalSource = sourceProperty is null ? null : GetString(order, sourceProperty);
        if (!string.IsNullOrWhiteSpace(originalSource))
        {
            SetValue(order, "shop_name", originalSource);
        }

        SetValue(order, "source", displayName);
    }

    private static void ApplyDerivedShippingWindow(
        JsonObject order,
        JsonArray products,
        bool useDepartmentStoreDatePolicy)
    {
        foreach (var field in ApplicationOwnedShippingFields)
        {
            RemoveValue(order, field);
        }

        var fallbackYear = TryReadYear(GetString(order, "order_date"));
        var parentProducts = products.OfType<JsonObject>().ToArray();
        foreach (var product in parentProducts)
        {
            ApplyDerivedShippingWindow(product, fallbackYear, useDepartmentStoreDatePolicy);
        }

        MirrorConsistentProductShippingWindow(order, parentProducts);
        if (CanApplyArrivalDateFallback(parentProducts))
        {
            ApplyArrivalDateFallback(order, useDepartmentStoreDatePolicy);
            ApplyNoteDateFallback(order, fallbackYear, useDepartmentStoreDatePolicy);
        }
    }

    private static void ApplyDerivedShippingWindow(
        JsonObject product,
        int? fallbackYear,
        bool useDepartmentStoreDatePolicy)
    {
        foreach (var field in ApplicationOwnedShippingFields)
        {
            RemoveValue(product, field);
        }

        var candidates = new List<ShippingWindow>();
        var hasInvalidExplicitWindow = false;
        var hasMultipleExplicitRanges = false;
        var name = GetString(product, "name") ?? string.Empty;
        var spec = GetString(product, "spec") ?? string.Empty;
        var year = TryReadYear(name) ?? fallbackYear;
        if (year is null)
        {
            return;
        }

        var specResult = ParseWindows(spec, year.Value, "product.spec", useDepartmentStoreDatePolicy);
        if (specResult.Status != WindowParseStatus.NotFound)
        {
            candidates.AddRange(specResult.Windows);
            hasInvalidExplicitWindow = specResult.Status == WindowParseStatus.Invalid;
            hasMultipleExplicitRanges = specResult.HasMultipleExplicitRanges;
        }
        else
        {
            var nameResult = ParseWindows(name, year.Value, "product.name", useDepartmentStoreDatePolicy);
            candidates.AddRange(nameResult.Windows);
            hasInvalidExplicitWindow = nameResult.Status == WindowParseStatus.Invalid;
            hasMultipleExplicitRanges = nameResult.HasMultipleExplicitRanges;
        }

        var distinct = candidates
            .DistinctBy(candidate => (candidate.Start, candidate.End, candidate.ShippingDate))
            .ToArray();
        if (hasMultipleExplicitRanges || distinct.Length > 1)
        {
            SetValue(product, "shipping_date_status", "conflict");
            return;
        }

        if (hasInvalidExplicitWindow)
        {
            SetValue(product, "shipping_date_status", "needs_review");
            return;
        }

        if (distinct.Length == 0)
        {
            return;
        }

        var selected = distinct[0];
        SetValue(product, "ship_window_start", selected.Start.ToString("yyyy-MM-dd"));
        SetValue(product, "ship_window_end", selected.End.ToString("yyyy-MM-dd"));
        SetValue(product, "derived_shipping_date", selected.ShippingDate.ToString("yyyy-MM-dd"));
        SetValue(product, "shipping_date_source", selected.Source);
        SetValue(product, "shipping_date_status", "derived");
    }

    private static void MirrorConsistentProductShippingWindow(
        JsonObject order,
        IReadOnlyList<JsonObject> products)
    {
        if (products.Count == 0)
        {
            return;
        }

        var signatures = products
            .Select(product => string.Join(
                "\u001f",
                ApplicationOwnedShippingFields.Select(field => GetString(product, field) ?? string.Empty)))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        var hasShippingMetadata = ApplicationOwnedShippingFields
            .Any(field => !string.IsNullOrWhiteSpace(GetString(products[0], field)));
        if (signatures.Length != 1 || !hasShippingMetadata)
        {
            return;
        }

        foreach (var field in ApplicationOwnedShippingFields)
        {
            var value = GetString(products[0], field);
            if (!string.IsNullOrWhiteSpace(value))
            {
                SetValue(order, field, value);
            }
        }
    }

    private static WindowParseResult ParseWindows(
        string value,
        int year,
        string source,
        bool useDepartmentStoreDatePolicy = false)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return new WindowParseResult(WindowParseStatus.NotFound, [], false);
        }

        var matches = MonthDayRangeRegex().Matches(value);
        if (matches.Count == 0)
        {
            return ExplicitRangeIntentRegex().IsMatch(value)
                ? new WindowParseResult(WindowParseStatus.Invalid, [], false)
                : new WindowParseResult(WindowParseStatus.NotFound, [], false);
        }

        var windows = new List<ShippingWindow>(matches.Count);
        var invalid = false;
        foreach (Match match in matches)
        {
            if (!int.TryParse(match.Groups["m1"].Value, out var startMonth) ||
                !int.TryParse(match.Groups["d1"].Value, out var startDay) ||
                !int.TryParse(match.Groups["d2"].Value, out var endDay))
            {
                invalid = true;
                continue;
            }

            var endMonth = match.Groups["m2"].Success &&
                           int.TryParse(match.Groups["m2"].Value, out var parsedEndMonth)
                ? parsedEndMonth
                : startMonth;
            DateOnly start;
            DateOnly end;
            try
            {
                start = new DateOnly(year, startMonth, startDay);
                var endYear = endMonth < startMonth ? year + 1 : year;
                end = new DateOnly(endYear, endMonth, endDay);
            }
            catch (ArgumentOutOfRangeException)
            {
                invalid = true;
                continue;
            }

            var shippingDate = start;
            while (shippingDate.DayOfWeek != DayOfWeek.Monday && shippingDate <= end)
            {
                shippingDate = shippingDate.AddDays(1);
            }

            if (shippingDate > end)
            {
                shippingDate = start;
                while (shippingDate <= end &&
                       shippingDate.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday)
                {
                    shippingDate = shippingDate.AddDays(1);
                }

                if (shippingDate > end)
                {
                    invalid = true;
                    continue;
                }
            }
            else if (shippingDate.AddDays(7) <= end)
            {
                invalid = true;
                continue;
            }

            windows.Add(new ShippingWindow(
                start,
                end,
                ShippingLeadTimePolicy.Adjust(
                    useDepartmentStoreDatePolicy ? shippingDate.AddDays(-1) : shippingDate),
                source));
        }

        return new WindowParseResult(
            invalid ? WindowParseStatus.Invalid : WindowParseStatus.Valid,
            windows,
            matches.Count > 1);
    }

    private static int? TryReadYear(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var match = YearRegex().Match(value);
        return match.Success && int.TryParse(match.Value, out var year) ? year : null;
    }

    private static void ApplyArrivalDateFallback(JsonObject order, bool useDepartmentStoreDatePolicy)
    {
        if (!string.IsNullOrWhiteSpace(GetString(order, "derived_shipping_date")))
        {
            return;
        }

        var value = GetString(order, "arrival_date");
        string[] formats = ["yyyy/M/d", "yyyy/MM/dd", "yyyy-M-d", "yyyy-MM-dd"];
        if (!DateOnly.TryParseExact(
                value,
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out var originalDate))
        {
            return;
        }

        var sourceDate = useDepartmentStoreDatePolicy ? originalDate.AddDays(-1) : originalDate;
        SetValue(order, "derived_shipping_date", ShippingLeadTimePolicy.Adjust(sourceDate).ToString("yyyy-MM-dd"));
        SetValue(order, "shipping_date_source", "order.arrival_date");
        SetValue(order, "shipping_date_status", "derived");
    }

    private static void ApplyNoteDateFallback(
        JsonObject order,
        int? fallbackYear,
        bool useDepartmentStoreDatePolicy)
    {
        if (!string.IsNullOrWhiteSpace(GetString(order, "derived_shipping_date")) ||
            fallbackYear is null)
        {
            return;
        }

        var result = ParseWindows(
            GetString(order, "note") ?? string.Empty,
            fallbackYear.Value,
            "order.note",
            useDepartmentStoreDatePolicy);
        var distinct = result.Windows
            .DistinctBy(candidate => (candidate.Start, candidate.End, candidate.ShippingDate))
            .ToArray();
        if (result.HasMultipleExplicitRanges || distinct.Length > 1)
        {
            SetValue(order, "shipping_date_status", "conflict");
            return;
        }

        if (result.Status == WindowParseStatus.Invalid)
        {
            SetValue(order, "shipping_date_status", "needs_review");
            return;
        }

        if (distinct.Length != 1)
        {
            return;
        }

        var selected = distinct[0];
        SetValue(order, "ship_window_start", selected.Start.ToString("yyyy-MM-dd"));
        SetValue(order, "ship_window_end", selected.End.ToString("yyyy-MM-dd"));
        SetValue(order, "derived_shipping_date", selected.ShippingDate.ToString("yyyy-MM-dd"));
        SetValue(order, "shipping_date_source", selected.Source);
        SetValue(order, "shipping_date_status", "derived");
    }

    private static bool CanApplyArrivalDateFallback(IReadOnlyList<JsonObject> products) =>
        products.All(product =>
            string.IsNullOrWhiteSpace(GetString(product, "derived_shipping_date")) &&
            GetString(product, "shipping_date_status") is not ("conflict" or "needs_review"));

    private static string? FindKey(JsonObject obj, string name) =>
        obj.Select(property => property.Key)
            .FirstOrDefault(key => key.Equals(name, StringComparison.OrdinalIgnoreCase));

    private static string? GetString(JsonObject obj, string name)
    {
        var key = FindKey(obj, name);
        if (key is null || obj[key] is not JsonValue value)
        {
            return null;
        }

        return value.TryGetValue<string>(out var text) ? text : value.ToJsonString().Trim('"');
    }

    private static void SetValue(JsonObject obj, string name, string value)
    {
        var key = FindKey(obj, name) ?? name;
        obj[key] = value;
    }

    private static void RemoveValue(JsonObject obj, string name)
    {
        var key = FindKey(obj, name);
        if (key is not null)
        {
            obj.Remove(key);
        }
    }

    [GeneratedRegex(@"(?<m1>\d{1,2})\s*(?:/|月)\s*(?<d1>\d{1,2})\s*日?\s*(?:-|~|～|－|至)\s*(?:(?<m2>\d{1,2})\s*(?:/|月)\s*)?(?<d2>\d{1,2})\s*日?")]
    private static partial Regex MonthDayRangeRegex();

    [GeneratedRegex(@"(?<!\d)20\d{2}(?!\d)")]
    private static partial Regex YearRegex();

    [GeneratedRegex(@"(?<!\d)\d{1,2}\s*(?:/|月)\s*\d{1,2}\s*日?\s*(?:-|~|～|－|至)")]
    private static partial Regex ExplicitRangeIntentRegex();

    private readonly record struct ShippingWindow(
        DateOnly Start,
        DateOnly End,
        DateOnly ShippingDate,
        string Source);

    private sealed record WindowParseResult(
        WindowParseStatus Status,
        IReadOnlyList<ShippingWindow> Windows,
        bool HasMultipleExplicitRanges);

    private enum WindowParseStatus
    {
        NotFound,
        Valid,
        Invalid
    }
}
