using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Encodings.Web;
using System.Text.Unicode;

namespace MidAutumnGiftBox.Core;

public static class PreviewSanitizer
{
    private static readonly JsonSerializerOptions CompactJsonOptions = new()
    {
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
    };

    private static readonly JsonSerializerOptions IndentedJsonOptions = new()
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
    };

    private static readonly HashSet<string> ApprovedProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "result", "ok", "rows", "data", "list", "total", "total_count",
        "maxpage", "nowpage", "pagesize",
        "shop_id", "shop_name", "shop_type", "shop_type_name",
        "order_no", "order_date", "arrival_date", "shipping_date", "source", "source_id", "source_key", "note",
        "ship_window_start", "ship_window_end", "derived_shipping_date", "shipping_date_source", "shipping_date_status",
        "operational_diagnostic_fields", "operational_diagnostics",
        "status_code", "status_name", "total_price", "products", "items",
        "sku", "item_no", "name", "spec", "price", "qty", "shipp_qty", "stock",
        "warehouse", "wh_id", "type"
    };
    public static string Sanitize(string json, Func<JsonElement, bool>? rowFilter = null)
    {
        using var document = JsonDocument.Parse(json);
        var sanitized = SanitizeElement(document.RootElement, rowFilter);
        return sanitized?.ToJsonString(IndentedJsonOptions) ?? "{}";
    }

    public static IReadOnlyList<IReadOnlyDictionary<string, string?>> ProjectRows(
        string json,
        Func<JsonElement, bool>? rowFilter = null)
    {
        using var document = JsonDocument.Parse(json);
        var rows = FindRows(document.RootElement);
        if (rows is null)
        {
            return Array.Empty<IReadOnlyDictionary<string, string?>>();
        }

        var projected = new List<IReadOnlyDictionary<string, string?>>();
        foreach (var row in rows.Value.EnumerateArray())
        {
            if (row.ValueKind != JsonValueKind.Object || (rowFilter is not null && !rowFilter(row)))
            {
                continue;
            }

            var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            foreach (var property in row.EnumerateObject())
            {
                if (!ApprovedProperties.Contains(property.Name))
                {
                    continue;
                }

                values[property.Name] = property.Value.ValueKind switch
                {
                    JsonValueKind.String when property.NameEquals("products") =>
                        SanitizeJsonText(property.Value.GetString())?.ToJsonString(CompactJsonOptions),
                    JsonValueKind.String => property.Value.GetString(),
                    JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False => property.Value.GetRawText(),
                    JsonValueKind.Null => null,
                    JsonValueKind.Array when property.NameEquals("products") =>
                        SanitizeElement(property.Value, null)?.ToJsonString(CompactJsonOptions),
                    _ => null
                };
            }

            projected.Add(values);
        }

        return projected;
    }

    public static bool TryReadResult(string json, out bool ok, out string message)
    {
        using var document = JsonDocument.Parse(json);
        var okElement = FindProperty(document.RootElement, "ok");
        var messageElement = FindProperty(document.RootElement, "message");
        message = messageElement is { ValueKind: JsonValueKind.String }
            ? messageElement.Value.GetString() ?? string.Empty
            : string.Empty;

        if (okElement is { ValueKind: JsonValueKind.True or JsonValueKind.False })
        {
            ok = okElement.Value.GetBoolean();
            return true;
        }

        ok = true;
        return false;
    }

    public static string? FindString(string json, string propertyName)
    {
        using var document = JsonDocument.Parse(json);
        var element = FindProperty(document.RootElement, propertyName);
        return element is { ValueKind: JsonValueKind.String } ? element.Value.GetString() : null;
    }

    public static int? FindInt(string json, string propertyName)
    {
        using var document = JsonDocument.Parse(json);
        var element = FindProperty(document.RootElement, propertyName);
        if (element is { ValueKind: JsonValueKind.Number } && element.Value.TryGetInt32(out var number))
        {
            return number;
        }

        if (element is { ValueKind: JsonValueKind.String } &&
            int.TryParse(element.Value.GetString(), out number))
        {
            return number;
        }

        return null;
    }

    private static JsonNode? SanitizeElement(JsonElement element, Func<JsonElement, bool>? rowFilter)
    {
        return element.ValueKind switch
        {
            JsonValueKind.Object => SanitizeObject(element, rowFilter),
            JsonValueKind.Array => SanitizeArray(element, rowFilter, false),
            JsonValueKind.String => JsonValue.Create(element.GetString()),
            JsonValueKind.Number => JsonNode.Parse(element.GetRawText()),
            JsonValueKind.True => JsonValue.Create(true),
            JsonValueKind.False => JsonValue.Create(false),
            JsonValueKind.Null => null,
            _ => null
        };
    }

    private static JsonObject SanitizeObject(JsonElement element, Func<JsonElement, bool>? rowFilter)
    {
        var result = new JsonObject();
        foreach (var property in element.EnumerateObject())
        {
            if (!ApprovedProperties.Contains(property.Name))
            {
                continue;
            }

            if (property.Value.ValueKind == JsonValueKind.String &&
                (property.NameEquals("products") || property.NameEquals("items")))
            {
                result[property.Name] = SanitizeJsonText(property.Value.GetString());
                continue;
            }

            var isRowsProperty = property.NameEquals("rows") ||
                                 property.NameEquals("data") ||
                                 property.NameEquals("list");
            result[property.Name] = property.Value.ValueKind == JsonValueKind.Array && isRowsProperty
                ? SanitizeArray(property.Value, rowFilter, true)
                : SanitizeElement(property.Value, rowFilter);
        }

        return result;
    }

    private static JsonNode? SanitizeJsonText(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
        {
            return new JsonArray();
        }

        try
        {
            using var document = JsonDocument.Parse(json);
            return SanitizeElement(document.RootElement, null);
        }
        catch (JsonException)
        {
            return new JsonArray();
        }
    }

    private static JsonArray SanitizeArray(
        JsonElement element,
        Func<JsonElement, bool>? rowFilter,
        bool applyRowFilter)
    {
        var result = new JsonArray();
        foreach (var item in element.EnumerateArray())
        {
            if (applyRowFilter && rowFilter is not null && !rowFilter(item))
            {
                continue;
            }

            result.Add(SanitizeElement(item, null));
        }

        return result;
    }

    private static JsonElement? FindRows(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if ((property.NameEquals("rows") || property.NameEquals("data") || property.NameEquals("list")) &&
                    property.Value.ValueKind == JsonValueKind.Array)
                {
                    return property.Value;
                }

                var nested = FindRows(property.Value);
                if (nested is not null)
                {
                    return nested;
                }
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                var nested = FindRows(item);
                if (nested is not null)
                {
                    return nested;
                }
            }
        }

        return null;
    }

    private static JsonElement? FindProperty(JsonElement element, string propertyName)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (property.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    return property.Value;
                }

                var nested = FindProperty(property.Value, propertyName);
                if (nested is not null)
                {
                    return nested;
                }
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                var nested = FindProperty(item, propertyName);
                if (nested is not null)
                {
                    return nested;
                }
            }
        }

        return null;
    }
}
