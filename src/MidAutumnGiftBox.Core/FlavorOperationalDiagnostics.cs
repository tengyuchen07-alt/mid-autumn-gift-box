using System.Text.Json.Nodes;

namespace MidAutumnGiftBox.Core;

public static class FlavorOperationalDiagnostics
{
    private static readonly HashSet<string> SafeValueFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "finish_time", "logistics", "logistics_code"
    };

    private static readonly string[] OperationalKeywords =
    [
        "transfer", "trans_", "shipping", "logistic", "delivery_", "preorder", "pre_order",
        "change_type", "order_type", "finish", "cancel", "return", "void", "ship_flag", "shipped"
    ];

    private static readonly string[] SensitiveKeywords =
    [
        "receiver", "recipient", "purchaser", "customer", "buyer", "contact", "phone", "mobile",
        "address", "email", "account", "token", "secret", "password", "note", "comment", "memo", "key"
    ];

    public static void Enrich(JsonObject order)
    {
        ArgumentNullException.ThrowIfNull(order);
        var fieldNames = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        var safeValues = new List<string>();

        foreach (var property in order.ToArray())
        {
            if (!IsScalar(property.Value) || !LooksOperational(property.Key))
            {
                continue;
            }

            fieldNames.Add(property.Key);
            if (SafeValueFields.Contains(property.Key))
            {
                safeValues.Add($"{property.Key}={ScalarText(property.Value)}");
            }
        }

        if (fieldNames.Count > 0)
        {
            order["operational_diagnostic_fields"] = string.Join(", ", fieldNames);
        }

        if (safeValues.Count > 0)
        {
            order["operational_diagnostics"] = string.Join("; ", safeValues);
        }
    }

    private static bool LooksOperational(string name)
    {
        if (SafeValueFields.Contains(name))
        {
            return true;
        }

        var normalized = name.ToLowerInvariant();
        return OperationalKeywords.Any(normalized.Contains) &&
               !SensitiveKeywords.Any(normalized.Contains);
    }

    private static bool IsScalar(JsonNode? value) => value is null or JsonValue;

    private static string ScalarText(JsonNode? value)
    {
        if (value is null)
        {
            return "null";
        }

        if (value is JsonValue jsonValue && jsonValue.TryGetValue<string>(out var text))
        {
            return text;
        }

        return value.ToJsonString();
    }
}
