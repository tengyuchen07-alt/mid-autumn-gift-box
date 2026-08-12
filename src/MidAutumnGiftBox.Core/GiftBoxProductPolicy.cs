using System.Text.Json;
using System.Text.Json.Nodes;

namespace MidAutumnGiftBox.Core;

public static class GiftBoxProductPolicy
{
    private const string EggYolkPastryText = "蛋黃酥";

    public static IReadOnlySet<string> TargetSkus { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "4710964232411",
        "4710964232565",
        "4710964232435"
    };

    public static bool ContainsRelevantProduct(JsonElement order)
    {
        if (order.ValueKind != JsonValueKind.Object ||
            !TryGetProperty(order, "products", out var products))
        {
            return false;
        }

        return EnumerateArray(products).Any(IsRelevantProduct);
    }

    public static bool IsRelevantProduct(JsonElement product)
    {
        if (product.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        var skuMatches = TryGetProperty(product, "sku", out var sku) &&
                         sku.ValueKind == JsonValueKind.String &&
                         TargetSkus.Contains(sku.GetString() ?? string.Empty);
        var itemNoMatches = TryGetProperty(product, "item_no", out var itemNo) &&
                            itemNo.ValueKind == JsonValueKind.String &&
                            GiftBoxItemPolicy.TryResolve(itemNo.GetString(), out _);
        var nameMatches = TryGetProperty(product, "name", out var name) &&
                          name.ValueKind == JsonValueKind.String &&
                          (name.GetString() ?? string.Empty).Contains(
                              EggYolkPastryText,
                              StringComparison.OrdinalIgnoreCase);
        if (itemNoMatches || skuMatches || nameMatches)
        {
            return true;
        }

        return TryGetProperty(product, "items", out var items) &&
               EnumerateArray(items).Any(IsRelevantProduct);
    }

    public static bool IsRelevantProduct(JsonObject product) =>
        IsRelevantProduct(JsonSerializer.SerializeToElement(product));

    private static IEnumerable<JsonElement> EnumerateArray(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in value.EnumerateArray())
            {
                yield return item;
            }

            yield break;
        }

        if (value.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(value.GetString()))
        {
            yield break;
        }

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(value.GetString()!);
        }
        catch (JsonException)
        {
            yield break;
        }

        using (document)
        {
            if (document.RootElement.ValueKind != JsonValueKind.Array)
            {
                yield break;
            }

            foreach (var item in document.RootElement.EnumerateArray())
            {
                yield return item.Clone();
            }
        }
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
