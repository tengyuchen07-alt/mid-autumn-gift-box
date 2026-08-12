using System.Text.RegularExpressions;

namespace MidAutumnGiftBox.Core;

public static class GiftBoxSizePolicy
{
    private static readonly IReadOnlyDictionary<string, int> SkuSizes =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["4710964232435"] = 3,
            ["4710964232411"] = 6,
            ["4710964232565"] = 9,
            ["1902226"] = 3,
            ["1902248"] = 6,
            ["1902247"] = 9
        };
    private static readonly Regex PieceCountPattern = new(
        @"(?:內含\s*|每盒[^\r\n]{0,80}?蛋黃酥\s*[xX×]?\s*|蛋黃酥\s*)(3|6|9)\s*入?",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static int? Resolve(string? sku, string? productName)
    {
        if (!string.IsNullOrWhiteSpace(sku) && SkuSizes.TryGetValue(sku.Trim(), out var size))
        {
            return size;
        }

        if (string.IsNullOrWhiteSpace(productName))
        {
            return null;
        }

        var match = PieceCountPattern.Match(productName);
        return match.Success && int.TryParse(match.Groups[1].Value, out size) && size is 3 or 6 or 9
            ? size
            : null;
    }

    public static int? ResolveWithUniqueTargetChildSku(
        string? parentSku,
        string? parentProductName,
        IEnumerable<string?> childSkus)
    {
        var parentSize = Resolve(parentSku, parentProductName);
        if (parentSize is not null)
        {
            return parentSize;
        }

        ArgumentNullException.ThrowIfNull(childSkus);
        var targetChildSkus = childSkus
            .Where(sku => !string.IsNullOrWhiteSpace(sku))
            .Select(sku => sku!.Trim())
            .Where(SkuSizes.ContainsKey)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return targetChildSkus.Length == 1
            ? SkuSizes[targetChildSkus[0]]
            : null;
    }

    public static int? ResolveParent(
        OrderLineSnapshot parent,
        IEnumerable<OrderLineSnapshot> sourceSnapshot)
    {
        ArgumentNullException.ThrowIfNull(parent);
        ArgumentNullException.ThrowIfNull(sourceSnapshot);
        return ResolveWithUniqueTargetChildSku(
            parent.Sku,
            parent.ProductName,
            sourceSnapshot
                .Where(line =>
                    line.LineLevel.Equals("item", StringComparison.OrdinalIgnoreCase) &&
                    line.SourceCode.Equals(parent.SourceCode, StringComparison.OrdinalIgnoreCase) &&
                    line.ExternalOrderNo.Equals(parent.ExternalOrderNo, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(line.ParentLineKey, parent.ExternalLineKey, StringComparison.OrdinalIgnoreCase))
                .Select(line => line.Sku));
    }
}
