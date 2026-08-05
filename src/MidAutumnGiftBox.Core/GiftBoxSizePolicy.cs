using System.Text.RegularExpressions;

namespace MidAutumnGiftBox.Core;

public static class GiftBoxSizePolicy
{
    private static readonly IReadOnlyDictionary<string, int> SkuSizes =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["4710964232435"] = 3,
            ["4710964232411"] = 6,
            ["4710964232565"] = 9
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
}
