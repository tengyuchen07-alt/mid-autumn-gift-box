namespace MidAutumnGiftBox.Core;

public enum GiftBoxItemRole
{
    DependentSingle,
    GiftDateAnchor,
    Independent
}

public sealed record GiftBoxItemDefinition(
    string ItemNo,
    string DisplayName,
    int PieceCount,
    GiftBoxItemRole Role);

public static class GiftBoxItemPolicy
{
    private static readonly IReadOnlyDictionary<string, GiftBoxItemDefinition> Items =
        new Dictionary<string, GiftBoxItemDefinition>(StringComparer.Ordinal)
        {
            ["1902199"] = new("1902199", "單顆(無盒)(透明袋)", 1, GiftBoxItemRole.DependentSingle),
            ["192199"] = new("192199", "單顆(無盒)(霧面袋)", 1, GiftBoxItemRole.DependentSingle),
            ["1902072"] = new("1902072", "單顆(有盒)(霧面袋)", 1, GiftBoxItemRole.DependentSingle),
            ["1902042"] = new("1902042", "3入條裝(無盒)(霧面袋)", 3, GiftBoxItemRole.Independent),
            ["1902226"] = new("1902226", "3入禮盒紙盒(黑金袋)", 3, GiftBoxItemRole.GiftDateAnchor),
            ["1902197"] = new("1902197", "6入鐵盒2024年鐵盒(黑金袋)", 6, GiftBoxItemRole.GiftDateAnchor),
            ["1902248"] = new("1902248", "6入鐵盒2026年鐵盒(黑金袋)", 6, GiftBoxItemRole.GiftDateAnchor),
            ["1902247"] = new("1902247", "9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)", 9, GiftBoxItemRole.GiftDateAnchor)
        };

    public static bool TryResolve(string? itemNo, out GiftBoxItemDefinition definition)
    {
        var normalized = itemNo?.Trim() ?? string.Empty;
        return Items.TryGetValue(normalized, out definition!);
    }
}
