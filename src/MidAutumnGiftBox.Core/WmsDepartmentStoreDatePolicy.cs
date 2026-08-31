namespace MidAutumnGiftBox.Core;

public static class WmsDepartmentStoreDatePolicy
{
    private static readonly HashSet<string> Channels = new(StringComparer.Ordinal)
    {
        "正-新光A4",
        "正-台中高鐵",
        "臨新竹大全聯",
        "正-夢時代",
        "正-京站",
        "臨新莊宏匯",
        "臨桃園大江",
        "臨大葉高島屋",
        "臨新光三越台南新天地"
    };

    internal static IEnumerable<string> AllChannels => Channels;

    public static bool Applies(string sourceCode, string? channelName) =>
        IsWmsSource(sourceCode) &&
        IsDepartmentStoreChannel(channelName);

    public static bool IsDepartmentStoreChannel(string? channelName) =>
        !string.IsNullOrWhiteSpace(channelName) &&
        Channels.Contains(channelName.Trim());

    public static string? ResolveOrderDate(OrderLineSnapshot line) =>
        Applies(line.SourceCode, line.ChannelName ?? line.ChannelCode)
            ? ShippingLeadTimePolicy.ResolveOneDayEarlier(
                line.ShipWindowStart,
                line.ShipWindowEnd,
                line.ArrivalDate,
                line.ShippingDateStatus)
            : ShippingLeadTimePolicy.Resolve(
                line.DerivedShippingDate,
                line.ArrivalDate,
                line.ShippingDateStatus);

    public static string? ResolveDeliveryDate(OrderLineSnapshot line) =>
        Applies(line.SourceCode, line.ChannelName ?? line.ChannelCode)
            ? ShippingLeadTimePolicy.ResolveOriginalOneDayEarlier(
                line.ShipWindowStart,
                line.ShipWindowEnd,
                line.ArrivalDate,
                line.ShippingDateStatus)
            : ShippingLeadTimePolicy.ResolveOriginal(
                line.ShipWindowStart,
                line.ShipWindowEnd,
                line.ArrivalDate,
                line.ShippingDateStatus);

    private static bool IsWmsSource(string sourceCode) =>
        sourceCode.Equals("site1", StringComparison.OrdinalIgnoreCase) ||
        sourceCode.Equals("site2", StringComparison.OrdinalIgnoreCase);
}
