namespace MidAutumnGiftBox.Core;

public static class FlavorProductPolicy
{
    public static bool AppliesTo(WmsSource source) =>
        source.Code.Equals("site2", StringComparison.OrdinalIgnoreCase) ||
        source.BaseUri.Host.Equals("flavor.wms.changliu.com.tw", StringComparison.OrdinalIgnoreCase);
}
