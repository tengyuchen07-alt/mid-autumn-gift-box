using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Unicode;

namespace MidAutumnGiftBox.Core;

public static class OrderSnapshotPreviewBuilder
{
    private static readonly JsonSerializerOptions CompactJsonOptions = new()
    {
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
    };

    public static PreviewResult Build(
        WmsSource source,
        IReadOnlyList<OrderLineSnapshot> snapshot,
        SyncSourceStatus? syncStatus)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(snapshot);

        var sourceLines = snapshot
            .Where(line => line.SourceCode.Equals(source.Code, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        var rows = new List<IReadOnlyDictionary<string, string?>>();
        var safeRows = new JsonArray();
        foreach (var orderGroup in sourceLines
                     .GroupBy(line => line.ExternalOrderNo, StringComparer.OrdinalIgnoreCase)
                     .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            var parentLines = orderGroup
                .Where(line => line.LineLevel.Equals("parent", StringComparison.OrdinalIgnoreCase))
                .OrderBy(line => line.ExternalLineKey, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var itemLines = orderGroup
                .Where(line => line.LineLevel.Equals("item", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (parentLines.Length == 0 && itemLines.Length > 0)
            {
                throw new InvalidDataException($"訂單 {orderGroup.Key} 的快照只有子商品，無法重建預覽。");
            }

            var representative = parentLines.FirstOrDefault() ?? orderGroup.First();
            var products = new JsonArray();
            foreach (var parent in parentLines)
            {
                var product = BuildProduct(parent);
                var children = new JsonArray();
                foreach (var item in itemLines
                             .Where(line => string.Equals(
                                 line.ParentLineKey, parent.ExternalLineKey, StringComparison.OrdinalIgnoreCase))
                             .OrderBy(line => line.ExternalLineKey, StringComparer.OrdinalIgnoreCase))
                {
                    children.Add(BuildProduct(item));
                }

                if (children.Count > 0)
                {
                    product["items"] = children;
                }

                products.Add(product);
            }

            var row = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase)
            {
                ["order_no"] = representative.ExternalOrderNo,
                ["order_date"] = representative.OrderDate,
                ["source"] = representative.ChannelName ?? ChannelDisplayName(representative.ChannelCode),
                ["shop_name"] = representative.ShopName,
                ["source_key"] = representative.ChannelCode,
                ["arrival_date"] = representative.ArrivalDate,
                ["ship_window_start"] = representative.ShipWindowStart,
                ["ship_window_end"] = representative.ShipWindowEnd,
                ["derived_shipping_date"] = representative.DerivedShippingDate,
                ["shipping_date_source"] = representative.ShippingDateSource,
                ["shipping_date_status"] = representative.ShippingDateStatus,
                ["status_code"] = representative.StatusCode,
                ["status_name"] = representative.StatusName,
                ["total_price"] = representative.TotalPrice,
                ["products"] = products.ToJsonString(CompactJsonOptions)
            };
            rows.Add(row);

            var safeRow = new JsonObject();
            foreach (var pair in row.Where(pair => !pair.Key.Equals("products", StringComparison.OrdinalIgnoreCase)))
            {
                safeRow[pair.Key] = pair.Value;
            }

            safeRow["products"] = products.DeepClone();
            safeRows.Add(safeRow);
        }

        var testedAt = syncStatus?.FinishedAt ??
                       sourceLines.Select(line => line.SynchronizedAt).DefaultIfEmpty(DateTimeOffset.Now).Max();
        var resultOk = syncStatus?.Status != "failed";
        var message = rows.Count == 0 ? "本機尚無已提交訂單快照。" : "已從本機已提交快照載入。";
        var safeEnvelope = new JsonObject
        {
            ["result"] = new JsonObject { ["ok"] = resultOk },
            ["data"] = new JsonObject { ["rows"] = safeRows }
        };

        return new PreviewResult(
            true,
            message,
            rows,
            PreviewSanitizer.Sanitize(safeEnvelope.ToJsonString(CompactJsonOptions)),
            new PreviewMetadata(
                source.Name,
                "/api_v1/order/order_query.php",
                testedAt,
                resultOk ? 200 : 0,
                syncStatus?.PageCount ?? 0,
                rows.Count,
                resultOk,
                message));
    }

    private static JsonObject BuildProduct(OrderLineSnapshot line)
    {
        var product = new JsonObject();
        SetIfPresent(product, "sku", line.Sku);
        SetIfPresent(product, "item_no", line.ItemNo);
        SetIfPresent(product, "type", line.ProductType);
        SetIfPresent(product, "name", line.ProductName);
        SetIfPresent(product, "spec", line.Spec);
        product["qty"] = line.Quantity;
        product["shipp_qty"] = line.ShippedQuantity;
        return product;
    }

    private static void SetIfPresent(JsonObject target, string name, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            target[name] = value;
        }
    }

    private static string? ChannelDisplayName(string? channelCode) =>
        channelCode?.Equals("shopee", StringComparison.OrdinalIgnoreCase) == true
            ? "蝦皮賣場"
            : channelCode;
}
