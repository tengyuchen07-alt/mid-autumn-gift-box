using MidAutumnGiftBox.Core;

if (args.Length != 2 || args[0] is not ("erp" or "manual" or "pos"))
{
    Console.Error.WriteLine("Usage: SpreadsheetImportInspector <erp|manual|pos> <xlsx-path>");
    return 2;
}

var batch = args[0] switch
{
    "erp" => SpreadsheetOrderImporter.ReadErp(args[1], DateTimeOffset.Now),
    "manual" => SpreadsheetOrderImporter.ReadManual(args[1], DateTimeOffset.Now),
    "pos" => SpreadsheetOrderImporter.ReadPos(args[1], DateTimeOffset.Now),
    _ => throw new ArgumentOutOfRangeException()
};

Console.WriteLine($"source={batch.SourceCode} rows={batch.Rows.Count} export={batch.ExportEntries.Count} excluded={batch.ExcludedRowCount}");
foreach (var row in batch.Rows)
{
    Console.WriteLine(string.Join(" | ",
        row.SourceRowNumber,
        row.ChannelName,
        row.ExternalOrderNo,
        row.Sku,
        row.Quantity?.ToString() ?? "<blank>",
        row.OriginalDeliveryDate ?? "<blank>",
        row.OrderDate ?? "<blank>",
        row.SourceLocation ?? "<blank>"));
}

return 0;
