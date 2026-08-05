using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using System.Globalization;

namespace MidAutumnGiftBox.Core;

public static class PreviewWorkbookExporter
{
    private static readonly string[] PreferredOrderColumns =
    [
        "order_no", "order_date", "source", "shop_name", "source_key", "arrival_date",
        "ship_window_start", "ship_window_end", "derived_shipping_date", "shipping_date_source",
        "status_code", "status_name", "total_price"
    ];

    private static readonly string[] ProductColumns =
    [
        "通路別", "賣場名稱", "訂單編號", "訂單成立日", "指定到貨日",
        "出貨區間起日", "出貨區間迄日", "推導出貨日",
        "禮盒SKU", "禮盒品號", "禮盒類型", "禮盒品名", "禮盒規格", "禮盒數量", "禮盒已出貨數量",
        "主商品SKU", "主商品品號", "主商品類型", "主商品品名", "主商品規格", "主商品數量", "主商品已出貨數量"
    ];

    private static readonly string[] ChangeColumns =
    [
        "系統異動ID", "同步時間", "網站來源", "通路別", "訂單編號", "訂單成立日", "指定出貨日",
        "SKU", "商品名稱", "異動前數量", "異動後數量", "異動數量", "異動類型", "API狀態",
        "需人工確認", "確認"
    ];

    public static void Export(string path, PreviewResult preview)
    {
        ExportNewWorkbook(path, preview, Array.Empty<OrderChangeEntry>(), includeChangeSheet: false);
    }

    public static void Export(
        string path,
        PreviewResult preview,
        IReadOnlyList<OrderChangeEntry> changeEntries)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(preview);
        ArgumentNullException.ThrowIfNull(changeEntries);

        var fullPath = Path.GetFullPath(path);
        if (File.Exists(fullPath))
        {
            UpdateExistingWorkbook(fullPath, preview, changeEntries);
            return;
        }

        ExportNewWorkbook(fullPath, preview, changeEntries, includeChangeSheet: true);
    }

    private static void ExportNewWorkbook(
        string path,
        PreviewResult preview,
        IReadOnlyList<OrderChangeEntry> changeEntries,
        bool includeChangeSheet)
    {
        var fullPath = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var orderTable = BuildOrderTable(preview.Rows);
        var productTable = BuildProductTable(preview.Rows);

        using var stream = new FileStream(fullPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false, Encoding.UTF8);
        WriteTextEntry(archive, "[Content_Types].xml", includeChangeSheet ? ContentTypesWithChangesXml : ContentTypesXml);
        WriteTextEntry(archive, "_rels/.rels", PackageRelationshipsXml);
        WriteTextEntry(archive, "xl/workbook.xml", includeChangeSheet ? WorkbookWithChangesXml : WorkbookXml);
        WriteTextEntry(archive, "xl/_rels/workbook.xml.rels",
            includeChangeSheet ? WorkbookRelationshipsWithChangesXml : WorkbookRelationshipsXml);
        WriteWorksheet(archive, "xl/worksheets/sheet1.xml", orderTable.Headers, orderTable.Rows);
        WriteWorksheet(archive, "xl/worksheets/sheet2.xml", ProductColumns, productTable);
        if (includeChangeSheet)
        {
            WriteWorksheet(archive, "xl/worksheets/sheet3.xml", ChangeColumns, BuildChangeRows(changeEntries));
        }
    }

    private static (IReadOnlyList<string> Headers, IReadOnlyList<IReadOnlyList<string?>> Rows) BuildOrderTable(
        IReadOnlyList<IReadOnlyDictionary<string, string?>> rows)
    {
        var discovered = rows.SelectMany(row => row.Keys)
            .Where(key => !key.Equals("products", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var headers = PreferredOrderColumns.Where(preferred =>
                discovered.Contains(preferred, StringComparer.OrdinalIgnoreCase))
            .Concat(discovered.Where(column =>
                !PreferredOrderColumns.Contains(column, StringComparer.OrdinalIgnoreCase)))
            .Append("商品摘要")
            .ToArray();

        var values = new List<IReadOnlyList<string?>>(rows.Count);
        foreach (var row in rows)
        {
            var line = new List<string?>(headers.Length);
            foreach (var header in headers)
            {
                line.Add(header == "商品摘要"
                    ? BuildProductSummary(GetValue(row, "products"))
                    : GetValue(row, header));
            }

            values.Add(line);
        }

        return (headers, values);
    }

    private static IReadOnlyList<IReadOnlyList<string?>> BuildChangeRows(
        IReadOnlyList<OrderChangeEntry> entries) =>
        entries.Select(BuildChangeRow).ToArray();

    private static IReadOnlyList<string?> BuildChangeRow(OrderChangeEntry entry) =>
    [
        entry.EntryId,
        entry.ObservedAt.ToLocalTime().ToString("yyyy/MM/dd HH:mm:ss"),
        entry.SourceName,
        entry.ChannelName,
        entry.ExternalOrderNo,
        entry.OrderDate,
        entry.DeliveryDate,
        entry.Sku,
        entry.ProductName,
        entry.PreviousQuantity.ToString(CultureInfo.InvariantCulture),
        entry.NewQuantity.ToString(CultureInfo.InvariantCulture),
        entry.QuantityChange.ToString(CultureInfo.InvariantCulture),
        entry.ChangeType,
        string.Join(" ", new[] { entry.StatusCode, entry.StatusName }
            .Where(value => !string.IsNullOrWhiteSpace(value))),
        entry.NeedsReview ? "是" : "否",
        string.Empty
    ];

    private static void UpdateExistingWorkbook(
        string fullPath,
        PreviewResult preview,
        IReadOnlyList<OrderChangeEntry> changeEntries)
    {
        var temporaryPath = fullPath + $".{Guid.NewGuid():N}.tmp";
        try
        {
            File.Copy(fullPath, temporaryPath, overwrite: true);
            using (var archive = ZipFile.Open(temporaryPath, ZipArchiveMode.Update))
            {
                var changeSheet = archive.GetEntry("xl/worksheets/sheet3.xml")
                    ?? throw new InvalidDataException(
                        "選取的 Excel 不是含有「異動紀錄」的中秋禮盒活頁簿，請另存新檔。");
                var orderTable = BuildOrderTable(preview.Rows);
                var productTable = BuildProductTable(preview.Rows);
                ReplaceWorksheet(archive, "xl/worksheets/sheet1.xml", orderTable.Headers, orderTable.Rows);
                ReplaceWorksheet(archive, "xl/worksheets/sheet2.xml", ProductColumns, productTable);
                AppendChangeRows(archive, changeSheet, changeEntries);
            }

            File.Move(temporaryPath, fullPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }
        }
    }

    private static void ReplaceWorksheet(
        ZipArchive archive,
        string path,
        IReadOnlyList<string> headers,
        IReadOnlyList<IReadOnlyList<string?>> rows)
    {
        archive.GetEntry(path)?.Delete();
        WriteWorksheet(archive, path, headers, rows);
    }

    private static void AppendChangeRows(
        ZipArchive archive,
        ZipArchiveEntry changeSheet,
        IReadOnlyList<OrderChangeEntry> changeEntries)
    {
        XDocument document;
        using (var input = changeSheet.Open())
        {
            document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        }

        XNamespace spreadsheet = SpreadsheetNamespace;
        var sheetData = document.Root?.Element(spreadsheet + "sheetData")
            ?? throw new InvalidDataException("Excel 的異動紀錄工作表格式不完整。");
        var sharedStrings = ReadSharedStrings(archive, spreadsheet);
        var knownIds = sheetData.Elements(spreadsheet + "row")
            .Skip(1)
            .Select(row => row.Elements(spreadsheet + "c")
                .FirstOrDefault(cell => ((string?)cell.Attribute("r"))?.StartsWith("A", StringComparison.OrdinalIgnoreCase) == true))
            .Where(cell => cell is not null)
            .Select(cell => ReadCellText(cell!, spreadsheet, sharedStrings))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var nextRowNumber = sheetData.Elements(spreadsheet + "row")
            .Select(row => int.TryParse((string?)row.Attribute("r"), out var number) ? number : 0)
            .DefaultIfEmpty(1)
            .Max() + 1;

        foreach (var entry in changeEntries)
        {
            if (!knownIds.Add(entry.EntryId))
            {
                continue;
            }

            sheetData.Add(CreateRowElement(nextRowNumber++, BuildChangeRow(entry), spreadsheet));
        }

        var autoFilter = document.Root?.Element(spreadsheet + "autoFilter");
        autoFilter?.SetAttributeValue("ref", $"A1:{ColumnName(ChangeColumns.Length)}{Math.Max(1, nextRowNumber - 1)}");
        changeSheet.Delete();
        var replacement = archive.CreateEntry("xl/worksheets/sheet3.xml", CompressionLevel.Optimal);
        using var output = replacement.Open();
        document.Save(output, SaveOptions.DisableFormatting);
    }

    private static IReadOnlyList<string> ReadSharedStrings(ZipArchive archive, XNamespace spreadsheet)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null)
        {
            return Array.Empty<string>();
        }

        using var input = entry.Open();
        var document = XDocument.Load(input);
        return document.Descendants(spreadsheet + "si")
            .Select(item => string.Concat(item.Descendants(spreadsheet + "t").Select(text => text.Value)))
            .ToArray();
    }

    private static string ReadCellText(
        XElement cell,
        XNamespace spreadsheet,
        IReadOnlyList<string> sharedStrings)
    {
        if ((string?)cell.Attribute("t") == "s" &&
            int.TryParse(cell.Element(spreadsheet + "v")?.Value, out var sharedIndex) &&
            sharedIndex >= 0 && sharedIndex < sharedStrings.Count)
        {
            return sharedStrings[sharedIndex];
        }

        return string.Concat(cell.Descendants(spreadsheet + "t").Select(text => text.Value));
    }

    private static XElement CreateRowElement(
        int rowNumber,
        IReadOnlyList<string?> values,
        XNamespace spreadsheet)
    {
        return new XElement(spreadsheet + "row",
            new XAttribute("r", rowNumber),
            values.Select((value, index) =>
                new XElement(spreadsheet + "c",
                    new XAttribute("r", $"{ColumnName(index + 1)}{rowNumber}"),
                    new XAttribute("t", "inlineStr"),
                    new XElement(spreadsheet + "is",
                        new XElement(spreadsheet + "t",
                            new XAttribute(XNamespace.Xml + "space", "preserve"),
                            RemoveInvalidXmlCharacters(value ?? string.Empty))))));
    }

    private static IReadOnlyList<IReadOnlyList<string?>> BuildProductTable(
        IReadOnlyList<IReadOnlyDictionary<string, string?>> rows)
    {
        var output = new List<IReadOnlyList<string?>>();
        foreach (var order in rows)
        {
            foreach (var product in ParseProducts(GetValue(order, "products")))
            {
                var children = ReadArray(product, "items").ToArray();
                if (children.Length == 0)
                {
                    output.Add(BuildProductLine(order, product, null));
                    continue;
                }

                foreach (var child in children)
                {
                    output.Add(BuildProductLine(order, product, child));
                }
            }
        }

        return output;
    }

    private static IReadOnlyList<string?> BuildProductLine(
        IReadOnlyDictionary<string, string?> order,
        JsonElement product,
        JsonElement? child)
    {
        return
        [
            GetValue(order, "source"),
            GetValue(order, "shop_name"),
            GetValue(order, "order_no"),
            GetValue(order, "order_date"),
            GetValue(order, "arrival_date"),
            GetValue(order, "ship_window_start"),
            GetValue(order, "ship_window_end"),
            GetValue(order, "derived_shipping_date"),
            ReadScalar(product, "sku"),
            ReadScalar(product, "item_no"),
            ReadScalar(product, "type"),
            ReadScalar(product, "name"),
            ReadScalar(product, "spec"),
            ReadScalar(product, "qty"),
            ReadScalar(product, "shipp_qty"),
            child is null ? null : ReadScalar(child.Value, "sku"),
            child is null ? null : ReadScalar(child.Value, "item_no"),
            child is null ? null : ReadScalar(child.Value, "type"),
            child is null ? null : ReadScalar(child.Value, "name"),
            child is null ? null : ReadScalar(child.Value, "spec"),
            child is null ? null : ReadScalar(child.Value, "qty"),
            child is null ? null : ReadScalar(child.Value, "shipp_qty")
        ];
    }

    private static string BuildProductSummary(string? productsJson)
    {
        return string.Join("；", ParseProducts(productsJson).Select(product =>
        {
            var sku = ReadScalar(product, "sku");
            var name = ReadScalar(product, "name");
            var quantity = ReadScalar(product, "qty");
            var children = ReadArray(product, "items")
                .Select(child => ReadScalar(child, "sku"))
                .Where(value => !string.IsNullOrWhiteSpace(value));
            var childText = string.Join("、", children);
            return $"{sku} {name} × {quantity}" +
                   (string.IsNullOrWhiteSpace(childText) ? string.Empty : $"（內含：{childText}）");
        }));
    }

    private static IEnumerable<JsonElement> ParseProducts(string? productsJson)
    {
        if (string.IsNullOrWhiteSpace(productsJson))
        {
            yield break;
        }

        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(productsJson);
        }
        catch (JsonException)
        {
            yield break;
        }

        using (document)
        {
            var root = document.RootElement;
            if (root.ValueKind == JsonValueKind.String)
            {
                foreach (var nested in ParseProducts(root.GetString()))
                {
                    yield return nested.Clone();
                }

                yield break;
            }

            if (root.ValueKind != JsonValueKind.Array)
            {
                yield break;
            }

            foreach (var product in root.EnumerateArray())
            {
                if (product.ValueKind == JsonValueKind.Object)
                {
                    yield return product.Clone();
                }
            }
        }
    }

    private static IEnumerable<JsonElement> ReadArray(JsonElement element, string propertyName)
    {
        if (!TryGetProperty(element, propertyName, out var value))
        {
            yield break;
        }

        if (value.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in value.EnumerateArray())
            {
                if (item.ValueKind == JsonValueKind.Object)
                {
                    yield return item;
                }
            }

            yield break;
        }

        if (value.ValueKind == JsonValueKind.String)
        {
            foreach (var item in ParseProducts(value.GetString()))
            {
                yield return item;
            }
        }
    }

    private static string? ReadScalar(JsonElement element, string propertyName)
    {
        if (!TryGetProperty(element, propertyName, out var value))
        {
            return null;
        }

        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False => value.GetRawText(),
            _ => null
        };
    }

    private static bool TryGetProperty(JsonElement element, string propertyName, out JsonElement value)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (property.Name.Equals(propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }

    private static string? GetValue(IReadOnlyDictionary<string, string?> row, string key)
    {
        if (row.TryGetValue(key, out var value))
        {
            return value;
        }

        var match = row.FirstOrDefault(pair => pair.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
        return match.Equals(default(KeyValuePair<string, string?>)) ? null : match.Value;
    }

    private static void WriteWorksheet(
        ZipArchive archive,
        string path,
        IReadOnlyList<string> headers,
        IReadOnlyList<IReadOnlyList<string?>> rows)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, XmlSettings);
        writer.WriteStartDocument(true);
        writer.WriteStartElement("worksheet", SpreadsheetNamespace);
        writer.WriteStartElement("sheetData", SpreadsheetNamespace);
        WriteRow(writer, 1, headers.Cast<string?>().ToArray());
        for (var index = 0; index < rows.Count; index++)
        {
            WriteRow(writer, index + 2, rows[index]);
        }

        writer.WriteEndElement();
        if (headers.Count > 0)
        {
            writer.WriteStartElement("autoFilter", SpreadsheetNamespace);
            writer.WriteAttributeString("ref", $"A1:{ColumnName(headers.Count)}{Math.Max(1, rows.Count + 1)}");
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void WriteRow(XmlWriter writer, int rowNumber, IReadOnlyList<string?> values)
    {
        writer.WriteStartElement("row", SpreadsheetNamespace);
        writer.WriteAttributeString("r", rowNumber.ToString());
        for (var index = 0; index < values.Count; index++)
        {
            writer.WriteStartElement("c", SpreadsheetNamespace);
            writer.WriteAttributeString("r", $"{ColumnName(index + 1)}{rowNumber}");
            writer.WriteAttributeString("t", "inlineStr");
            writer.WriteStartElement("is", SpreadsheetNamespace);
            writer.WriteStartElement("t", SpreadsheetNamespace);
            writer.WriteAttributeString("xml", "space", null, "preserve");
            writer.WriteString(RemoveInvalidXmlCharacters(values[index] ?? string.Empty));
            writer.WriteEndElement();
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
    }

    private static string RemoveInvalidXmlCharacters(string value)
    {
        StringBuilder? cleaned = null;
        var offset = 0;
        foreach (var rune in value.EnumerateRunes())
        {
            var codePoint = rune.Value;
            var isValid = codePoint is 0x9 or 0xA or 0xD ||
                          codePoint is >= 0x20 and <= 0xD7FF ||
                          codePoint is >= 0xE000 and <= 0xFFFD ||
                          codePoint is >= 0x10000 and <= 0x10FFFF;
            if (isValid)
            {
                if (cleaned is not null)
                {
                    cleaned.Append(rune.ToString());
                }
            }
            else
            {
                cleaned ??= new StringBuilder(value.Length).Append(value.AsSpan(0, offset));
            }

            offset += rune.Utf16SequenceLength;
        }

        return cleaned?.ToString() ?? value;
    }

    private static string ColumnName(int column)
    {
        var result = string.Empty;
        while (column > 0)
        {
            column--;
            result = (char)('A' + column % 26) + result;
            column /= 26;
        }

        return result;
    }

    private static void WriteTextEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XmlWriterSettings XmlSettings = new()
    {
        Encoding = new UTF8Encoding(false),
        Indent = false,
        CloseOutput = false
    };

    private const string ContentTypesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
        """;

    private const string ContentTypesWithChangesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
        """;

    private const string PackageRelationshipsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """;

    private const string WorkbookXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="訂單預覽" sheetId="1" r:id="rId1"/>
            <sheet name="商品明細" sheetId="2" r:id="rId2"/>
          </sheets>
        </workbook>
        """;

    private const string WorkbookWithChangesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="訂單預覽" sheetId="1" r:id="rId1"/>
            <sheet name="商品明細" sheetId="2" r:id="rId2"/>
            <sheet name="異動紀錄" sheetId="3" r:id="rId3"/>
          </sheets>
        </workbook>
        """;

    private const string WorkbookRelationshipsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
        </Relationships>
        """;

    private const string WorkbookRelationshipsWithChangesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
        </Relationships>
        """;
}
