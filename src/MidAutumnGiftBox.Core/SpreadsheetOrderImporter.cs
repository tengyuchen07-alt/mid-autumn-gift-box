using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace MidAutumnGiftBox.Core;

public sealed record ImportedSpreadsheetRow(
    int SourceRowNumber,
    string ChannelName,
    string ExternalOrderNo,
    string Sku,
    string ProductName,
    decimal? Quantity,
    string? OriginalDeliveryDate,
    string? OrderDate,
    string? SourceLocation = null,
    string? Note = null);

public sealed record ImportedSpreadsheetBatch(
    string SourceCode,
    string SourceName,
    string FileName,
    IReadOnlyList<ImportedSpreadsheetRow> Rows,
    IReadOnlyList<OrderChangeEntry> ExportEntries,
    int ExcludedRowCount,
    DateTimeOffset ImportedAt);

public static class SpreadsheetOrderImporter
{
    private static readonly HashSet<string> GiftBoxSkus =
        ["1902226", "1902248", "1902247"];
    private static readonly HashSet<string> ManualSkus =
        ["1902226", "1902248", "1902247", "1902199", "192199", "1902072", "1902042", "1902197"];
    private static readonly IReadOnlyDictionary<string, string> PosExternalLineKeys =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["1902072"] = "pos:846489166CD6145F31897D1C",
            ["1902248"] = "pos:4D030B253668ECAAE3A78025",
            ["1902247"] = "pos:A126030F8A145DB13E3A211F"
        };
    private static readonly string[] DateFormats =
        ["yyyy/M/d", "yyyy/MM/dd", "yyyy-M-d", "yyyy-MM-dd", "M/d/yyyy", "MM/dd/yyyy"];

    public static ImportedSpreadsheetBatch ReadErp(string path, DateTimeOffset importedAt) =>
        Read(
            path,
            importedAt,
            "erp",
            "ERP",
            new ImportColumns(
                Channel: "客戶簡稱",
                Date: "單據日期",
                OrderNumber: "銷貨單號",
                Sku: "品號",
                ProductName: "品名",
                Quantity: "銷貨數量",
                Note: "備註"),
            ManualSkus,
            requireOrderNumber: true,
            requireDateAndQuantity: true);

    public static ImportedSpreadsheetBatch ReadManual(string path, DateTimeOffset importedAt) =>
        Read(
            path,
            importedAt,
            "manual_excel",
            "手打單 Excel",
            new ImportColumns(
                Channel: "通路",
                Date: "抵達日",
                OrderNumber: null,
                Sku: "料號",
                ProductName: "品項",
                Quantity: "數量",
                Location: "地點",
                Note: "備註"),
            ManualSkus,
            requireOrderNumber: false,
            requireDateAndQuantity: false);

    public static ImportedSpreadsheetBatch ReadPos(
        string path,
        DateTimeOffset importedAt,
        Func<string, string, DateOnly>? missingPickupDateResolver = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException("找不到選取的 POS Excel 檔案。", fullPath);
        }

        string[] requiredHeaders =
            ["預購取貨日", "訂單來源", "訂單編號", "SKU", "貨號", "品名", "出貨數"];
        var worksheets = ReadWorksheets(fullPath);
        if (!TryFindHeaderRow(worksheets, requiredHeaders, out var worksheet, out var headerRow))
        {
            throw new InvalidDataException($"POS機 Excel 缺少必要欄位：{string.Join("、", requiredHeaders)}。");
        }

        var indexes = requiredHeaders.ToDictionary(
            header => header,
            header => headerRow.Cells.Single(cell => HeaderEquals(cell.Value.Text, header)).Key,
            StringComparer.Ordinal);
        string Text(WorksheetRow row, string header) =>
            row.Cells.TryGetValue(indexes[header], out var cell) ? cell.Text.Trim() : string.Empty;

        var rows = new List<ImportedSpreadsheetRow>();
        var entries = new List<OrderChangeEntry>();
        var excluded = 0;
        foreach (var worksheetRow in worksheet.Rows.Where(row => row.Number > headerRow.Number))
        {
            var sourceSku = NormalizeSku(Text(worksheetRow, "SKU"));
            var itemNumber = NormalizeSku(Text(worksheetRow, "貨號"));
            var productName = Text(worksheetRow, "品名");
            if (string.IsNullOrWhiteSpace(sourceSku) &&
                string.IsNullOrWhiteSpace(itemNumber) &&
                string.IsNullOrWhiteSpace(productName))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(itemNumber))
            {
                excluded++;
                continue;
            }

            if (!PosExternalLineKeys.TryGetValue(itemNumber, out var externalLineKey))
            {
                excluded++;
                continue;
            }

            var channel = Text(worksheetRow, "訂單來源");
            if (string.IsNullOrWhiteSpace(channel))
            {
                throw new InvalidDataException($"POS機 Excel 第 {worksheetRow.Number} 列的訂單來源為空白。");
            }
            var externalOrderNo = Text(worksheetRow, "訂單編號");
            if (string.IsNullOrWhiteSpace(externalOrderNo))
            {
                throw new InvalidDataException($"POS機 Excel 第 {worksheetRow.Number} 列的訂單編號為空白。");
            }
            var pickupDateText = Text(worksheetRow, "預購取貨日");
            var pickupDate = ParsePosPickupDate(pickupDateText);
            if (pickupDate is null && !string.IsNullOrWhiteSpace(pickupDateText))
            {
                throw new InvalidDataException(
                    $"POS機 Excel 第 {worksheetRow.Number} 列的預購取貨日格式不正確。");
            }
            if (pickupDate is null && missingPickupDateResolver is null)
            {
                throw new InvalidDataException(
                    $"POS機 Excel 第 {worksheetRow.Number} 列的預購取貨日為空白。");
            }
            var quantity = ParseQuantity(Text(worksheetRow, "出貨數"));
            if (quantity is null)
            {
                throw new InvalidDataException($"POS機 Excel 第 {worksheetRow.Number} 列的出貨數為空白或不是數字。");
            }

            var firstOrderDate = pickupDate is null
                ? NextTuesdayOrThursday(missingPickupDateResolver!(externalOrderNo, externalLineKey))
                : (DateOnly?)null;
            var deliveryDate = pickupDate is null
                ? firstOrderDate!.Value.AddDays(7)
                : ShippingLeadTimePolicy.AdjustDeliveryDate(pickupDate.Value.AddDays(-1));
            var orderDate = pickupDate is null
                ? firstOrderDate!.Value
                : ShippingLeadTimePolicy.Adjust(deliveryDate);
            var importedRow = new ImportedSpreadsheetRow(
                worksheetRow.Number,
                channel,
                externalOrderNo,
                string.IsNullOrWhiteSpace(itemNumber) ? sourceSku : itemNumber,
                productName,
                quantity,
                deliveryDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                orderDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                Note: pickupDate is null
                    ? "未指定預購日"
                    : TextByHeader(worksheetRow, headerRow, "系統內部備註"));
            rows.Add(importedRow);

            entries.Add(ToEntry(
                importedRow,
                "pos_excel",
                "POS機",
                externalLineKey,
                importedAt));
        }

        return new ImportedSpreadsheetBatch(
            "pos_excel",
            "POS機",
            Path.GetFileName(fullPath),
            rows,
            entries,
            excluded,
            importedAt);
    }

    private static ImportedSpreadsheetBatch Read(
        string path,
        DateTimeOffset importedAt,
        string sourceCode,
        string sourceName,
        ImportColumns columns,
        IReadOnlySet<string> acceptedSkus,
        bool requireOrderNumber,
        bool requireDateAndQuantity)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException("找不到選取的 Excel 檔案。", fullPath);
        }

        var workbooks = ReadWorksheets(fullPath);
        var requiredHeaders = new[]
            {
                columns.Channel, columns.Date, columns.OrderNumber, columns.Sku, columns.Quantity
            }
            .Where(header => header is not null)
            .Cast<string>()
            .ToArray();
        if (!TryFindHeaderRow(workbooks, requiredHeaders, out var workbook, out var headerRow))
        {
            throw new InvalidDataException($"{sourceName} Excel 缺少必要欄位：{string.Join("、", requiredHeaders)}。");
        }
        var selectedWorkbook = workbook;

        var indexes = requiredHeaders.ToDictionary(
            header => header,
            header => headerRow.Cells.Single(cell => HeaderEquals(cell.Value.Text, header)).Key,
            StringComparer.Ordinal);
        string Text(WorksheetRow row, string? header) =>
            header is not null && indexes.TryGetValue(header, out var index) && row.Cells.TryGetValue(index, out var cell)
                ? cell.Text.Trim()
                : string.Empty;

        var sourceFileIdentity = StableSourceFileIdentity(fullPath);
        var rows = new List<ImportedSpreadsheetRow>();
        var entries = new List<OrderChangeEntry>();
        var excluded = 0;
        var occurrence = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var worksheetRow in selectedWorkbook.Rows.Where(row => row.Number > headerRow.Number))
        {
            var sku = NormalizeSku(Text(worksheetRow, columns.Sku));
            if (string.IsNullOrWhiteSpace(sku))
            {
                continue;
            }
            if (!acceptedSkus.Contains(sku))
            {
                excluded++;
                continue;
            }

            var channel = Text(worksheetRow, columns.Channel);
            if (string.IsNullOrWhiteSpace(channel))
            {
                throw new InvalidDataException($"{sourceName} Excel 第 {worksheetRow.Number} 列的通路別為空白。");
            }

            var externalOrderNo = Text(worksheetRow, columns.OrderNumber);
            if (requireOrderNumber && string.IsNullOrWhiteSpace(externalOrderNo))
            {
                throw new InvalidDataException($"{sourceName} Excel 第 {worksheetRow.Number} 列的訂單編號為空白。");
            }
            if (!requireOrderNumber)
            {
                externalOrderNo = $"手打檔-{sourceFileIdentity[..10]}-{worksheetRow.Number}";
            }

            var dateText = Text(worksheetRow, columns.Date);
            var originalDate = ParseDate(dateText);
            var quantityText = Text(worksheetRow, columns.Quantity);
            var quantity = ParseQuantity(quantityText);
            if (requireDateAndQuantity && originalDate is null)
            {
                throw new InvalidDataException($"{sourceName} Excel 第 {worksheetRow.Number} 列的日期為空白或格式不正確。");
            }
            if (requireDateAndQuantity && quantity is null)
            {
                throw new InvalidDataException($"{sourceName} Excel 第 {worksheetRow.Number} 列的數量為空白或不是數字。");
            }

            DateOnly? deliveryDate = originalDate is null
                ? null
                : ShippingLeadTimePolicy.AdjustDeliveryDate(originalDate.Value);
            var normalizedOriginalDate = deliveryDate?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            var orderDate = deliveryDate is null
                ? null
                : ShippingLeadTimePolicy.Adjust(deliveryDate.Value)
                    .ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            var productName = TextByHeader(worksheetRow, headerRow, columns.ProductName);
            if (string.IsNullOrWhiteSpace(productName))
            {
                productName = ProductNameForSku(sku);
            }
            var importedRow = new ImportedSpreadsheetRow(
                worksheetRow.Number,
                channel,
                externalOrderNo,
                sku,
                productName,
                quantity,
                normalizedOriginalDate,
                orderDate,
                TextByHeader(worksheetRow, headerRow, columns.Location),
                TextByHeader(worksheetRow, headerRow, columns.Note));
            rows.Add(importedRow);
            if (quantity is null)
            {
                continue;
            }

            var occurrenceKey = $"{externalOrderNo}\u001f{sku}";
            occurrence.TryGetValue(occurrenceKey, out var count);
            count++;
            occurrence[occurrenceKey] = count;
            var externalLineKey = $"{sourceCode}:{sku}:{count}";
            entries.Add(ToEntry(
                importedRow,
                sourceCode,
                sourceName,
                externalLineKey,
                importedAt));
        }

        return new ImportedSpreadsheetBatch(
            sourceCode,
            sourceName,
            Path.GetFileName(fullPath),
            rows,
            entries,
            excluded,
            importedAt);
    }

    private static OrderChangeEntry ToEntry(
        ImportedSpreadsheetRow row,
        string sourceCode,
        string sourceName,
        string externalLineKey,
        DateTimeOffset importedAt)
    {
        var identity = string.Join("\u001f", sourceCode, row.ExternalOrderNo, externalLineKey,
            row.ChannelName, row.Sku, row.Quantity, row.OriginalDeliveryDate);
        var entryId = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)));
        var quantity = row.Quantity ?? 0m;
        return new OrderChangeEntry(
            entryId,
            importedAt,
            sourceCode,
            sourceName,
            row.ExternalOrderNo,
            externalLineKey,
            row.ChannelName,
            row.OriginalDeliveryDate,
            row.OrderDate,
            row.Sku,
            row.ProductName,
            0m,
            quantity,
            quantity,
            "檔案匯入",
            "IMPORTED",
            "已匯入",
            false,
            string.Empty,
            row.OriginalDeliveryDate,
            Note: row.Note,
            SourceLocation: row.SourceLocation);
    }

    private static string NormalizeSku(string value) =>
        decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var number) &&
        decimal.Truncate(number) == number
            ? number.ToString("0", CultureInfo.InvariantCulture)
            : value.Trim();

    private static DateOnly? ParseDate(string value)
    {
        if (DateOnly.TryParseExact(value.Trim(), DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var date))
        {
            return date;
        }

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
        {
            try
            {
                return DateOnly.FromDateTime(DateTime.FromOADate(serial));
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        return null;
    }

    private static DateOnly? ParsePosPickupDate(string value)
    {
        var normalized = value.Trim();
        var suffixIndex = normalized.IndexOfAny(['(', '（']);
        if (suffixIndex >= 0)
        {
            normalized = normalized[..suffixIndex].Trim();
        }
        return ParseDate(normalized);
    }

    private static DateOnly NextTuesdayOrThursday(DateOnly firstImportedOn)
    {
        var candidate = firstImportedOn;
        while (candidate.DayOfWeek is not (DayOfWeek.Tuesday or DayOfWeek.Thursday))
        {
            candidate = candidate.AddDays(1);
        }

        return candidate;
    }

    private static decimal? ParseQuantity(string value) =>
        decimal.TryParse(value.Trim(), NumberStyles.Number | NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture, out var quantity)
            ? quantity
            : null;

    private static bool HeaderEquals(string value, string expected) =>
        string.Equals(value.Trim(), expected, StringComparison.Ordinal);

    private static bool TryFindHeaderRow(
        IReadOnlyList<WorksheetDocument> worksheets,
        IReadOnlyCollection<string> requiredHeaders,
        out WorksheetDocument worksheet,
        out WorksheetRow headerRow)
    {
        foreach (var candidateWorksheet in worksheets)
        {
            foreach (var candidateRow in candidateWorksheet.Rows)
            {
                var remainingHeaders = new HashSet<string>(requiredHeaders, StringComparer.Ordinal);
                foreach (var cell in candidateRow.Cells.Values)
                {
                    remainingHeaders.Remove(cell.Text.Trim());
                }

                if (remainingHeaders.Count == 0)
                {
                    worksheet = candidateWorksheet;
                    headerRow = candidateRow;
                    return true;
                }
            }
        }

        worksheet = null!;
        headerRow = null!;
        return false;
    }

    private static string TextByHeader(WorksheetRow row, WorksheetRow headerRow, string? header)
    {
        if (header is null) return string.Empty;
        var headerCell = headerRow.Cells.FirstOrDefault(cell => HeaderEquals(cell.Value.Text, header));
        return headerCell.Value is not null && row.Cells.TryGetValue(headerCell.Key, out var cell)
            ? cell.Text.Trim()
            : string.Empty;
    }

    private static string ProductNameForSku(string sku) => sku switch
    {
        "1902226" => "蛋黃酥3入禮盒",
        "1902248" => "蛋黃酥6入禮盒",
        "1902247" => "蛋黃酥9入禮盒",
        "1902199" => "蛋黃酥(1入)-無盒",
        "192199" => "蛋黃酥(1入)-無盒霧面袋",
        "1902072" => "蛋黃酥(1入)組-有盒",
        "1902042" => "蛋黃酥3入條裝-無盒",
        "1902197" => "蛋黃酥6入鐵盒2024",
        _ => sku
    };

    private static string StableSourceFileIdentity(string path)
    {
        var fileName = Path.GetFileName(path).Trim().ToUpperInvariant();
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(fileName)));
    }

    private static IReadOnlyList<WorksheetDocument> ReadWorksheets(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        ValidateArchiveSafety(archive);
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var sharedStrings = ReadSharedStrings(archive, spreadsheet);
        var sheetEntries = archive.Entries
            .Where(entry => entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
                            entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (sheetEntries.Length == 0)
        {
            throw new InvalidDataException("Excel 沒有可讀取的工作表。");
        }

        return sheetEntries.Select(sheetEntry =>
        {
            EnsureSafeXmlEntry(sheetEntry);
            XDocument document;
            using (var stream = sheetEntry.Open()) document = XDocument.Load(stream);
            var rows = document.Descendants(spreadsheet + "row")
                .Select(row => new WorksheetRow(
                    int.TryParse((string?)row.Attribute("r"), out var rowNumber) ? rowNumber : 0,
                    row.Elements(spreadsheet + "c").ToDictionary(
                        cell => ColumnIndex((string?)cell.Attribute("r")),
                        cell => new WorksheetCell(ReadCellText(cell, spreadsheet, sharedStrings)))))
                .ToArray();
            return new WorksheetDocument(rows);
        }).ToArray();
    }

    private static IReadOnlyList<string> ReadSharedStrings(ZipArchive archive, XNamespace spreadsheet)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null) return [];
        EnsureSafeXmlEntry(entry);
        using var stream = entry.Open();
        return XDocument.Load(stream).Descendants(spreadsheet + "si")
            .Select(item => string.Concat(item.Descendants(spreadsheet + "t").Select(text => text.Value)))
            .ToArray();
    }

    private static string ReadCellText(
        XElement cell,
        XNamespace spreadsheet,
        IReadOnlyList<string> sharedStrings)
    {
        if ((string?)cell.Attribute("t") == "s" &&
            int.TryParse(cell.Element(spreadsheet + "v")?.Value, out var index) &&
            index >= 0 && index < sharedStrings.Count)
        {
            return sharedStrings[index];
        }

        if ((string?)cell.Attribute("t") == "inlineStr")
        {
            return string.Concat(cell.Descendants(spreadsheet + "t").Select(text => text.Value));
        }

        return cell.Element(spreadsheet + "v")?.Value ?? string.Empty;
    }

    private static int ColumnIndex(string? reference)
    {
        var result = 0;
        foreach (var character in reference?.TakeWhile(char.IsLetter) ?? [])
        {
            result = result * 26 + char.ToUpperInvariant(character) - 'A' + 1;
        }
        return result;
    }

    private static void EnsureSafeXmlEntry(ZipArchiveEntry entry)
    {
        const long maximumXmlBytes = 20 * 1024 * 1024;
        if (entry.Length > maximumXmlBytes)
        {
            throw new InvalidDataException($"Excel 工作表過大，已停止讀取：{entry.FullName}。");
        }
    }

    private static void ValidateArchiveSafety(ZipArchive archive)
    {
        const int maximumEntries = 200;
        const long maximumTotalBytes = 100 * 1024 * 1024;
        if (archive.Entries.Count > maximumEntries || archive.Entries.Sum(entry => entry.Length) > maximumTotalBytes)
        {
            throw new InvalidDataException("Excel 檔案過大或內容異常，為保護系統已停止讀取。");
        }
    }

    private sealed record ImportColumns(
        string Channel,
        string Date,
        string? OrderNumber,
        string Sku,
        string? ProductName,
        string Quantity,
        string? Location = null,
        string? Note = null);
    private sealed record WorksheetCell(string Text);
    private sealed record WorksheetRow(int Number, IReadOnlyDictionary<int, WorksheetCell> Cells);
    private sealed record WorksheetDocument(IReadOnlyList<WorksheetRow> Rows);
}
