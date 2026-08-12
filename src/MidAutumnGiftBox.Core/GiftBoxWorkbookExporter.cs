using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace MidAutumnGiftBox.Core;

public static class GiftBoxWorkbookExporter
{
    private static readonly IReadOnlyDictionary<string, string> WebsiteChannelLocations =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["蝦皮賣場"] = "創鵬",
            ["電子商務-二聯客戶"] = "彰廠",
            ["正-新光A4"] = "彰廠",
            ["正-台中高鐵"] = "彰廠",
            ["臨新竹大全聯"] = "彰廠",
            ["正-夢時代"] = "彰廠",
            ["正-京站"] = "彰廠",
            ["MOMO"] = "創鵬",
            ["好的文創"] = "彰廠",
            ["美安"] = "創鵬",
            ["YAHOO"] = "彰廠",
            ["阿瘦"] = "彰廠",
            ["臨新莊宏匯"] = "彰廠",
            ["臨桃園大江"] = "彰廠",
            ["臨大葉高島屋"] = "彰廠",
            ["臨新光三越台南新天地"] = "彰廠"
        };
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly string[] DetailHeaders =
        ["通路別", "品名", "指定到貨日", "下單日", "數量", "確認", "地點", "訂單編號", "備註", "訂單鍵", "來源明細指紋"];
    private static readonly string[] PendingDateHeaders =
        ["通路別", "品名", "指定到貨日", "數量", "確認", "地點", "訂單編號", "備註", "訂單鍵", "來源明細指紋"];
    private static readonly string[] LegacyPendingDateHeaders =
        ["通路別", "品名", "數量", "確認", "地點", "訂單編號", "備註", "訂單鍵", "來源明細指紋"];
    private static readonly string[] StatisticsHeaders =
    [
        "下單日", "到貨日", "地點",
        "單顆(無盒)(透明袋)",
        "單顆(無盒)(霧面袋)",
        "單顆(有盒)(霧面袋)",
        "3入條裝(無盒)(霧面袋)",
        "3入禮盒紙盒(黑金袋)",
        "6入鐵盒2024年鐵盒(黑金袋)",
        "6入鐵盒2026年鐵盒(黑金袋)",
        "9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)"
    ];
    private static readonly string[][] LegacyStatisticsHeaders =
    [
        ["下單日", "地點", "一入", "三入", "六入", "九入", "合計"],
        ["下單日", "一入", "三入", "六入", "九入"]
    ];
    private const string DetailSheetName = "訂單明細";
    private const string DetailSheetPath = "xl/worksheets/sheet1.xml";
    private const string PendingDateSheetName = "日期待確認";
    private const string PendingDateSheetPath = "xl/worksheets/sheet2.xml";
    private const string StatisticsSheetName = "統計";
    private const string StatisticsSheetPath = "xl/worksheets/sheet3.xml";
    private static readonly XmlWriterSettings XmlSettings = new()
    {
        Encoding = new UTF8Encoding(false),
        Indent = false,
        CloseOutput = false
    };

    public static void Export(string path, IReadOnlyList<OrderChangeEntry> entries)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(entries);
        var fullPath = Path.GetFullPath(path);
        using var workbookLock = AcquireWorkbookLock(fullPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (File.Exists(fullPath))
        {
            AppendExisting(fullPath, entries);
            return;
        }

        using var stream = new FileStream(fullPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: false, Encoding.UTF8);
        WriteTextEntry(archive, "[Content_Types].xml", ContentTypesXml);
        WriteTextEntry(archive, "_rels/.rels", PackageRelationshipsXml);
        WriteTextEntry(archive, "xl/workbook.xml", WorkbookXml);
        WriteTextEntry(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationshipsXml);
        WriteTextEntry(archive, "xl/styles.xml", StylesXml);
        var details = DetailAggregates(entries);
        WriteDetailWorksheet(archive, details, []);
        WritePendingDateWorksheet(archive, PendingDateAggregates(entries), []);
        WriteStatisticsWorksheet(archive, details);
    }

    public static void Export(
        string path,
        IReadOnlyList<OrderChangeEntry> entries,
        IReadOnlyList<OrderLineSnapshot> snapshots)
    {
        ArgumentNullException.ThrowIfNull(entries);
        ArgumentNullException.ThrowIfNull(snapshots);
        Export(path, ResolveSnapshotDates(entries, snapshots));
    }

    private static IReadOnlyList<OrderChangeEntry> ResolveSnapshotDates(
        IReadOnlyList<OrderChangeEntry> entries,
        IReadOnlyList<OrderLineSnapshot> snapshots)
    {
        var snapshotDates = snapshots
            .Where(snapshot => snapshot.LineLevel.Equals("parent", StringComparison.OrdinalIgnoreCase))
            .GroupBy(SnapshotLineKey, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.Key,
                group =>
                {
                    var snapshot = group.OrderByDescending(item => item.SynchronizedAt).First();
                    return new SnapshotDateResolution(
                        ShippingLeadTimePolicy.ResolveOriginal(
                            snapshot.ShipWindowStart,
                            snapshot.ShipWindowEnd,
                            snapshot.ArrivalDate,
                            snapshot.ShippingDateStatus),
                        ShippingLeadTimePolicy.Resolve(
                            snapshot.DerivedShippingDate,
                            snapshot.ArrivalDate,
                            snapshot.ShippingDateStatus),
                        IsDateBlocked(snapshot.ShippingDateStatus),
                        ResolveSnapshotGiftBoxSize(snapshot, snapshots),
                        snapshot.Note);
                },
                StringComparer.OrdinalIgnoreCase);
        return entries
            .Select(entry =>
            {
                var normalizedEntry = string.IsNullOrWhiteSpace(entry.OriginalDeliveryDate) &&
                                      !string.IsNullOrWhiteSpace(entry.DeliveryDate)
                    ? entry with { OriginalDeliveryDate = entry.DeliveryDate }
                    : entry;
                if (!snapshotDates.TryGetValue(EntryLineKey(entry), out var resolution))
                {
                    return normalizedEntry;
                }

                var enrichedEntry = normalizedEntry with
                {
                    GiftBoxSize = normalizedEntry.GiftBoxSize ?? resolution.GiftBoxSize,
                    Note = resolution.Note
                };

                if (resolution.BlocksDate)
                {
                    return enrichedEntry with
                    {
                        DeliveryDate = null,
                        OriginalDeliveryDate = null
                    };
                }

                return enrichedEntry with
                {
                    DeliveryDate = resolution.DeliveryDate ?? enrichedEntry.DeliveryDate,
                    OriginalDeliveryDate = resolution.OriginalDeliveryDate ??
                                           resolution.DeliveryDate ??
                                           enrichedEntry.OriginalDeliveryDate
                };
            })
            .ToArray();
    }

    public static IReadOnlyList<OrderChangeEntry> ResolveEntries(
        IReadOnlyList<OrderChangeEntry> entries,
        IReadOnlyList<OrderLineSnapshot> snapshots)
    {
        ArgumentNullException.ThrowIfNull(entries);
        ArgumentNullException.ThrowIfNull(snapshots);
        return ResolveSnapshotDates(entries, snapshots);
    }

    public static IReadOnlyList<WorkbookEditableRow> ReadEditableRows(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException("找不到先前正式匯出的 Excel。", fullPath);
        }

        using var archive = ZipFile.OpenRead(fullPath);
        ValidateArchiveSafety(archive);
        ValidateWorkbook(archive);
        XNamespace spreadsheet = SpreadsheetNamespace;
        var sharedStrings = ReadSharedStrings(archive, spreadsheet);
        var rows = new List<WorkbookEditableRow>();
        ReadRows(DetailSheetPath, "detail", "K", "J", "C", "D", "G", "F", "A", "H", "B");
        if (HasLegacyPendingHeaders(archive, spreadsheet, sharedStrings))
        {
            ReadRows(PendingDateSheetPath, "pending", "I", "H", null, null, "E", "D", "A", "F", "B");
        }
        else
        {
            ReadRows(PendingDateSheetPath, "pending", "J", "I", "C", null, "F", "E", "A", "G", "B");
        }
        return rows;

        void ReadRows(
            string sheetPath,
            string sheetKind,
            string fingerprintColumn,
            string aggregateIdColumn,
            string? originalDeliveryDateColumn,
            string? orderDateColumn,
            string locationColumn,
            string confirmationColumn,
            string channelColumn,
            string orderNumberColumn,
            string productColumn)
        {
            var entry = archive.GetEntry(sheetPath)
                ?? throw new InvalidDataException($"正式 Excel 缺少 {sheetKind} 工作表。");
            using var input = entry.Open();
            var document = XDocument.Load(input);
            var sheetData = document.Root?.Element(spreadsheet + "sheetData")
                ?? throw new InvalidDataException($"正式 Excel 的 {sheetKind} 工作表缺少資料。");
            foreach (var row in sheetData.Elements(spreadsheet + "row").Skip(1))
            {
                string Cell(string column, string fieldName)
                {
                    var cell = FindCell(row, column);
                    if (cell is null) return string.Empty;
                    if (cell.Element(spreadsheet + "f") is not null)
                    {
                        throw new InvalidDataException(
                            $"Excel 的「{fieldName}」欄位不可使用公式；請改為固定值後再匯出。");
                    }
                    return ReadCellText(cell, spreadsheet, sharedStrings);
                }

                string PendingSpecifiedDateCell(string column)
                {
                    var cell = FindCell(row, column);
                    if (cell is null) return string.Empty;
                    if (cell.Element(spreadsheet + "f") is not null)
                    {
                        throw new InvalidDataException(
                            "Excel 的「指定到貨日」欄位不可使用公式；請改為固定日期後再匯出。");
                    }

                    if (cell.Attribute("t") is not null)
                    {
                        var textValue = ReadCellText(cell, spreadsheet, sharedStrings).Trim();
                        if (textValue.Length == 0) return string.Empty;
                        throw new InvalidDataException(
                            "日期待確認的「指定到貨日」必須是 Excel 日期固定值，不可輸入文字。");
                    }

                    var raw = cell.Element(spreadsheet + "v")?.Value?.Trim() ?? string.Empty;
                    if (raw.Length == 0) return string.Empty;
                    if (
                        !double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial) ||
                        serial is < 1 or > 2958465)
                    {
                        throw new InvalidDataException(
                            "日期待確認的「指定到貨日」必須是 Excel 日期固定值，不可輸入文字。");
                    }

                    try
                    {
                        return DateOnly.FromDateTime(DateTime.FromOADate(serial))
                            .ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                    }
                    catch (ArgumentException)
                    {
                        throw new InvalidDataException(
                            "日期待確認的「指定到貨日」不是有效的 Excel 日期。");
                    }
                }

                var fingerprints = Cell(fingerprintColumn, "明細識別值")
                    .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                if (fingerprints.Length == 0) continue;
                rows.Add(new WorkbookEditableRow(
                    ManualOverrideWorkflow.BuildRowKey(fingerprints),
                    sheetKind,
                    fingerprints,
                    originalDeliveryDateColumn is null
                        ? string.Empty
                        : sheetKind.Equals("pending", StringComparison.OrdinalIgnoreCase)
                            ? PendingSpecifiedDateCell(originalDeliveryDateColumn)
                            : NormalizeEditableDate(Cell(originalDeliveryDateColumn, "指定到貨日")),
                    orderDateColumn is null
                        ? string.Empty
                        : NormalizeEditableDate(Cell(orderDateColumn, "下單日")),
                    Cell(locationColumn, "地點"),
                    Cell(confirmationColumn, "確認"),
                    Cell(channelColumn, "通路別"),
                    Cell(orderNumberColumn, "訂單編號"),
                    Cell(productColumn, "品名"),
                    Cell(aggregateIdColumn, "訂單識別值")));
            }
        }

        static bool HasLegacyPendingHeaders(
            ZipArchive workbook,
            XNamespace spreadsheetNamespace,
            IReadOnlyList<string> strings)
        {
            using var input = (workbook.GetEntry(PendingDateSheetPath)
                               ?? throw new InvalidDataException("正式 Excel 缺少日期待確認工作表。")).Open();
            var document = XDocument.Load(input);
            var header = document.Descendants(spreadsheetNamespace + "row").FirstOrDefault();
            var headers = header?.Elements(spreadsheetNamespace + "c")
                .Select(cell => ReadCellText(cell, spreadsheetNamespace, strings))
                .ToArray() ?? [];
            return headers.SequenceEqual(LegacyPendingDateHeaders, StringComparer.Ordinal);
        }
    }

    private static IReadOnlyList<OrderAggregate> DetailAggregates(
        IReadOnlyList<OrderChangeEntry> entries) =>
        entries
            .Select(entry => (Entry: entry, Product: ResolveProduct(entry)))
            .Where(item => item.Product is not null &&
                           !string.IsNullOrWhiteSpace(item.Entry.DeliveryDate))
            .GroupBy(
                item => $"{item.Entry.SourceCode.ToUpperInvariant()}\u001f" +
                        $"{item.Entry.ExternalOrderNo.Trim().ToUpperInvariant()}\u001f" +
                        $"{item.Product!.Key}\u001f" +
                        $"{NormalizeDate(item.Entry.OriginalDeliveryDate ?? item.Entry.DeliveryDate)}",
                StringComparer.Ordinal)
            .Select(group =>
            {
                var first = group.First();
                var entry = first.Entry;
                var product = first.Product!;
                var channel = DisplayChannel(entry);
                var originalDeliveryDate = entry.HasManualOriginalDeliveryDate
                    ? NormalizeDate(entry.OriginalDeliveryDate)
                    : NormalizeDate(entry.OriginalDeliveryDate ?? entry.DeliveryDate);
                var deliveryDate = NormalizeDate(entry.DeliveryDate);
                var manualLocation = group.Select(item => item.Entry)
                    .Where(item => item.HasManualLocation)
                    .MaxBy(item => item.ObservedAt);
                var manualConfirmation = group.Select(item => item.Entry)
                    .Where(item => item.HasManualConfirmation)
                    .MaxBy(item => item.ObservedAt);
                var sourceLineFingerprints = group
                    .Select(item => BuildSourceLineFingerprint(item.Entry))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(fingerprint => fingerprint, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                return new OrderAggregate(
                    entry.SourceCode,
                    manualLocation?.SourceLocation ?? ResolveLatestSourceLocation(group.Select(item => item.Entry)),
                    manualLocation is not null,
                    manualConfirmation?.Confirmation ?? string.Empty,
                    manualConfirmation is not null,
                    group.Any(item => item.Entry.IgnoreExistingLocation),
                    group.Any(item => item.Entry.IgnoreExistingConfirmation),
                    channel,
                    product.Name,
                    product.Size,
                    originalDeliveryDate,
                    deliveryDate,
                    group.Sum(item => item.Entry.QuantityChange),
                    entry.ExternalOrderNo.Trim(),
                    ResolveLatestNote(group.Select(item => item.Entry)),
                    BuildAggregateId(product.Key, entry.SourceCode, entry.ExternalOrderNo, originalDeliveryDate),
                    sourceLineFingerprints);
            })
            .Where(aggregate => aggregate.Quantity != 0m)
            .OrderBy(aggregate => aggregate.OrderDate, StringComparer.Ordinal)
            .ThenBy(aggregate => aggregate.ChannelName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(aggregate => aggregate.ProductName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(aggregate => aggregate.OriginalDeliveryDate, StringComparer.Ordinal)
            .ThenBy(aggregate => aggregate.ExternalOrderNo, StringComparer.OrdinalIgnoreCase)
            .ToArray();

    private static IReadOnlyList<PendingDateAggregate> PendingDateAggregates(
        IReadOnlyList<OrderChangeEntry> entries) =>
        entries
            .Select(entry => (Entry: entry, Product: ResolveProduct(entry)))
            .Where(item => item.Product is not null &&
                           string.IsNullOrWhiteSpace(item.Entry.DeliveryDate))
            .GroupBy(
                item => $"{item.Entry.SourceCode.ToUpperInvariant()}\u001f" +
                        $"{item.Entry.ExternalOrderNo.Trim().ToUpperInvariant()}\u001f{item.Product!.Key}",
                StringComparer.Ordinal)
            .Select(group =>
            {
                var first = group.First();
                var entry = first.Entry;
                var product = first.Product!;
                var fingerprints = group.Select(item => BuildSourceLineFingerprint(item.Entry))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                var manualLocation = group.Select(item => item.Entry)
                    .Where(item => item.HasManualLocation)
                    .MaxBy(item => item.ObservedAt);
                var manualConfirmation = group.Select(item => item.Entry)
                    .Where(item => item.HasManualConfirmation)
                    .MaxBy(item => item.ObservedAt);
                return new PendingDateAggregate(
                    entry.SourceCode,
                    manualLocation?.SourceLocation ?? ResolveLatestSourceLocation(group.Select(item => item.Entry)),
                    manualLocation is not null,
                    manualConfirmation?.Confirmation ?? string.Empty,
                    manualConfirmation is not null,
                    group.Any(item => item.Entry.IgnoreExistingLocation),
                    group.Any(item => item.Entry.IgnoreExistingConfirmation),
                    DisplayChannel(entry),
                    product.Name,
                    group.Sum(item => item.Entry.QuantityChange),
                    entry.ExternalOrderNo.Trim(),
                    ResolveLatestNote(group.Select(item => item.Entry)),
                    BuildAggregateId(product.Key, entry.SourceCode, entry.ExternalOrderNo, string.Empty),
                    fingerprints);
            })
            .Where(aggregate => aggregate.Quantity != 0m)
            .OrderBy(aggregate => aggregate.ChannelName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(aggregate => aggregate.ProductName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(aggregate => aggregate.ExternalOrderNo, StringComparer.OrdinalIgnoreCase)
            .ToArray();

    private static void AppendExisting(string fullPath, IReadOnlyList<OrderChangeEntry> entries)
    {
        var temporaryPath = fullPath + $".{Guid.NewGuid():N}.tmp";
        try
        {
            File.Copy(fullPath, temporaryPath, overwrite: true);
            using (var archive = ZipFile.Open(temporaryPath, ZipArchiveMode.Update))
            {
                ValidateArchiveSafety(archive);
                ValidateWorkbook(archive);
                EnsureStyleParts(archive);
                EnsureCalculationSettings(archive);
                var details = DetailAggregates(entries);
                RewriteDetailWorksheet(archive, details);
                RewritePendingDateWorksheet(archive, PendingDateAggregates(entries));
                RewriteStatisticsWorksheet(archive, details);
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

    private static void ValidateWorkbook(ZipArchive archive)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new InvalidDataException("選取的 Excel 不是中秋禮盒統計活頁簿，請另存新檔。");
        XNamespace spreadsheet = SpreadsheetNamespace;
        string?[] names;
        using (var input = workbookEntry.Open())
        {
            var document = XDocument.Load(input);
            names = document.Descendants(spreadsheet + "sheet")
                .Select(sheet => (string?)sheet.Attribute("name"))
                .Where(name => name is not null)
                .ToArray();
        }
        string[] expectedNames = [DetailSheetName, PendingDateSheetName, StatisticsSheetName];
        if (!names.SequenceEqual(expectedNames, StringComparer.Ordinal))
        {
            throw new InvalidDataException(
                "選取的 Excel 不是含有「訂單明細、日期待確認、統計」的新版中秋禮盒統計活頁簿，請另存新檔。");
        }

        var sharedStrings = ReadSharedStrings(archive, spreadsheet);
        var detailEntry = archive.GetEntry(DetailSheetPath)
            ?? throw new InvalidDataException("選取的 Excel 缺少訂單明細工作表，請另存新檔。");
        using (var detailInput = detailEntry.Open())
        {
            var worksheet = XDocument.Load(detailInput);
            var firstRow = worksheet.Descendants(spreadsheet + "row")
                .FirstOrDefault(row => (string?)row.Attribute("r") == "1");
            var headers = firstRow?.Elements(spreadsheet + "c")
                .Select(cell => ReadCellText(cell, spreadsheet, sharedStrings))
                .ToArray() ?? [];
            var hiddenIdColumn = worksheet.Descendants(spreadsheet + "col").Any(column =>
                (string?)column.Attribute("min") == "10" &&
                (string?)column.Attribute("max") == "11" &&
                (string?)column.Attribute("hidden") == "1");
            if (!headers.SequenceEqual(DetailHeaders, StringComparer.Ordinal) || !hiddenIdColumn)
            {
                throw new InvalidDataException(
                    "選取的 Excel 不是由本程式建立的訂單明細活頁簿，請另存新檔。");
            }
        }

        var pendingEntry = archive.GetEntry(PendingDateSheetPath)
            ?? throw new InvalidDataException("選取的 Excel 缺少日期待確認工作表，請另存新檔。");
        using var pendingInput = pendingEntry.Open();
        var pendingDocument = XDocument.Load(pendingInput);
        var pendingFirstRow = pendingDocument.Descendants(spreadsheet + "row")
            .FirstOrDefault(row => (string?)row.Attribute("r") == "1");
        var pendingHeaders = pendingFirstRow?.Elements(spreadsheet + "c")
            .Select(cell => ReadCellText(cell, spreadsheet, sharedStrings))
            .ToArray() ?? [];
        var isCurrentPending = pendingHeaders.SequenceEqual(PendingDateHeaders, StringComparer.Ordinal);
        var isLegacyPending = pendingHeaders.SequenceEqual(LegacyPendingDateHeaders, StringComparer.Ordinal);
        var pendingHiddenIds = pendingDocument.Descendants(spreadsheet + "col").Any(column =>
            (string?)column.Attribute("min") == (isCurrentPending ? "9" : "8") &&
            (string?)column.Attribute("max") == (isCurrentPending ? "10" : "9") &&
            (string?)column.Attribute("hidden") == "1");
        if ((!isCurrentPending && !isLegacyPending) || !pendingHiddenIds)
        {
            throw new InvalidDataException("日期待確認工作表格式不正確，請另存新檔。");
        }

        var statisticsEntry = archive.GetEntry(StatisticsSheetPath)
            ?? throw new InvalidDataException("選取的 Excel 缺少統計工作表，請另存新檔。");
        using var statisticsInput = statisticsEntry.Open();
        var statisticsDocument = XDocument.Load(statisticsInput);
        var statisticsFirstRow = statisticsDocument.Descendants(spreadsheet + "row")
            .FirstOrDefault(row => (string?)row.Attribute("r") == "1");
        var statisticsHeaders = statisticsFirstRow?.Elements(spreadsheet + "c")
            .Select(cell => ReadCellText(cell, spreadsheet, sharedStrings))
            .ToArray() ?? [];
        if (!statisticsHeaders.SequenceEqual(StatisticsHeaders, StringComparer.Ordinal) &&
            !LegacyStatisticsHeaders.Any(legacy =>
                statisticsHeaders.SequenceEqual(legacy, StringComparer.Ordinal)))
        {
            throw new InvalidDataException("統計工作表格式不正確，請另存新檔。");
        }
    }

    private static void WriteDetailWorksheet(
        ZipArchive archive,
        IReadOnlyList<OrderAggregate> aggregates,
        IReadOnlyList<ExistingAggregateRow> existingRows)
    {
        var entry = archive.CreateEntry(DetailSheetPath, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, XmlSettings);
        writer.WriteStartDocument(true);
        writer.WriteStartElement("worksheet", SpreadsheetNamespace);
        WriteHiddenIdColumns(writer, 10, 11);
        writer.WriteStartElement("sheetData", SpreadsheetNamespace);
        WriteHeaderRow(writer, DetailHeaders);
        for (var index = 0; index < aggregates.Count; index++)
        {
            WriteDetailRow(
                writer,
                index + 2,
                aggregates[index],
                ResolveLocation(aggregates[index], existingRows),
                ResolveConfirmation(aggregates[index], existingRows));
        }

        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void RewriteDetailWorksheet(
        ZipArchive archive,
        IReadOnlyList<OrderAggregate> aggregates)
    {
        var worksheetEntry = archive.GetEntry(DetailSheetPath)
            ?? throw new InvalidDataException("訂單明細工作表不完整，請另存新檔。");
        XDocument document;
        using (var input = worksheetEntry.Open())
        {
            document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        }

        XNamespace spreadsheet = SpreadsheetNamespace;
        var sharedStrings = ReadSharedStrings(archive, spreadsheet);
        var existingRows = ReadExistingRows(document, spreadsheet, sharedStrings, "J", "G", "F", "K");
        worksheetEntry.Delete();
        WriteDetailWorksheet(archive, aggregates, existingRows);
    }

    private static void WritePendingDateWorksheet(
        ZipArchive archive,
        IReadOnlyList<PendingDateAggregate> aggregates,
        IReadOnlyList<ExistingAggregateRow> existingRows)
    {
        var entry = archive.CreateEntry(PendingDateSheetPath, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, XmlSettings);
        writer.WriteStartDocument(true);
        writer.WriteStartElement("worksheet", SpreadsheetNamespace);
        WriteHiddenIdColumns(writer, 9, 10);
        writer.WriteStartElement("sheetData", SpreadsheetNamespace);
        WriteHeaderRow(writer, PendingDateHeaders);

        for (var index = 0; index < aggregates.Count; index++)
        {
            var rowNumber = index + 2;
            var aggregate = aggregates[index];
            var confirmation = aggregate.HasManualConfirmation || aggregate.IgnoreExistingConfirmation
                ? aggregate.Confirmation
                : ResolveConfirmation(aggregate.AggregateId, aggregate.SourceLineFingerprints, existingRows);
            var location = aggregate.HasManualLocation
                ? aggregate.SourceLocation
                : aggregate.IgnoreExistingLocation
                    ? ResolveLocation(
                        aggregate.SourceCode,
                        aggregate.ChannelName,
                        aggregate.SourceLocation,
                        null)
                : ResolveLocation(
                    aggregate.SourceCode,
                    aggregate.ChannelName,
                    aggregate.SourceLocation,
                    ResolveExistingLocation(aggregate.AggregateId, aggregate.SourceLineFingerprints, existingRows));
            writer.WriteStartElement("row", SpreadsheetNamespace);
            writer.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
            WriteTextCell(writer, $"A{rowNumber}", aggregate.ChannelName);
            WriteTextCell(writer, $"B{rowNumber}", aggregate.ProductName);
            WriteTextCell(writer, $"C{rowNumber}", string.Empty);
            WriteNumberCell(writer, $"D{rowNumber}", aggregate.Quantity);
            WriteTextCell(writer, $"E{rowNumber}", confirmation);
            WriteTextCell(writer, $"F{rowNumber}", location);
            WriteTextCell(writer, $"G{rowNumber}", aggregate.ExternalOrderNo);
            WriteTextCell(writer, $"H{rowNumber}", aggregate.Note);
            WriteTextCell(writer, $"I{rowNumber}", aggregate.AggregateId);
            WriteTextCell(writer, $"J{rowNumber}", string.Join(",", aggregate.SourceLineFingerprints));
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void RewritePendingDateWorksheet(
        ZipArchive archive,
        IReadOnlyList<PendingDateAggregate> aggregates)
    {
        var existing = archive.GetEntry(PendingDateSheetPath)
            ?? throw new InvalidDataException("日期待確認工作表不完整，請另存新檔。");
        XDocument document;
        using (var input = existing.Open()) document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        XNamespace spreadsheet = SpreadsheetNamespace;
        var sharedStrings = ReadSharedStrings(archive, spreadsheet);
        var headers = document.Descendants(spreadsheet + "row").FirstOrDefault()?
            .Elements(spreadsheet + "c")
            .Select(cell => ReadCellText(cell, spreadsheet, sharedStrings))
            .ToArray() ?? [];
        var existingRows = headers.SequenceEqual(LegacyPendingDateHeaders, StringComparer.Ordinal)
            ? ReadExistingRows(document, spreadsheet, sharedStrings, "H", "E", "D", "I")
            : ReadExistingRows(document, spreadsheet, sharedStrings, "I", "F", "E", "J");
        existing.Delete();
        WritePendingDateWorksheet(archive, aggregates, existingRows);
    }

    private static void WriteStatisticsWorksheet(
        ZipArchive archive,
        IReadOnlyList<OrderAggregate> aggregates)
    {
        var entry = archive.CreateEntry(StatisticsSheetPath, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, XmlSettings);
        writer.WriteStartDocument(true);
        writer.WriteStartElement("worksheet", SpreadsheetNamespace);
        WriteHiddenIdColumns(writer, 12, 13);
        writer.WriteStartElement("sheetData", SpreadsheetNamespace);
        WriteHeaderRow(writer, StatisticsHeaders);
        var lastRow = aggregates.Count + 1;
        for (var index = 0; index < aggregates.Count; index++)
        {
            var rowNumber = index + 2;
            writer.WriteStartElement("row", SpreadsheetNamespace);
            writer.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
            var uniqueSourceRowFormula =
                $"IF(AND('訂單明細'!D{rowNumber}<>\"\"," +
                $"COUNTIFS('訂單明細'!$D$2:D{rowNumber},'訂單明細'!D{rowNumber}," +
                $"'訂單明細'!$C$2:C{rowNumber},'訂單明細'!C{rowNumber}," +
                $"'訂單明細'!$G$2:G{rowNumber},'訂單明細'!G{rowNumber})=1),ROW()-1,\"\")";
            var compactIndexFormula =
                $"IFERROR(_xlfn.AGGREGATE(15,6,$L$2:$L${lastRow}/($L$2:$L${lastRow}<>\"\")," +
                $"ROW()-1),\"\")";
            WriteFormulaCell(writer, $"A{rowNumber}",
                $"IF($M{rowNumber}=\"\",\"\",INDEX('訂單明細'!$D$2:$D${lastRow},$M{rowNumber}))", 2);
            WriteFormulaCell(writer, $"B{rowNumber}",
                $"IF($M{rowNumber}=\"\",\"\",INDEX('訂單明細'!$C$2:$C${lastRow},$M{rowNumber}))", 2);
            WriteFormulaCell(writer, $"C{rowNumber}",
                $"IF($M{rowNumber}=\"\",\"\",IF(INDEX('訂單明細'!$G$2:$G${lastRow},$M{rowNumber})=\"\",\"未設定\",INDEX('訂單明細'!$G$2:$G${lastRow},$M{rowNumber})))");
            for (var productIndex = 0; productIndex < 8; productIndex++)
            {
                var column = ColumnName(productIndex + 4);
                var formula = $"IF($M{rowNumber}=\"\",\"\",SUMIFS('訂單明細'!$E$2:$E${lastRow}," +
                              $"'訂單明細'!$D$2:$D${lastRow},$A{rowNumber}," +
                              $"'訂單明細'!$C$2:$C${lastRow},$B{rowNumber}," +
                              $"'訂單明細'!$G$2:$G${lastRow},IF($C{rowNumber}=\"未設定\",\"\",$C{rowNumber})," +
                              $"'訂單明細'!$B$2:$B${lastRow},{column}$1))";
                WriteFormulaCell(writer, $"{column}{rowNumber}", formula);
            }
            WriteFormulaCell(writer, $"L{rowNumber}", uniqueSourceRowFormula);
            WriteFormulaCell(writer, $"M{rowNumber}", compactIndexFormula);
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void RewriteStatisticsWorksheet(
        ZipArchive archive,
        IReadOnlyList<OrderAggregate> aggregates)
    {
        var existing = archive.GetEntry(StatisticsSheetPath)
            ?? throw new InvalidDataException("統計工作表不完整，請另存新檔。");
        existing.Delete();
        WriteStatisticsWorksheet(archive, aggregates);
    }

    private static void WriteHiddenIdColumns(XmlWriter writer, int minimum, int maximum)
    {
        writer.WriteStartElement("cols", SpreadsheetNamespace);
        writer.WriteStartElement("col", SpreadsheetNamespace);
        writer.WriteAttributeString("min", minimum.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("max", maximum.ToString(CultureInfo.InvariantCulture));
        writer.WriteAttributeString("hidden", "1");
        writer.WriteAttributeString("width", "1");
        writer.WriteAttributeString("customWidth", "1");
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteHeaderRow(XmlWriter writer, IReadOnlyList<string> headers)
    {
        writer.WriteStartElement("row", SpreadsheetNamespace);
        writer.WriteAttributeString("r", "1");
        for (var index = 0; index < headers.Count; index++)
        {
            WriteTextCell(writer, $"{ColumnName(index + 1)}1", headers[index]);
        }

        writer.WriteEndElement();
    }

    private static void WriteDetailRow(
        XmlWriter writer,
        int rowNumber,
        OrderAggregate aggregate,
        string location,
        string confirmation)
    {
        writer.WriteStartElement("row", SpreadsheetNamespace);
        writer.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
        WriteTextCell(writer, $"A{rowNumber}", aggregate.ChannelName);
        WriteTextCell(writer, $"B{rowNumber}", aggregate.ProductName);
        WriteDateCell(writer, $"C{rowNumber}", aggregate.OriginalDeliveryDate);
        WriteDateCell(writer, $"D{rowNumber}", aggregate.OrderDate);
        WriteNumberCell(writer, $"E{rowNumber}", aggregate.Quantity);
        WriteTextCell(writer, $"F{rowNumber}", confirmation);
        WriteTextCell(writer, $"G{rowNumber}", location);
        WriteTextCell(writer, $"H{rowNumber}", aggregate.ExternalOrderNo);
        WriteTextCell(writer, $"I{rowNumber}", aggregate.Note);
        WriteTextCell(writer, $"J{rowNumber}", aggregate.AggregateId);
        WriteTextCell(writer, $"K{rowNumber}", string.Join(",", aggregate.SourceLineFingerprints));
        writer.WriteEndElement();
    }

    private static XElement? FindCell(XElement row, string column) =>
        row.Elements(XName.Get("c", SpreadsheetNamespace))
            .FirstOrDefault(cell => ((string?)cell.Attribute("r"))?.StartsWith(
                column, StringComparison.OrdinalIgnoreCase) == true);


    private static void WriteTextCell(XmlWriter writer, string reference, string value)
    {
        writer.WriteStartElement("c", SpreadsheetNamespace);
        writer.WriteAttributeString("r", reference);
        writer.WriteAttributeString("s", "1");
        writer.WriteAttributeString("t", "inlineStr");
        writer.WriteStartElement("is", SpreadsheetNamespace);
        writer.WriteStartElement("t", SpreadsheetNamespace);
        writer.WriteAttributeString("xml", "space", null, "preserve");
        writer.WriteString(RemoveInvalidXmlCharacters(value));
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteNumberCell(XmlWriter writer, string reference, decimal value)
    {
        writer.WriteStartElement("c", SpreadsheetNamespace);
        writer.WriteAttributeString("r", reference);
        writer.WriteAttributeString("s", "1");
        writer.WriteStartElement("v", SpreadsheetNamespace);
        writer.WriteString(value.ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteDateCell(XmlWriter writer, string reference, string value)
    {
        if (!DateOnly.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var date))
        {
            WriteTextCell(writer, reference, value);
            return;
        }

        writer.WriteStartElement("c", SpreadsheetNamespace);
        writer.WriteAttributeString("r", reference);
        writer.WriteAttributeString("s", "2");
        writer.WriteStartElement("v", SpreadsheetNamespace);
        writer.WriteString(date.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteFormulaCell(
        XmlWriter writer,
        string reference,
        string formula,
        int styleIndex = 1)
    {
        writer.WriteStartElement("c", SpreadsheetNamespace);
        writer.WriteAttributeString("r", reference);
        writer.WriteAttributeString("s", styleIndex.ToString(CultureInfo.InvariantCulture));
        writer.WriteStartElement("f", SpreadsheetNamespace);
        writer.WriteString(formula);
        writer.WriteEndElement();
        writer.WriteStartElement("v", SpreadsheetNamespace);
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static IReadOnlyList<ExistingAggregateRow> ReadExistingRows(
        XDocument document,
        XNamespace spreadsheet,
        IReadOnlyList<string> sharedStrings,
        string aggregateIdColumn,
        string locationColumn,
        string confirmationColumn,
        string fingerprintColumn)
    {
        var sheetData = document.Root?.Element(spreadsheet + "sheetData")
            ?? throw new InvalidDataException("中秋禮盒統計工作表格式不完整。");
        return sheetData.Elements(spreadsheet + "row")
            .Skip(1)
            .Select(row =>
            {
                string Cell(string column)
                {
                    var cell = FindCell(row, column);
                    return cell is null ? string.Empty : ReadCellText(cell, spreadsheet, sharedStrings);
                }
                string ManualCell(string column)
                {
                    var cell = FindCell(row, column);
                    return cell is null ? string.Empty : ReadManualCellText(cell, spreadsheet, sharedStrings);
                }

                return new ExistingAggregateRow(
                    Cell(aggregateIdColumn),
                    ManualCell(locationColumn),
                    ManualCell(confirmationColumn),
                    Cell(fingerprintColumn).Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase));
            })
            .ToArray();
    }

    private static string ResolveConfirmation(
        OrderAggregate aggregate,
        IReadOnlyList<ExistingAggregateRow> existingRows)
        => aggregate.HasManualConfirmation || aggregate.IgnoreExistingConfirmation
            ? aggregate.Confirmation
            : ResolveConfirmation(aggregate.AggregateId, aggregate.SourceLineFingerprints, existingRows);

    private static string ResolveConfirmation(
        string aggregateId,
        IReadOnlyList<string> sourceLineFingerprints,
        IReadOnlyList<ExistingAggregateRow> existingRows)
    {
        var exact = existingRows.FirstOrDefault(row =>
            row.AggregateId.Equals(aggregateId, StringComparison.OrdinalIgnoreCase));
        if (exact is not null && !string.IsNullOrWhiteSpace(exact.Confirmation))
        {
            return exact.Confirmation;
        }

        var fingerprints = sourceLineFingerprints.ToHashSet(StringComparer.OrdinalIgnoreCase);
        return existingRows.FirstOrDefault(row =>
                   !string.IsNullOrWhiteSpace(row.Confirmation) &&
                   row.SourceLineFingerprints.Overlaps(fingerprints))
               ?.Confirmation ?? string.Empty;
    }

    private static string ResolveLocation(
        OrderAggregate aggregate,
        IReadOnlyList<ExistingAggregateRow> existingRows)
        => aggregate.HasManualLocation
            ? aggregate.SourceLocation
            : aggregate.IgnoreExistingLocation
                ? ResolveLocation(
                    aggregate.SourceCode,
                    aggregate.ChannelName,
                    aggregate.SourceLocation,
                    null)
            : ResolveLocation(
                aggregate.SourceCode,
                aggregate.ChannelName,
                aggregate.SourceLocation,
                ResolveExistingLocation(aggregate.AggregateId, aggregate.SourceLineFingerprints, existingRows));

    private static string? ResolveExistingLocation(
        string aggregateId,
        IReadOnlyList<string> sourceLineFingerprints,
        IReadOnlyList<ExistingAggregateRow> existingRows)
    {
        var exact = existingRows.FirstOrDefault(row =>
            row.AggregateId.Equals(aggregateId, StringComparison.OrdinalIgnoreCase));
        if (exact is not null && !string.IsNullOrWhiteSpace(exact.Location))
        {
            return exact.Location;
        }

        var fingerprints = sourceLineFingerprints.ToHashSet(StringComparer.OrdinalIgnoreCase);
        return existingRows.FirstOrDefault(row =>
                   !string.IsNullOrWhiteSpace(row.Location) &&
                   row.SourceLineFingerprints.Overlaps(fingerprints))
               ?.Location;
    }

    private static string ResolveLocation(
        string sourceCode,
        string channelName,
        string? sourceLocation,
        string? existingLocation)
    {
        if (!string.IsNullOrWhiteSpace(existingLocation))
        {
            return existingLocation;
        }

        if (!string.IsNullOrWhiteSpace(sourceLocation))
        {
            return sourceLocation.Trim();
        }

        if (sourceCode.Equals("pos_excel", StringComparison.OrdinalIgnoreCase))
        {
            return "躉泰";
        }

        if (sourceCode.Equals("erp", StringComparison.OrdinalIgnoreCase) ||
            sourceCode.Equals("erp_excel", StringComparison.OrdinalIgnoreCase))
        {
            return "彰廠";
        }

        var normalizedChannel = channelName.Trim();
        if ((sourceCode.Equals("site1", StringComparison.OrdinalIgnoreCase) ||
             sourceCode.Equals("site2", StringComparison.OrdinalIgnoreCase)) &&
            WebsiteChannelLocations.TryGetValue(normalizedChannel, out var websiteLocation))
        {
            return websiteLocation;
        }

        return string.Empty;
    }

    private static string DisplayChannel(OrderChangeEntry entry) =>
        (entry.ChannelName ?? entry.SourceName).Trim();

    private static string ResolveLatestNote(IEnumerable<OrderChangeEntry> entries) =>
        entries.MaxBy(entry => entry.ObservedAt)!.Note?.Trim() ?? string.Empty;

    private static string ResolveLatestSourceLocation(IEnumerable<OrderChangeEntry> entries) =>
        entries.Where(entry => !string.IsNullOrWhiteSpace(entry.SourceLocation))
            .MaxBy(entry => entry.ObservedAt)?.SourceLocation?.Trim() ?? string.Empty;

    private static int? ResolveGiftBoxSize(OrderChangeEntry entry) =>
        entry.GiftBoxSize ?? GiftBoxSizePolicy.Resolve(entry.Sku, entry.ProductName);

    private static ProductDescriptor? ResolveProduct(OrderChangeEntry entry)
    {
        if (entry.SourceCode.Equals("pos_excel", StringComparison.OrdinalIgnoreCase))
        {
            var productKey = $"pos:{entry.ExternalLineKey}";
            return entry.Sku?.Trim() switch
            {
                "1902072" => new ProductDescriptor(productKey, "單顆(有盒)(霧面袋)", 1),
                "1902248" => new ProductDescriptor(productKey, "6入鐵盒2026年鐵盒(黑金袋)", 6),
                "1902247" => new ProductDescriptor(productKey, "9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)", 9),
                _ => null
            };
        }

        if ((entry.SourceCode.Equals("site1", StringComparison.OrdinalIgnoreCase) ||
             entry.SourceCode.Equals("site2", StringComparison.OrdinalIgnoreCase)) &&
            GiftBoxItemPolicy.TryResolve(entry.Sku, out var wmsItem))
        {
            return new ProductDescriptor($"sku:{wmsItem.ItemNo}", wmsItem.DisplayName, wmsItem.PieceCount);
        }

        if (entry.SourceCode.Equals("manual_excel", StringComparison.OrdinalIgnoreCase) ||
            entry.SourceCode.Equals("erp", StringComparison.OrdinalIgnoreCase))
        {
            if (entry.Sku == "1902199") return new ProductDescriptor("sku:1902199", "單顆(無盒)(透明袋)", 1);
            if (entry.Sku == "192199") return new ProductDescriptor("sku:192199", "單顆(無盒)(霧面袋)", 1);
            if (entry.Sku == "1902072") return new ProductDescriptor("sku:1902072", "單顆(有盒)(霧面袋)", 1);
            if (entry.Sku == "1902042") return new ProductDescriptor("sku:1902042", "3入條裝(無盒)(霧面袋)", 3);
            if (entry.Sku == "1902226") return new ProductDescriptor("sku:1902226", "3入禮盒紙盒(黑金袋)", 3);
            if (entry.Sku == "1902197") return new ProductDescriptor("sku:1902197", "6入鐵盒2024年鐵盒(黑金袋)", 6);
            if (entry.Sku == "1902248") return new ProductDescriptor("sku:1902248", "6入鐵盒2026年鐵盒(黑金袋)", 6);
            if (entry.Sku == "1902247") return new ProductDescriptor("sku:1902247", "9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)", 9);
        }

        return ResolveGiftBoxSize(entry) switch
        {
            3 => new ProductDescriptor("size:3", "3入禮盒紙盒(黑金袋)", 3),
            6 => new ProductDescriptor("size:6", "6入鐵盒2026年鐵盒(黑金袋)", 6),
            9 => new ProductDescriptor("size:9", "9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)", 9),
            _ => null
        };
    }

    private static string NormalizeDate(string? value)
    {
        var normalized = value?.Trim() ?? string.Empty;
        string[] formats = ["yyyy/M/d", "yyyy/MM/dd", "yyyy-M-d", "yyyy-MM-dd"];
        return DateOnly.TryParseExact(
            normalized,
            formats,
            CultureInfo.InvariantCulture,
            DateTimeStyles.None,
            out var date)
            ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : normalized;
    }

    private static string NormalizeEditableDate(string? value)
    {
        var normalized = value?.Trim() ?? string.Empty;
        if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial) &&
            serial is >= 1 and <= 2958465)
        {
            try
            {
                return DateOnly.FromDateTime(DateTime.FromOADate(serial))
                    .ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            catch (ArgumentException)
            {
                // Preserve an invalid user-entered value so it remains visible for correction.
            }
        }

        return NormalizeDate(normalized);
    }

    private static bool IsDateBlocked(string? shippingDateStatus) =>
        shippingDateStatus?.Trim() is { } status &&
        (status.Equals("conflict", StringComparison.OrdinalIgnoreCase) ||
         status.Equals("needs_review", StringComparison.OrdinalIgnoreCase));

    private static string BuildAggregateId(
        string productKey,
        string sourceCode,
        string externalOrderNo,
        string originalDeliveryDate)
    {
        var identity = $"{productKey}\u001f{sourceCode.ToUpperInvariant()}\u001f" +
                       $"{externalOrderNo.Trim().ToUpperInvariant()}\u001f{originalDeliveryDate}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity)));
    }

    private static string BuildSourceLineFingerprint(OrderChangeEntry entry)
        => ManualOverrideWorkflow.BuildEntryFingerprint(entry);

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
            int.TryParse(cell.Element(spreadsheet + "v")?.Value, out var index) &&
            index >= 0 && index < sharedStrings.Count)
        {
            return sharedStrings[index];
        }

        var inlineText = string.Concat(cell.Descendants(spreadsheet + "t").Select(text => text.Value));
        return inlineText.Length > 0 ? inlineText : cell.Element(spreadsheet + "v")?.Value ?? string.Empty;
    }

    private static string ReadManualCellText(
        XElement cell,
        XNamespace spreadsheet,
        IReadOnlyList<string> sharedStrings)
    {
        if (cell.Element(spreadsheet + "f") is not null)
        {
            throw new InvalidDataException(
                "Excel 的「地點」或「確認」欄位不可使用公式；原檔未被修改，請改為固定值後再匯出。");
        }
        return ReadCellText(cell, spreadsheet, sharedStrings);
    }

    private static void ValidateArchiveSafety(ZipArchive archive)
    {
        const int maximumEntries = 200;
        const long maximumEntryBytes = 20 * 1024 * 1024;
        const long maximumTotalBytes = 100 * 1024 * 1024;
        if (archive.Entries.Count > maximumEntries ||
            archive.Entries.Any(entry => entry.Length > maximumEntryBytes) ||
            archive.Entries.Sum(entry => entry.Length) > maximumTotalBytes)
        {
            throw new InvalidDataException(
                "Excel 檔案過大或內容異常，為保護系統已停止更新；原檔未被修改。");
        }
    }

    private static string RemoveInvalidXmlCharacters(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var rune in value.EnumerateRunes())
        {
            var codePoint = rune.Value;
            if (codePoint is 0x9 or 0xA or 0xD ||
                codePoint is >= 0x20 and <= 0xD7FF ||
                codePoint is >= 0xE000 and <= 0xFFFD ||
                codePoint is >= 0x10000 and <= 0x10FFFF)
            {
                builder.Append(rune.ToString());
            }
        }

        return builder.ToString();
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

    private static string EntryLineKey(OrderChangeEntry entry) =>
        $"{entry.SourceCode}\u001f{entry.ExternalOrderNo}\u001f{entry.ExternalLineKey}";

    private static string SnapshotLineKey(OrderLineSnapshot snapshot) =>
        $"{snapshot.SourceCode}\u001f{snapshot.ExternalOrderNo}\u001f{snapshot.ExternalLineKey}";

    private static int? ResolveSnapshotGiftBoxSize(
        OrderLineSnapshot parent,
        IReadOnlyList<OrderLineSnapshot> snapshots) =>
        GiftBoxSizePolicy.ResolveParent(parent, snapshots);

    private static SourceSyncFileLock AcquireWorkbookLock(string fullPath)
    {
        var lockDirectory = Path.Combine(Path.GetTempPath(), "MidAutumnGiftBox", "locks");
        var lockCode = "excel-" + Convert.ToHexString(
            SHA256.HashData(Encoding.UTF8.GetBytes(fullPath.ToUpperInvariant())))[..24];
        return SourceSyncFileLock.TryAcquire(lockDirectory, lockCode)
               ?? throw new IOException("這份 Excel 正由另一個程式匯出，請稍後再試。");
    }

    private static void EnsureStyleParts(ZipArchive archive)
    {
        var existingStyles = archive.GetEntry("xl/styles.xml");
        existingStyles?.Delete();
        WriteTextEntry(archive, "xl/styles.xml", StylesXml);

        XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        var contentTypesEntry = archive.GetEntry("[Content_Types].xml")
            ?? throw new InvalidDataException("Excel 缺少 Content Types。");
        XDocument contentTypesDocument;
        using (var input = contentTypesEntry.Open()) contentTypesDocument = XDocument.Load(input);
        if (!contentTypesDocument.Descendants(contentTypes + "Override").Any(element =>
                (string?)element.Attribute("PartName") == "/xl/styles.xml"))
        {
            contentTypesDocument.Root?.Add(new XElement(contentTypes + "Override",
                new XAttribute("PartName", "/xl/styles.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")));
        }
        contentTypesEntry.Delete();
        WriteDocumentEntry(archive, "[Content_Types].xml", contentTypesDocument);

        XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        var relationshipsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new InvalidDataException("Excel 缺少 Workbook Relationships。");
        XDocument relationshipsDocument;
        using (var input = relationshipsEntry.Open()) relationshipsDocument = XDocument.Load(input);
        const string stylesRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        if (!relationshipsDocument.Descendants(relationships + "Relationship").Any(element =>
                (string?)element.Attribute("Type") == stylesRelationshipType))
        {
            relationshipsDocument.Root?.Add(new XElement(relationships + "Relationship",
                new XAttribute("Id", "rIdStyles"),
                new XAttribute("Type", stylesRelationshipType),
                new XAttribute("Target", "styles.xml")));
        }
        relationshipsEntry.Delete();
        WriteDocumentEntry(archive, "xl/_rels/workbook.xml.rels", relationshipsDocument);
    }

    private static void EnsureCalculationSettings(ZipArchive archive)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new InvalidDataException("選取的 Excel 缺少活頁簿資訊，請另存新檔。");
        XDocument workbook;
        using (var input = workbookEntry.Open()) workbook = XDocument.Load(input);
        XNamespace spreadsheet = SpreadsheetNamespace;
        var calculation = workbook.Root?.Element(spreadsheet + "calcPr");
        if (calculation is null)
        {
            calculation = new XElement(spreadsheet + "calcPr");
            workbook.Root?.Add(calculation);
        }
        calculation.SetAttributeValue("calcId", "191029");
        calculation.SetAttributeValue("calcMode", "auto");
        calculation.SetAttributeValue("fullCalcOnLoad", "1");
        calculation.SetAttributeValue("forceFullCalc", "1");
        workbookEntry.Delete();
        WriteDocumentEntry(archive, "xl/workbook.xml", workbook);
    }

    private static void WriteDocumentEntry(ZipArchive archive, string path, XDocument document)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var output = entry.Open();
        document.Save(output);
    }

    private static void WriteTextEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private sealed record OrderAggregate(
        string SourceCode,
        string SourceLocation,
        bool HasManualLocation,
        string Confirmation,
        bool HasManualConfirmation,
        bool IgnoreExistingLocation,
        bool IgnoreExistingConfirmation,
        string ChannelName,
        string ProductName,
        int ProductSize,
        string OriginalDeliveryDate,
        string OrderDate,
        decimal Quantity,
        string ExternalOrderNo,
        string Note,
        string AggregateId,
        IReadOnlyList<string> SourceLineFingerprints);

    private sealed record ExistingAggregateRow(
        string AggregateId,
        string Location,
        string Confirmation,
        HashSet<string> SourceLineFingerprints);

    private sealed record PendingDateAggregate(
        string SourceCode,
        string SourceLocation,
        bool HasManualLocation,
        string Confirmation,
        bool HasManualConfirmation,
        bool IgnoreExistingLocation,
        bool IgnoreExistingConfirmation,
        string ChannelName,
        string ProductName,
        decimal Quantity,
        string ExternalOrderNo,
        string Note,
        string AggregateId,
        IReadOnlyList<string> SourceLineFingerprints);

    private sealed record ProductDescriptor(string Key, string Name, int Size);

    private sealed record SnapshotDateResolution(
        string? OriginalDeliveryDate,
        string? DeliveryDate,
        bool BlocksDate,
        int? GiftBoxSize,
        string? Note);

    private const string ContentTypesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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
            <sheet name="訂單明細" sheetId="1" r:id="rId1"/>
            <sheet name="日期待確認" sheetId="2" r:id="rId2"/>
            <sheet name="統計" sheetId="3" r:id="rId3"/>
          </sheets>
          <calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>
        </workbook>
        """;

    private const string WorkbookRelationshipsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
          <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
        """;

    private const string StylesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy/mm/dd"/></numFmts>
          <fonts count="2">
            <font><sz val="11"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
            <font><sz val="12"/><name val="標楷體"/><family val="3"/><charset val="136"/></font>
          </fonts>
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
          </fills>
          <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
          <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
          <cellXfs count="3">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>
          </cellXfs>
          <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        </styleSheet>
        """;
}
