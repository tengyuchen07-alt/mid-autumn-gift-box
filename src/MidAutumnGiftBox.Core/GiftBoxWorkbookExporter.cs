using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace MidAutumnGiftBox.Core;

public static class GiftBoxWorkbookExporter
{
    private const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly string[] Headers = ["通路別", "訂單編號", "指定出貨日", "數量", "確認", "系統異動ID"];
    private static readonly (int Size, string Name, string Path)[] Sheets =
    [
        (3, "三入", "xl/worksheets/sheet1.xml"),
        (6, "六入", "xl/worksheets/sheet2.xml"),
        (9, "九入", "xl/worksheets/sheet3.xml")
    ];
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
        foreach (var sheet in Sheets)
        {
            WriteWorksheet(archive, sheet.Path, EntriesForSize(entries, sheet.Size));
        }
    }

    private static IReadOnlyList<OrderChangeEntry> EntriesForSize(
        IReadOnlyList<OrderChangeEntry> entries,
        int size) =>
        entries
            .Where(entry => GiftBoxSizePolicy.Resolve(entry.Sku, entry.ProductName) == size)
            .OrderBy(entry => entry.ObservedAt)
            .ThenBy(entry => entry.SourceCode, StringComparer.OrdinalIgnoreCase)
            .ThenBy(entry => entry.ExternalOrderNo, StringComparer.OrdinalIgnoreCase)
            .ToArray();

    private static void AppendExisting(string fullPath, IReadOnlyList<OrderChangeEntry> entries)
    {
        var temporaryPath = fullPath + $".{Guid.NewGuid():N}.tmp";
        try
        {
            File.Copy(fullPath, temporaryPath, overwrite: true);
            using (var archive = ZipFile.Open(temporaryPath, ZipArchiveMode.Update))
            {
                ValidateWorkbook(archive);
                foreach (var sheet in Sheets)
                {
                    AppendWorksheet(archive, sheet.Path, EntriesForSize(entries, sheet.Size));
                }
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
        if (!names.SequenceEqual(Sheets.Select(sheet => sheet.Name), StringComparer.Ordinal))
        {
            throw new InvalidDataException(
                "選取的 Excel 不是含有「三入、六入、九入」的中秋禮盒統計活頁簿，請另存新檔。");
        }

        var sharedStrings = ReadSharedStrings(archive, spreadsheet);
        foreach (var sheet in Sheets)
        {
            var worksheetEntry = archive.GetEntry(sheet.Path)
                ?? throw new InvalidDataException("選取的 Excel 缺少中秋禮盒統計工作表，請另存新檔。");
            using var worksheetInput = worksheetEntry.Open();
            var worksheet = XDocument.Load(worksheetInput);
            var firstRow = worksheet.Descendants(spreadsheet + "row")
                .FirstOrDefault(row => (string?)row.Attribute("r") == "1");
            var headers = firstRow?.Elements(spreadsheet + "c")
                .Select(cell => ReadCellText(cell, spreadsheet, sharedStrings))
                .ToArray() ?? [];
            var hiddenIdColumn = worksheet.Descendants(spreadsheet + "col").Any(column =>
                (string?)column.Attribute("min") == "6" &&
                (string?)column.Attribute("max") == "6" &&
                (string?)column.Attribute("hidden") == "1");
            if (!headers.SequenceEqual(Headers, StringComparer.Ordinal) || !hiddenIdColumn)
            {
                throw new InvalidDataException(
                    "選取的 Excel 不是由本程式建立的三種禮盒活頁簿，請另存新檔。");
            }
        }
    }

    private static void WriteWorksheet(
        ZipArchive archive,
        string path,
        IReadOnlyList<OrderChangeEntry> entries)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = XmlWriter.Create(stream, XmlSettings);
        writer.WriteStartDocument(true);
        writer.WriteStartElement("worksheet", SpreadsheetNamespace);
        WriteHiddenIdColumn(writer);
        writer.WriteStartElement("sheetData", SpreadsheetNamespace);
        WriteHeaderRow(writer);
        for (var index = 0; index < entries.Count; index++)
        {
            WriteDataRow(writer, index + 2, entries[index]);
        }

        writer.WriteEndElement();
        writer.WriteStartElement("autoFilter", SpreadsheetNamespace);
        writer.WriteAttributeString("ref", $"A1:E{Math.Max(1, entries.Count + 1)}");
        writer.WriteEndElement();
        writer.WriteEndElement();
        writer.WriteEndDocument();
    }

    private static void AppendWorksheet(
        ZipArchive archive,
        string path,
        IReadOnlyList<OrderChangeEntry> entries)
    {
        var worksheetEntry = archive.GetEntry(path)
            ?? throw new InvalidDataException("中秋禮盒統計工作表不完整，請另存新檔。");
        XDocument document;
        using (var input = worksheetEntry.Open())
        {
            document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        }

        XNamespace spreadsheet = SpreadsheetNamespace;
        var sheetData = document.Root?.Element(spreadsheet + "sheetData")
            ?? throw new InvalidDataException("中秋禮盒統計工作表格式不完整。");
        var sharedStrings = ReadSharedStrings(archive, spreadsheet);
        var knownIds = sheetData.Elements(spreadsheet + "row")
            .Skip(1)
            .Select(row => row.Elements(spreadsheet + "c")
                .FirstOrDefault(cell => ((string?)cell.Attribute("r"))?.StartsWith("F", StringComparison.OrdinalIgnoreCase) == true))
            .Where(cell => cell is not null)
            .Select(cell => ReadCellText(cell!, spreadsheet, sharedStrings))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var nextRow = sheetData.Elements(spreadsheet + "row")
            .Select(row => int.TryParse((string?)row.Attribute("r"), out var number) ? number : 0)
            .DefaultIfEmpty(1)
            .Max() + 1;
        foreach (var entry in entries)
        {
            if (knownIds.Add(entry.EntryId))
            {
                sheetData.Add(CreateDataRow(nextRow++, entry, spreadsheet));
            }
        }

        document.Root?.Element(spreadsheet + "autoFilter")
            ?.SetAttributeValue("ref", $"A1:E{Math.Max(1, nextRow - 1)}");
        worksheetEntry.Delete();
        var replacement = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var output = replacement.Open();
        document.Save(output, SaveOptions.DisableFormatting);
    }

    private static void WriteHiddenIdColumn(XmlWriter writer)
    {
        writer.WriteStartElement("cols", SpreadsheetNamespace);
        writer.WriteStartElement("col", SpreadsheetNamespace);
        writer.WriteAttributeString("min", "6");
        writer.WriteAttributeString("max", "6");
        writer.WriteAttributeString("hidden", "1");
        writer.WriteAttributeString("width", "1");
        writer.WriteAttributeString("customWidth", "1");
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteHeaderRow(XmlWriter writer)
    {
        writer.WriteStartElement("row", SpreadsheetNamespace);
        writer.WriteAttributeString("r", "1");
        for (var index = 0; index < Headers.Length; index++)
        {
            WriteTextCell(writer, $"{ColumnName(index + 1)}1", Headers[index]);
        }

        writer.WriteEndElement();
    }

    private static void WriteDataRow(XmlWriter writer, int rowNumber, OrderChangeEntry entry)
    {
        writer.WriteStartElement("row", SpreadsheetNamespace);
        writer.WriteAttributeString("r", rowNumber.ToString(CultureInfo.InvariantCulture));
        WriteTextCell(writer, $"A{rowNumber}", entry.ChannelName ?? entry.SourceName);
        WriteTextCell(writer, $"B{rowNumber}", entry.ExternalOrderNo);
        WriteTextCell(writer, $"C{rowNumber}", entry.DeliveryDate ?? string.Empty);
        WriteNumberCell(writer, $"D{rowNumber}", entry.QuantityChange);
        WriteTextCell(writer, $"E{rowNumber}", string.Empty);
        WriteTextCell(writer, $"F{rowNumber}", entry.EntryId);
        writer.WriteEndElement();
    }

    private static XElement CreateDataRow(int rowNumber, OrderChangeEntry entry, XNamespace spreadsheet) =>
        new(spreadsheet + "row",
            new XAttribute("r", rowNumber),
            CreateTextCell($"A{rowNumber}", entry.ChannelName ?? entry.SourceName, spreadsheet),
            CreateTextCell($"B{rowNumber}", entry.ExternalOrderNo, spreadsheet),
            CreateTextCell($"C{rowNumber}", entry.DeliveryDate ?? string.Empty, spreadsheet),
            new XElement(spreadsheet + "c",
                new XAttribute("r", $"D{rowNumber}"),
                new XElement(spreadsheet + "v", entry.QuantityChange.ToString(CultureInfo.InvariantCulture))),
            CreateTextCell($"E{rowNumber}", string.Empty, spreadsheet),
            CreateTextCell($"F{rowNumber}", entry.EntryId, spreadsheet));

    private static XElement CreateTextCell(string reference, string value, XNamespace spreadsheet) =>
        new(spreadsheet + "c",
            new XAttribute("r", reference),
            new XAttribute("t", "inlineStr"),
            new XElement(spreadsheet + "is",
                new XElement(spreadsheet + "t",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    RemoveInvalidXmlCharacters(value))));

    private static void WriteTextCell(XmlWriter writer, string reference, string value)
    {
        writer.WriteStartElement("c", SpreadsheetNamespace);
        writer.WriteAttributeString("r", reference);
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
        writer.WriteStartElement("v", SpreadsheetNamespace);
        writer.WriteString(value.ToString(CultureInfo.InvariantCulture));
        writer.WriteEndElement();
        writer.WriteEndElement();
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
            int.TryParse(cell.Element(spreadsheet + "v")?.Value, out var index) &&
            index >= 0 && index < sharedStrings.Count)
        {
            return sharedStrings[index];
        }

        return string.Concat(cell.Descendants(spreadsheet + "t").Select(text => text.Value));
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

    private static void WriteTextEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private const string ContentTypesXml = """
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
            <sheet name="三入" sheetId="1" r:id="rId1"/>
            <sheet name="六入" sheetId="2" r:id="rId2"/>
            <sheet name="九入" sheetId="3" r:id="rId3"/>
          </sheets>
        </workbook>
        """;

    private const string WorkbookRelationshipsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
        </Relationships>
        """;
}
