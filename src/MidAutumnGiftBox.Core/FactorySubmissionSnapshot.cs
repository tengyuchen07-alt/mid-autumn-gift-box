namespace MidAutumnGiftBox.Core;

public sealed record FactorySubmissionLine(
    string OrderDate,
    string DeliveryDate,
    string Location,
    string ProductName,
    decimal Quantity);

public sealed record FactorySubmissionSnapshot(
    DateTimeOffset SubmittedAt,
    string WorkbookPath,
    IReadOnlyList<FactorySubmissionLine> Lines);
