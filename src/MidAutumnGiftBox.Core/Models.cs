namespace MidAutumnGiftBox.Core;

public sealed record WmsSource(string Code, string Name, Uri BaseUri);

public sealed record ApiCredentials(string ApiId, string ApiKey)
{
    public bool IsComplete => !string.IsNullOrWhiteSpace(ApiId) && !string.IsNullOrWhiteSpace(ApiKey);
}

public static class WmsQueryPolicy
{
    private static readonly TimeZoneInfo TaipeiTimeZone =
        TimeZoneInfo.FindSystemTimeZoneById("Taipei Standard Time");

    public static DateOnly InitialOrderDate { get; } = new(2026, 7, 1);

    public static DateOnly GetTaipeiDate(DateTimeOffset instant) =>
        DateOnly.FromDateTime(TimeZoneInfo.ConvertTime(instant, TaipeiTimeZone).DateTime);
}

public sealed record SyncWindow(DateOnly FromDate, DateOnly ThroughDate, DateTimeOffset StartedAt);

public static class SyncWindowPolicy
{
    public static SyncWindow Create(DateTimeOffset? lastSuccessAt, DateTimeOffset startedAt)
    {
        _ = lastSuccessAt;
        var throughDate = WmsQueryPolicy.GetTaipeiDate(startedAt);
        return new SyncWindow(WmsQueryPolicy.InitialOrderDate, throughDate, startedAt);
    }
}

public sealed record PreviewResult(
    bool IsSuccess,
    string Message,
    IReadOnlyList<IReadOnlyDictionary<string, string?>> Rows,
    string SafeJson,
    PreviewMetadata Metadata);

public sealed record PreviewMetadata(
    string SourceName,
    string Endpoint,
    DateTimeOffset TestedAt,
    int HttpStatusCode,
    int PageCount,
    int RowCount,
    bool ResultOk,
    string ResultMessage);

public class WmsApiException : Exception
{
    public WmsApiException(string message, Exception? innerException = null, int? statusCode = null)
        : base(message, innerException)
    {
        StatusCode = statusCode;
    }

    public int? StatusCode { get; }
}
