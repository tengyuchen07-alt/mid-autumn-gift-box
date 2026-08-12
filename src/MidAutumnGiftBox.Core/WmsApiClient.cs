using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Collections.Concurrent;
using System.Security.Cryptography;

namespace MidAutumnGiftBox.Core;

public sealed class WmsApiClient
{
    private const int MaxFlavorBomLookupsPerPreview = 100;
    private readonly HttpClient _httpClient;
    private readonly Func<int, CancellationToken, Task> _retryDelay;
    private readonly ConcurrentDictionary<string, TokenCacheEntry> _tokenCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, SemaphoreSlim> _tokenLocks = new(StringComparer.OrdinalIgnoreCase);

    public WmsApiClient(
        HttpClient httpClient,
        Func<int, CancellationToken, Task>? retryDelay = null)
    {
        _httpClient = httpClient;
        _retryDelay = retryDelay ?? ((attempt, cancellationToken) =>
            Task.Delay(TimeSpan.FromMilliseconds(250 * Math.Pow(2, attempt - 1)), cancellationToken));
    }

    public async Task<PreviewResult> GetShopsAsync(
        WmsSource source,
        ApiCredentials credentials,
        CancellationToken cancellationToken = default)
    {
        return await GetPreviewAsync(
            source,
            credentials,
            "/api_v1/shop/shop_list.php",
            cancellationToken);
    }

    public async Task<PreviewResult> GetOrderByNumberAsync(
        WmsSource source,
        ApiCredentials credentials,
        string orderNumber,
        CancellationToken cancellationToken = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(orderNumber);
        var normalizedOrderNumber = orderNumber.Trim();
        var path = "/api_v1/order/order_query.php" +
                   $"?order_no={Uri.EscapeDataString(normalizedOrderNumber)}&nowpage=1&pagesize=50";
        return await GetPreviewAsync(
            source,
            credentials,
            path,
            cancellationToken,
            row => string.Equals(
                GetElementString(row, "order_no"),
                normalizedOrderNumber,
                StringComparison.OrdinalIgnoreCase));
    }

    public Task<PreviewResult> GetPendingOrdersSinceInitialDateAsync(
        WmsSource source,
        ApiCredentials credentials,
        DateOnly throughDate,
        CancellationToken cancellationToken = default) =>
        GetPendingOrdersAsync(
            source,
            credentials,
            WmsQueryPolicy.InitialOrderDate,
            throughDate,
            cancellationToken);

    public async Task<PreviewResult> GetPendingOrdersAsync(
        WmsSource source,
        ApiCredentials credentials,
        DateOnly fromDate,
        DateOnly throughDate,
        CancellationToken cancellationToken = default)
    {
        var isFlavor = FlavorProductPolicy.AppliesTo(source);
        if (fromDate < WmsQueryPolicy.InitialOrderDate)
        {
            throw new ArgumentException(
                $"查詢起日不可早於 {WmsQueryPolicy.InitialOrderDate:yyyy/MM/dd}。",
                nameof(fromDate));
        }

        if (throughDate < fromDate)
        {
            throw new ArgumentException("查詢結束日不可早於查詢起日。", nameof(throughDate));
        }

        bool IsRecordablePendingOrder(JsonElement row)
        {
            if (row.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            var isPending = false;
            foreach (var property in row.EnumerateObject())
            {
                if (property.Name.Equals("status_code", StringComparison.OrdinalIgnoreCase) &&
                    property.Value.ValueKind == JsonValueKind.String)
                {
                    isPending = string.Equals(property.Value.GetString(), "F", StringComparison.OrdinalIgnoreCase);
                }
            }

            var isExplicitFlavorNonPageOrder = isFlavor && string.Equals(
                GetElementString(row, "logistics_code")?.Trim(),
                "none",
                StringComparison.OrdinalIgnoreCase);
            return isPending && !isExplicitFlavorNonPageOrder && GiftBoxProductPolicy.ContainsRelevantProduct(row);
        }

        if (!credentials.IsComplete)
        {
            throw new ArgumentException("請輸入 API ID 與 API Key。", nameof(credentials));
        }

        var token = await GetTokenAsync(source, credentials, cancellationToken);
        var rows = new List<IReadOnlyDictionary<string, string?>>();
        var safePages = new List<string>();
        var page = 1;
        var maxPage = 1;
        var lastStatusCode = 0;
        var lastResultMessage = string.Empty;
        var flavorBomCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        do
        {
            var path = "/api_v1/order/order_query.php" +
                       $"?date_min={fromDate:yyyy-MM-dd}" +
                       $"&date_max={throughDate:yyyy-MM-dd}" +
                       $"&status=F&nowpage={page}&pagesize=500";
            var authorizedResult = await SendAuthorizedForJsonAsync(
                source,
                credentials,
                path,
                token,
                cancellationToken);
            token = authorizedResult.Token;
            var httpResult = authorizedResult.Response;
            lastStatusCode = httpResult.StatusCode;
            var json = httpResult.Body;

            var hasResult = PreviewSanitizer.TryReadResult(json, out var ok, out var message);
            lastResultMessage = string.IsNullOrWhiteSpace(message)
                ? string.Empty
                : SafeMessage(message, string.Empty, credentials.ApiId, credentials.ApiKey, token);
            if (hasResult && !ok)
            {
                var safeFailureMessage = SafeMessage(
                    message, "WMS 回覆訂單查詢失敗。", credentials.ApiId, credentials.ApiKey, token);
                return new PreviewResult(
                    false,
                    safeFailureMessage,
                    Array.Empty<IReadOnlyDictionary<string, string?>>(),
                    "{}",
                    new PreviewMetadata(source.Name, "/api_v1/order/order_query.php", DateTimeOffset.UtcNow,
                        httpResult.StatusCode, page, 0, false, safeFailureMessage));
            }

            var normalized = await NormalizeOrderPageAsync(
                source,
                credentials,
                json,
                token,
                isFlavor,
                flavorBomCache,
                cancellationToken);
            json = normalized.Json;
            token = normalized.Token;

            rows.AddRange(PreviewSanitizer.ProjectRows(json, IsRecordablePendingOrder));
            safePages.Add(PreviewSanitizer.Sanitize(json, IsRecordablePendingOrder));
            maxPage = Math.Clamp(PreviewSanitizer.FindInt(json, "maxpage") ?? 1, 1, 10_000);
            page++;
        }
        while (page <= maxPage);

        var safeJson = safePages.Count == 1
            ? safePages[0]
            : "[\n" + string.Join(",\n", safePages) + "\n]";
        return new PreviewResult(
            true,
            $"查詢成功，訂單成立日 {fromDate:yyyy/MM/dd} 至 {throughDate:yyyy/MM/dd}，共 {rows.Count} 筆包含蛋黃酥品項的待處理訂單。",
            rows,
            safeJson,
            new PreviewMetadata(
                source.Name,
                "/api_v1/order/order_query.php",
                DateTimeOffset.UtcNow,
                lastStatusCode,
                safePages.Count,
                rows.Count,
                true,
                string.IsNullOrWhiteSpace(lastResultMessage) ? "查詢成功。" : lastResultMessage));
    }

    private async Task<NormalizedFlavorPage> NormalizeOrderPageAsync(
        WmsSource source,
        ApiCredentials credentials,
        string json,
        string token,
        bool enableFlavorBom,
        IDictionary<string, string> bomCache,
        CancellationToken cancellationToken)
    {
        JsonNode? root;
        try
        {
            root = JsonNode.Parse(json);
        }
        catch (JsonException)
        {
            return new NormalizedFlavorPage(json, token);
        }

        if (root is null)
        {
            return new NormalizedFlavorPage(json, token);
        }

        var rows = SelectOrderRowsArray(root);
        if (rows is null)
        {
            return new NormalizedFlavorPage(root.ToJsonString(), token);
        }

        var currentToken = token;
        foreach (var order in rows.OfType<JsonObject>())
        {
            if (!string.Equals(GetString(order, "status_code"), "F", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (enableFlavorBom)
            {
                FlavorOperationalDiagnostics.Enrich(order);
            }

            var productsKey = FindKey(order, "products");
            if (productsKey is null)
            {
                continue;
            }

            var products = ParseArrayNode(order[productsKey]);
            if (products is null)
            {
                order[productsKey] = new JsonArray();
                continue;
            }

            order[productsKey] = products;
            foreach (var product in products.OfType<JsonObject>())
            {
                var itemsKey = FindKey(product, "items");
                var existingItems = itemsKey is null ? null : ParseArrayNode(product[itemsKey]);
                if (itemsKey is not null && existingItems is not null)
                {
                    product[itemsKey] = existingItems;
                }

                var isCombine = string.Equals(
                    GetString(product, "type"),
                    "combine",
                    StringComparison.OrdinalIgnoreCase);
                if (!enableFlavorBom || !isCombine || existingItems is { Count: > 0 })
                {
                    continue;
                }

                var sku = GetString(product, "sku");
                if (string.IsNullOrWhiteSpace(sku))
                {
                    continue;
                }

                if (!bomCache.TryGetValue(sku, out var cachedItems))
                {
                    if (bomCache.Count >= MaxFlavorBomLookupsPerPreview)
                    {
                        throw new WmsApiException(
                            $"Flavor 組合品 BOM 查詢超過單次上限 {MaxFlavorBomLookupsPerPreview} 個不同 SKU，" +
                            "為避免長時間等待，本次預覽已停止；請縮小訂單範圍或請 WMS 提供批次 BOM。");
                    }

                    var path = $"/api_v1/product/bom.php?sku={Uri.EscapeDataString(sku)}";
                    var bomResult = await SendAuthorizedForJsonAsync(
                        source,
                        credentials,
                        path,
                        currentToken,
                        cancellationToken);
                    currentToken = bomResult.Token;
                    if (PreviewSanitizer.TryReadResult(
                            bomResult.Response.Body,
                            out var bomOk,
                            out var bomMessage) && !bomOk)
                    {
                        throw new WmsApiException(SafeMessage(
                            bomMessage,
                            $"Flavor BOM 查詢失敗（SKU：{sku}）。",
                            credentials.ApiId,
                            credentials.ApiKey,
                            currentToken));
                    }

                    cachedItems = ExtractBomItems(bomResult.Response.Body).ToJsonString();
                    bomCache[sku] = cachedItems;
                }

                product[itemsKey ?? "items"] = JsonNode.Parse(cachedItems);
            }

            for (var index = products.Count - 1; index >= 0; index--)
            {
                if (products[index] is not JsonObject product ||
                    !GiftBoxProductPolicy.IsRelevantProduct(product))
                {
                    products.RemoveAt(index);
                }
            }

            OrderPreviewEnricher.Enrich(order, products);
        }

        return new NormalizedFlavorPage(root.ToJsonString(), currentToken);
    }

    private static JsonArray ExtractBomItems(string json)
    {
        JsonNode? root;
        try
        {
            root = JsonNode.Parse(json);
        }
        catch (JsonException)
        {
            return new JsonArray();
        }

        var rows = FindRowsArray(root);
        var result = new JsonArray();
        if (rows is null)
        {
            return result;
        }

        foreach (var item in rows.OfType<JsonObject>())
        {
            var mapped = new JsonObject();
            CopyScalar(item, mapped, "sku");
            CopyScalar(item, mapped, "item_no");
            CopyScalar(item, mapped, "type");
            CopyScalar(item, mapped, "name");
            CopyScalar(item, mapped, "spec");
            CopyScalar(item, mapped, "shipp_qty");
            var quantityKey = FindKey(item, "quantity");
            var qtyKey = FindKey(item, "qty");
            var quantity = quantityKey is null ? (qtyKey is null ? null : item[qtyKey]) : item[quantityKey];
            if (quantity is not null)
            {
                mapped["qty"] = quantity.DeepClone();
            }

            if (mapped.Count > 0)
            {
                result.Add(mapped);
            }
        }

        return result;
    }

    private static JsonArray? SelectOrderRowsArray(JsonNode root)
    {
        if (root is not JsonObject envelope)
        {
            return null;
        }

        var dataKey = FindKey(envelope, "data");
        if (dataKey is not null && envelope[dataKey] is JsonObject data)
        {
            var dataRowsKey = FindKey(data, "rows");
            if (dataRowsKey is not null && data[dataRowsKey] is JsonArray dataRows)
            {
                var resultKey = FindKey(envelope, "result");
                if (resultKey is not null && envelope[resultKey] is JsonObject result)
                {
                    var alternateRowsKey = FindKey(result, "rows");
                    if (alternateRowsKey is not null)
                    {
                        result.Remove(alternateRowsKey);
                    }
                }

                return dataRows;
            }
        }

        var resultObjectKey = FindKey(envelope, "result");
        if (resultObjectKey is not null && envelope[resultObjectKey] is JsonObject resultObject)
        {
            var resultRowsKey = FindKey(resultObject, "rows");
            if (resultRowsKey is not null && resultObject[resultRowsKey] is JsonArray resultRows)
            {
                return resultRows;
            }
        }

        var directRowsKey = FindKey(envelope, "rows");
        return directRowsKey is not null ? envelope[directRowsKey] as JsonArray : null;
    }

    private static JsonArray? FindRowsArray(JsonNode? node)
    {
        if (node is JsonObject obj)
        {
            foreach (var property in obj)
            {
                if (property.Key.Equals("rows", StringComparison.OrdinalIgnoreCase) &&
                    property.Value is JsonArray rows)
                {
                    return rows;
                }

                var nested = FindRowsArray(property.Value);
                if (nested is not null)
                {
                    return nested;
                }
            }
        }
        else if (node is JsonArray array)
        {
            foreach (var item in array)
            {
                var nested = FindRowsArray(item);
                if (nested is not null)
                {
                    return nested;
                }
            }
        }

        return null;
    }

    private static JsonArray? ParseArrayNode(JsonNode? node)
    {
        if (node is JsonArray array)
        {
            return (JsonArray)array.DeepClone();
        }

        if (node is not JsonValue value || !value.TryGetValue<string>(out var text) || string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        try
        {
            return JsonNode.Parse(text) as JsonArray;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    private static string? FindKey(JsonObject obj, string name) =>
        obj.Select(property => property.Key)
            .FirstOrDefault(key => key.Equals(name, StringComparison.OrdinalIgnoreCase));

    private static string? GetElementString(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        foreach (var property in element.EnumerateObject())
        {
            if (property.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                return property.Value.ValueKind == JsonValueKind.String
                    ? property.Value.GetString()
                    : property.Value.ToString();
            }
        }

        return null;
    }

    private static string? GetString(JsonObject obj, string name)
    {
        var key = FindKey(obj, name);
        if (key is null || obj[key] is not JsonValue value)
        {
            return null;
        }

        return value.TryGetValue<string>(out var text) ? text : value.ToJsonString().Trim('"');
    }

    private static void CopyScalar(JsonObject source, JsonObject destination, string name)
    {
        var key = FindKey(source, name);
        if (key is not null && source[key] is JsonValue value)
        {
            destination[name] = value.DeepClone();
        }
    }

    private async Task<PreviewResult> GetPreviewAsync(
        WmsSource source,
        ApiCredentials credentials,
        string path,
        CancellationToken cancellationToken,
        Func<JsonElement, bool>? rowFilter = null)
    {
        if (!credentials.IsComplete)
        {
            throw new ArgumentException("請輸入 API ID 與 API Key。", nameof(credentials));
        }

        var token = await GetTokenAsync(source, credentials, cancellationToken);
        var authorizedResult = await SendAuthorizedForJsonAsync(
            source,
            credentials,
            path,
            token,
            cancellationToken);
        token = authorizedResult.Token;
        var httpResult = authorizedResult.Response;
        var json = httpResult.Body;
        var hasResult = PreviewSanitizer.TryReadResult(json, out var ok, out var message);
        var safeResultMessage = string.IsNullOrWhiteSpace(message)
            ? string.Empty
            : SafeMessage(message, string.Empty, credentials.ApiId, credentials.ApiKey, token);
        if (hasResult && !ok)
        {
            var safeFailureMessage = SafeMessage(
                message, "WMS 回覆查詢失敗。", credentials.ApiId, credentials.ApiKey, token);
            return new PreviewResult(false, safeFailureMessage,
                Array.Empty<IReadOnlyDictionary<string, string?>>(), "{}",
                new PreviewMetadata(source.Name, new Uri(source.BaseUri, path).AbsolutePath,
                    DateTimeOffset.UtcNow, httpResult.StatusCode, 1, 0, false, safeFailureMessage));
        }

        var projectedRows = PreviewSanitizer.ProjectRows(json, rowFilter);
        return new PreviewResult(
            true,
            string.IsNullOrWhiteSpace(message)
                ? "查詢成功。"
                : SafeMessage(message, "查詢成功。", credentials.ApiId, credentials.ApiKey, token),
            projectedRows,
            PreviewSanitizer.Sanitize(json, rowFilter),
            new PreviewMetadata(
                source.Name,
                new Uri(source.BaseUri, path).AbsolutePath,
                DateTimeOffset.UtcNow,
                httpResult.StatusCode,
                1,
                projectedRows.Count,
                true,
                string.IsNullOrWhiteSpace(safeResultMessage) ? "查詢成功。" : safeResultMessage));
    }

    private async Task<string> GetTokenAsync(
        WmsSource source,
        ApiCredentials credentials,
        CancellationToken cancellationToken)
    {
        var cacheKey = BuildTokenCacheKey(source);
        var credentialFingerprint = Convert.ToHexString(
            SHA256.HashData(Encoding.UTF8.GetBytes($"{credentials.ApiId}\0{credentials.ApiKey}")));
        if (TryGetCachedToken(cacheKey, credentialFingerprint, out var cachedToken))
        {
            return cachedToken;
        }

        var tokenLock = _tokenLocks.GetOrAdd(cacheKey, _ => new SemaphoreSlim(1, 1));
        await tokenLock.WaitAsync(cancellationToken);
        try
        {
            if (TryGetCachedToken(cacheKey, credentialFingerprint, out cachedToken))
            {
                return cachedToken;
            }

        using var request = new HttpRequestMessage(
            HttpMethod.Get,
            new Uri(source.BaseUri, "/api_v1/token/authorize.php"));
        var basicValue = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{credentials.ApiId}:{credentials.ApiKey}"));
        request.Headers.Authorization = new AuthenticationHeaderValue("Basic", basicValue);

        var json = (await SendForJsonAsync(request, cancellationToken)).Body;
        if (PreviewSanitizer.TryReadResult(json, out var ok, out var message) && !ok)
        {
            throw new WmsApiException(SafeMessage(
                message, "API ID 或 API Key 驗證失敗。", credentials.ApiId, credentials.ApiKey));
        }

        var token = PreviewSanitizer.FindString(json, "access_token");
        if (string.IsNullOrWhiteSpace(token))
        {
            throw new WmsApiException("WMS 未回傳 access_token，請確認 API ID、API Key 與端點設定。");
        }

        var expiresInSeconds = Math.Clamp(PreviewSanitizer.FindInt(json, "expires_in") ?? 3_600, 60, 86_400);
        _tokenCache[cacheKey] = new TokenCacheEntry(
            credentialFingerprint,
            token,
            DateTimeOffset.UtcNow.AddSeconds(expiresInSeconds));
        return token;
        }
        finally
        {
            tokenLock.Release();
        }
    }

    private async Task<HttpJsonResponse> SendForJsonAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken)
    {
        for (var attempt = 1; attempt <= 3; attempt++)
        {
            try
            {
                using var retryRequest = CloneRequest(request);
                using var response = await _httpClient.SendAsync(
                    retryRequest,
                    HttpCompletionOption.ResponseHeadersRead,
                    cancellationToken);
                var body = await response.Content.ReadAsStringAsync(cancellationToken);
                if (!response.IsSuccessStatusCode)
                {
                    if (attempt < 3 && IsTransient(response.StatusCode))
                    {
                        await _retryDelay(attempt, cancellationToken);
                        continue;
                    }

                    throw new WmsApiException(
                        $"WMS 連線失敗（HTTP {(int)response.StatusCode}）。",
                        statusCode: (int)response.StatusCode);
                }

                if (string.IsNullOrWhiteSpace(body))
                {
                    throw new WmsApiException("WMS 回傳空白內容。");
                }

                try
                {
                    using var _ = JsonDocument.Parse(body);
                }
                catch (JsonException exception)
                {
                    throw new WmsApiException("WMS 回傳格式不是有效 JSON，請確認 API 端點或聯絡系統供應商。", exception);
                }

                return new HttpJsonResponse(body, (int)response.StatusCode);
            }
            catch (HttpRequestException exception) when (attempt < 3)
            {
                await _retryDelay(attempt, cancellationToken);
                _ = exception;
            }
            catch (TaskCanceledException) when (!cancellationToken.IsCancellationRequested && attempt < 3)
            {
                await _retryDelay(attempt, cancellationToken);
            }
            catch (WmsApiException)
            {
                throw;
            }
            catch (TaskCanceledException exception) when (!cancellationToken.IsCancellationRequested)
            {
                throw new WmsApiException("WMS 連線逾時，請稍後再試。", exception);
            }
            catch (HttpRequestException exception)
            {
                throw new WmsApiException("無法連接 WMS，請確認公司網路、網址或 TLS 設定。", exception);
            }
        }

        throw new WmsApiException("WMS 連線失敗，已完成三次嘗試。");
    }

    private bool TryGetCachedToken(string cacheKey, string credentialFingerprint, out string token)
    {
        if (_tokenCache.TryGetValue(cacheKey, out var entry) &&
            entry.CredentialFingerprint == credentialFingerprint &&
            entry.ExpiresAt > DateTimeOffset.UtcNow.AddMinutes(2))
        {
            token = entry.Token;
            return true;
        }

        token = string.Empty;
        return false;
    }

    private async Task<AuthorizedHttpResult> SendAuthorizedForJsonAsync(
        WmsSource source,
        ApiCredentials credentials,
        string path,
        string token,
        CancellationToken cancellationToken)
    {
        var currentToken = token;
        for (var authorizationAttempt = 1; authorizationAttempt <= 2; authorizationAttempt++)
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, new Uri(source.BaseUri, path));
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", currentToken);
            try
            {
                var response = await SendForJsonAsync(request, cancellationToken);
                return new AuthorizedHttpResult(response, currentToken);
            }
            catch (WmsApiException exception) when (
                authorizationAttempt == 1 && exception.StatusCode is 401 or 403)
            {
                InvalidateCachedToken(source, currentToken);
                currentToken = await GetTokenAsync(source, credentials, cancellationToken);
            }
        }

        throw new WmsApiException("WMS 授權失效，重新取得 Token 後仍無法完成查詢。");
    }

    private void InvalidateCachedToken(WmsSource source, string token)
    {
        var cacheKey = BuildTokenCacheKey(source);
        if (_tokenCache.TryGetValue(cacheKey, out var entry) && entry.Token == token)
        {
            _tokenCache.TryRemove(cacheKey, out _);
        }
    }

    private static string BuildTokenCacheKey(WmsSource source) =>
        $"{source.Code}|{source.BaseUri.GetLeftPart(UriPartial.Authority)}";

    private static HttpRequestMessage CloneRequest(HttpRequestMessage request)
    {
        var clone = new HttpRequestMessage(request.Method, request.RequestUri)
        {
            Version = request.Version,
            VersionPolicy = request.VersionPolicy
        };
        foreach (var header in request.Headers)
        {
            clone.Headers.TryAddWithoutValidation(header.Key, header.Value);
        }

        return clone;
    }

    private static bool IsTransient(System.Net.HttpStatusCode statusCode)
    {
        var value = (int)statusCode;
        return statusCode is System.Net.HttpStatusCode.RequestTimeout or
               System.Net.HttpStatusCode.TooManyRequests || value >= 500;
    }

    private static string SafeMessage(string? message, string fallback, params string?[] secrets)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return fallback;
        }

        var normalized = message.Replace('\r', ' ').Replace('\n', ' ').Trim();
        if (secrets.Length >= 2 &&
            !string.IsNullOrWhiteSpace(secrets[0]) &&
            !string.IsNullOrWhiteSpace(secrets[1]))
        {
            var basicCredential = Convert.ToBase64String(
                Encoding.UTF8.GetBytes($"{secrets[0]}:{secrets[1]}"));
            normalized = normalized.Replace(basicCredential, "[已遮罩]", StringComparison.Ordinal);
        }

        foreach (var secret in secrets)
        {
            if (!string.IsNullOrEmpty(secret))
            {
                normalized = normalized.Replace(secret, "[已遮罩]", StringComparison.Ordinal);
            }
        }

        return normalized.Length <= 300 ? normalized : normalized[..300];
    }

    private sealed record TokenCacheEntry(
        string CredentialFingerprint,
        string Token,
        DateTimeOffset ExpiresAt);

    private sealed record HttpJsonResponse(string Body, int StatusCode);

    private sealed record AuthorizedHttpResult(HttpJsonResponse Response, string Token);

    private sealed record NormalizedFlavorPage(string Json, string Token);
}
