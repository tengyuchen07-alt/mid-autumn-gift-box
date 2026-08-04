using System.Net;
using System.Net.Http.Headers;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using MidAutumnGiftBox.Core;

var tests = new (string Name, Func<Task> Run)[]
{
    ("店舖預覽會取得 Token、使用 Bearer 並移除敏感資料", ShopPreviewUsesBearerAndRemovesSensitiveData),
    ("訂單預覽依成立日範圍取得蛋黃酥資料且不以指定到貨日排除", PendingOrderPreviewUsesOrderDateRange),
    ("睿驛訂單以 source_key 顯示通路並從商品 spec 推導出貨星期一", ReyiChannelAndShippingWindowUseVerifiedFields),
    ("模糊或衝突商品日期不會產生指定出貨日", ShippingDerivationRejectsFuzzyAndConflictingText),
    ("Flavor 預覽保留指定 SKU 或蛋黃酥商品並展開文字商品與 BOM", FlavorPreviewFiltersProductsAndExpandsBom),
    ("Flavor BOM 業務失敗會顯示安全錯誤而不是空明細", FlavorBomBusinessFailureIsReported),
    ("Flavor BOM 超過單次上限會停止並顯示完整性錯誤", FlavorBomLookupLimitStopsIncompletePreview),
    ("預覽可以匯出含訂單與商品明細的 Excel", PreviewCanBeExportedToExcel),
    ("兩網站憑證彼此獨立且畫面狀態不回顯 Key", CredentialsAreIsolatedAndKeyIsNeverDisplayed),
    ("WMS 錯誤訊息不會回顯 API Key", ErrorMessageNeverEchoesApiKey),
    ("成功回應的訊息也不會洩漏 Key 或 Token", SuccessMessageNeverLeaksSecrets),
    ("WMS 回傳非 JSON 時會轉為可理解錯誤", InvalidJsonBecomesSafeWmsError),
    ("暫時性伺服器錯誤最多重試後可以成功", TransientServerErrorsAreRetried),
    ("有效 Token 會快取並供後續查詢使用", ValidTokenIsCached),
    ("Bearer Token 被撤銷時會重新授權一次", RevokedTokenIsRefreshedOnce),
    ("API ID 被修改時不可默默使用舊 Key", EditedApiIdCannotReuseSavedKey),
    ("同步失敗會保留該來源先前的最後成功時間", SyncFailurePreservesLastSuccessfulRun),
    ("同步查詢日期固定使用台北時區", SyncQueryDateUsesTaipeiTimeZone),
    ("來源訂單快照重跑會更新商品且不影響另一來源", OrderSnapshotReplacesOneSourceWithoutDuplicates),
};

var failed = 0;
foreach (var test in tests)
{
    try
    {
        await test.Run();
        Console.WriteLine($"PASS {test.Name}");
    }
    catch (Exception exception)
    {
        failed++;
        Console.Error.WriteLine($"FAIL {test.Name}");
        Console.Error.WriteLine(exception.Message);
    }
}

Console.WriteLine($"Tests: {tests.Length}, Passed: {tests.Length - failed}, Failed: {failed}");
return failed == 0 ? 0 : 1;

static async Task ShopPreviewUsesBearerAndRemovesSensitiveData()
{
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath == "/api_v1/token/authorize.php")
        {
            Equal("Basic", request.Headers.Authorization?.Scheme, "Token request must use Basic auth.");
            Equal(
                Convert.ToBase64String(Encoding.UTF8.GetBytes("api-user:api-secret")),
                request.Headers.Authorization?.Parameter,
                "Token request credentials are incorrect.");

            return Json("""
                {"result":{"ok":true,"access_token":"temporary-token","token_type":"bearer","expires_in":3600}}
                """);
        }

        Equal("Bearer", request.Headers.Authorization?.Scheme, "Shop request must use Bearer auth.");
        Equal("temporary-token", request.Headers.Authorization?.Parameter, "Bearer token is incorrect.");
        return Json("""
            {
              "result": {
                "ok": true,
                "rows": [
                  {"shop_id":"momo","shop_name":"MOMO","shop_type":"platform","receiver_name":"王小明","phone":"0912345678"}
                ]
              },
              "access_token":"must-not-leak"
            }
            """);
    });

    var client = new WmsApiClient(new HttpClient(handler));
    var result = await client.GetShopsAsync(
        new WmsSource("site1", "網站1", new Uri("https://notflavor.example")),
        new ApiCredentials("api-user", "api-secret"));

    True(result.IsSuccess, "Shop preview should succeed.");
    Equal(1, result.Rows.Count, "Shop row count is incorrect.");
    Equal("momo", result.Rows[0]["shop_id"], "shop_id is incorrect.");
    Equal("網站1", result.Metadata.SourceName, "Preview source metadata is incorrect.");
    Equal("/api_v1/shop/shop_list.php", result.Metadata.Endpoint, "Preview endpoint metadata is incorrect.");
    Equal(200, result.Metadata.HttpStatusCode, "Preview HTTP metadata is incorrect.");
    Equal(1, result.Metadata.PageCount, "Shop preview page metadata is incorrect.");
    True(result.Metadata.ResultOk, "Preview result.ok metadata is incorrect.");
    Contains("MOMO", result.SafeJson, "Safe JSON should keep shop_name.");
    DoesNotContain("王小明", result.SafeJson, "Safe JSON leaked receiver_name.");
    DoesNotContain("0912345678", result.SafeJson, "Safe JSON leaked phone.");
    DoesNotContain("temporary-token", result.SafeJson, "Safe JSON leaked bearer token.");
    DoesNotContain("must-not-leak", result.SafeJson, "Safe JSON leaked access_token.");
}

static async Task PendingOrderPreviewUsesOrderDateRange()
{
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath == "/api_v1/token/authorize.php")
        {
            return Json("{" + "\"result\":{\"ok\":true,\"access_token\":\"token\"}}" );
        }

        Contains("status=F", request.RequestUri.Query, "Pending order status was not sent.");
        Contains("pagesize=500", request.RequestUri.Query, "Order page size was not sent.");
        Contains("date_min=2026-07-01", request.RequestUri.Query, "Initial order date was not sent.");
        Contains("date_max=2026-08-20", request.RequestUri.Query, "Current execution date was not sent.");
        DoesNotContain("arrival_date=", request.RequestUri.Query, "arrival_date is not a documented query parameter.");

        if (request.RequestUri.Query.Contains("nowpage=2", StringComparison.Ordinal))
        {
            return Json("""
                {
                  "result": {
                    "ok": true,
                    "maxpage": 2,
                    "rows": [{
                      "order_no":"A003",
                      "order_date":"2026/08/10 09:00:00",
                      "arrival_date":"2026/08/13",
                      "source":"shopline",
                      "status_code":"F",
                      "products":[{"sku":"MOON-03","name":"第二頁蛋黃酥禮盒","qty":3}]
                    }]
                  }
                }
                """);
        }

        Contains("nowpage=1", request.RequestUri.Query, "First order page was not requested.");
        return Json("""
            {
              "result": {
                "ok": true,
                "maxpage": 2,
                "rows": [
                  {
                    "order_no":"A001",
                    "order_date":"2026/08/12 10:20:00",
                    "arrival_date":"2026/08/13",
                    "source":"momo",
                    "status_code":"F",
                    "total_price":1200,
                    "receiver_name":"不應顯示",
                    "address":"台北市不應顯示",
                    "products":[{"sku":"MOON-01","name":"中秋蛋黃酥禮盒","qty":2,"shipp_qty":0,"phone":"0900000000"}]
                  },
                  {
                    "order_no":"A002",
                    "order_date":"2026/08/12 11:00:00",
                    "arrival_date":"2026/08/14",
                    "source":"momo",
                    "status_code":"F",
                    "products":[{"sku":"MOON-02","name":"其他日期蛋黃酥禮盒","qty":1}]
                  },
                  {
                    "order_no":"A004",
                    "order_date":"2026/08/12 12:00:00",
                    "arrival_date":"2026/08/13",
                    "source":"momo",
                    "status_code":"W",
                    "products":[{"sku":"MOON-04","name":"非待處理蛋黃酥禮盒","qty":1}]
                  },
                  {
                    "order_no":"A005",
                    "order_date":"2026/08/12 13:00:00",
                    "arrival_date":null,
                    "source":"momo",
                    "status_code":"F",
                    "products":[{"sku":"MOON-05","name":"缺少到貨日蛋黃酥禮盒","qty":1}]
                  },
                  {
                    "order_no":"A006",
                    "order_date":"2026/08/12 14:00:00",
                    "arrival_date":"2026/08/15",
                    "source":"momo",
                    "status_code":"F",
                    "products":[{"sku":"OTHER-01","name":"一般芝麻粉","qty":1}]
                  },
                  {
                    "order_no":"A007",
                    "order_date":"2026/08/12 15:00:00",
                    "arrival_date":null,
                    "source":"momo",
                    "status_code":"F",
                    "products":[{"sku":"SET-01","name":"中秋組合","qty":1,"items":[{"sku":"EGG-01","name":"黑芝麻蛋黃酥","qty":3}]}]
                  }
                ]
              }
            }
            """);
    });

    var client = new WmsApiClient(new HttpClient(handler));
    var result = await client.GetPendingOrdersSinceInitialDateAsync(
        new WmsSource("site1", "網站1", new Uri("https://example.test")),
        new ApiCredentials("id", "key"),
        new DateOnly(2026, 8, 20));

    True(result.IsSuccess, "Order preview should succeed.");
    Equal(5, result.Rows.Count, "All pending orders containing egg-yolk pastry products should be visible.");
    Equal("A001", result.Rows[0]["order_no"], "order_no is incorrect.");
    Equal(2, result.Metadata.PageCount, "Order preview page metadata is incorrect.");
    Equal("2026/08/13", result.Rows[0]["arrival_date"], "arrival_date is incorrect.");
    Contains("MOON-01", result.SafeJson, "Safe JSON should keep product sku.");
    Contains("A003", result.SafeJson, "Safe JSON should include matching rows from later pages.");
    Contains("A002", result.SafeJson, "Safe JSON should retain other valid arrival dates.");
    DoesNotContain("A004", result.SafeJson, "Safe JSON included a non-pending order.");
    Contains("A005", result.SafeJson, "An egg-yolk pastry order without arrival_date should remain visible.");
    Contains("A007", result.SafeJson, "An order with an egg-yolk pastry nested item should remain visible.");
    DoesNotContain("A006", result.SafeJson, "An unrelated product order should be excluded.");
    DoesNotContain("不應顯示", result.SafeJson, "Safe JSON leaked receiver data.");
    DoesNotContain("台北市", result.SafeJson, "Safe JSON leaked address.");
    DoesNotContain("0900000000", result.SafeJson, "Safe JSON leaked phone.");
}

static async Task ReyiChannelAndShippingWindowUseVerifiedFields()
{
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath == "/api_v1/token/authorize.php")
        {
            return Json("{\"result\":{\"ok\":true,\"access_token\":\"token\"}}");
        }

        return Json("""
            {
              "result":{"ok":true},
              "data":{"maxpage":1,"rows":[{
                "order_no":"260804EXFYR5FK",
                "order_date":"2026/08/04 11:36:25",
                "arrival_date":null,
                "source":"芝初 SesaOle 芝麻食品 專業製造商",
                "source_key":"shopee",
                "status_code":"F",
                "products":[{
                  "shop_id":"56362838419",
                  "sku":null,
                  "type":"shop",
                  "name":"SesaOle【芝初】預購_ 2026 黑芝麻Q潤蛋黃酥大禮盒(內含9入蛋黃酥+芝麻粉7g/包 x3)_8月中起出貨",
                  "spec":"8/22-8/28當週出貨",
                  "qty":1,
                  "items":[{
                    "sku":"4710964232565",
                    "item_no":"1902247",
                    "type":"warehouse",
                    "name":"芝初黑芝麻Q潤蛋黃酥9入禮盒",
                    "qty":1
                  }]
                }]
              }]}
            }
            """);
    });
    var client = new WmsApiClient(new HttpClient(handler));
    var result = await client.GetPendingOrdersSinceInitialDateAsync(
        new WmsSource("site1", "睿驛", new Uri("https://example.test")),
        new ApiCredentials("id", "key"),
        new DateOnly(2026, 8, 4));

    Equal(1, result.Rows.Count, "The verified Reyi order should remain visible.");
    var row = result.Rows[0];
    Equal("蝦皮賣場", row["source"], "Order source should use the source_key channel mapping.");
    Equal("芝初 SesaOle 芝麻食品 專業製造商", row["shop_name"], "Original source should remain as shop name.");
    Equal("shopee", row["source_key"], "Original source_key should remain available.");
    Equal(null, row["arrival_date"], "Derived shipping date must not overwrite arrival_date.");
    Equal("2026-08-22", row["ship_window_start"], "Shipping-window start is incorrect.");
    Equal("2026-08-28", row["ship_window_end"], "Shipping-window end is incorrect.");
    Equal("2026-08-24", row["derived_shipping_date"], "The Monday inside the shipping window is incorrect.");
    Equal("product.spec", row["shipping_date_source"], "Shipping-date source should identify product spec.");
    Contains("8/22-8/28當週出貨", row["products"] ?? string.Empty, "Product spec should remain visible.");
    Contains("蝦皮賣場", result.SafeJson, "Safe JSON should contain the channel display name.");
    Contains("2026-08-24", result.SafeJson, "Safe JSON should contain the derived shipping date.");
}

static async Task ShippingDerivationRejectsFuzzyAndConflictingText()
{
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath == "/api_v1/token/authorize.php")
        {
            return Json("{\"result\":{\"ok\":true,\"access_token\":\"token\"}}");
        }

        return Json("""
            {"result":{"ok":true},"data":{"maxpage":1,"rows":[
              {
                "order_no":"FUZZY",
                "order_date":"2026/08/04 09:00:00",
                "source":"原賣場",
                "source_key":"unknown-channel",
                "status_code":"F",
                "products":[{"name":"2026 蛋黃酥禮盒_8月中起出貨","spec":"","qty":1}]
              },
              {
                "order_no":"CONFLICT",
                "order_date":"2026/08/04 10:00:00",
                "status_code":"F",
                "products":[
                  {"name":"2026 蛋黃酥禮盒A","spec":"8/22-8/28當週出貨","qty":1},
                  {"name":"2026 蛋黃酥禮盒B","spec":"8/29-9/4當週出貨","qty":1}
                ]
              },
              {
                "order_no":"SPEC-AMBIGUOUS",
                "order_date":"2026/08/04 11:00:00",
                "status_code":"F",
                "derived_shipping_date":"2099-01-01",
                "ship_window_start":"2099-01-01",
                "products":[{
                  "name":"2026 蛋黃酥禮盒_8/22-8/28當週出貨",
                  "spec":"8/1-8/31當月出貨",
                  "qty":1
                }]
              },
              {
                "order_no":"MULTI-RANGE",
                "order_date":"2026/08/04 12:00:00",
                "status_code":"F",
                "products":[{
                  "name":"2026 蛋黃酥禮盒",
                  "spec":"8/1-8/31或8/29-9/4出貨",
                  "qty":1
                }]
              },
              {
                "order_no":"MALFORMED",
                "order_date":"2026/08/04 13:00:00",
                "status_code":"F",
                "products":[{
                  "name":"2026 蛋黃酥禮盒_8/22-8/28當週出貨",
                  "spec":"8/22-",
                  "qty":1
                }]
              },
              {
                "order_no":"CROSS-YEAR",
                "order_date":"2026/12/20 09:00:00",
                "status_code":"F",
                "products":[{
                  "name":"2026 蛋黃酥跨年禮盒",
                  "spec":"12/29-1/4當週出貨",
                  "qty":1
                }]
              }
            ]}}
            """);
    });
    var client = new WmsApiClient(new HttpClient(handler));
    var result = await client.GetPendingOrdersSinceInitialDateAsync(
        new WmsSource("site1", "睿驛", new Uri("https://example.test")),
        new ApiCredentials("id", "key"),
        new DateOnly(2026, 8, 4));

    Equal(6, result.Rows.Count, "All egg-yolk pastry orders should remain visible.");
    var fuzzy = result.Rows.Single(row => row["order_no"] == "FUZZY");
    Equal("原賣場", fuzzy["source"], "Unknown source_key should preserve the original source.");
    True(!fuzzy.ContainsKey("derived_shipping_date"), "Fuzzy month text must not derive an exact date.");
    var conflict = result.Rows.Single(row => row["order_no"] == "CONFLICT");
    Equal("conflict", conflict["shipping_date_status"], "Conflicting windows should be marked for review.");
    True(!conflict.ContainsKey("derived_shipping_date"), "Conflicting windows must not choose one date.");
    var ambiguous = result.Rows.Single(row => row["order_no"] == "SPEC-AMBIGUOUS");
    Equal("needs_review", ambiguous["shipping_date_status"], "An explicit non-unique spec should require review.");
    True(!ambiguous.ContainsKey("derived_shipping_date"), "Invalid spec must not fall back to name or retain API-supplied derived data.");
    True(!ambiguous.ContainsKey("ship_window_start"), "Application-owned window fields must be cleared before derivation.");
    var multiple = result.Rows.Single(row => row["order_no"] == "MULTI-RANGE");
    Equal("conflict", multiple["shipping_date_status"], "Multiple windows in one spec should be marked conflict.");
    True(!multiple.ContainsKey("derived_shipping_date"), "Multiple windows must not choose the first date.");
    var malformed = result.Rows.Single(row => row["order_no"] == "MALFORMED");
    Equal("needs_review", malformed["shipping_date_status"], "Malformed explicit spec should require review.");
    True(!malformed.ContainsKey("derived_shipping_date"), "Malformed spec must not fall back to a valid range in name.");
    var crossYear = result.Rows.Single(row => row["order_no"] == "CROSS-YEAR");
    Equal("2026-12-29", crossYear["ship_window_start"], "Cross-year start is incorrect.");
    Equal("2027-01-04", crossYear["ship_window_end"], "Cross-year end is incorrect.");
    Equal("2027-01-04", crossYear["derived_shipping_date"], "Cross-year Monday is incorrect.");
}

static async Task FlavorPreviewFiltersProductsAndExpandsBom()
{
    const string productsText = """
        [
          {
            "sku":"4710964232411",
            "type":"combine",
            "name":"目標禮盒一",
            "qty":2,
            "shipp_qty":1,
            "phone":"不應顯示"
          },
          {
            "sku":"4710964232565",
            "type":"combine",
            "name":"目標禮盒二",
            "qty":1,
            "items":[
              {"sku":"MAIN-001","item_no":"M001","name":"主商品一","qty":3,"shipp_qty":2,"exp":"不應顯示"}
            ]
          },
          {"sku":"4710964232435","type":"shop","name":"目標禮盒三","qty":4},
          {"sku":"EGG-EXTRA","type":"shop","name":"額外蛋黃酥禮盒","qty":5},
          {"sku":"BUNDLE-EGG","type":"combine","name":"中秋組合","qty":1},
          {"sku":"OTHER-999","type":"shop","name":"其他商品","qty":99}
        ]
        """;

    var bomRequests = 0;
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath == "/api_v1/token/authorize.php")
        {
            return Json("{\"result\":{\"ok\":true,\"access_token\":\"token\"}}");
        }

        if (request.RequestUri.AbsolutePath == "/api_v1/product/bom.php")
        {
            bomRequests++;
            if (request.RequestUri.Query.Contains("sku=BUNDLE-EGG", StringComparison.Ordinal))
            {
                return Json("""
                    {"result":{"ok":true},"data":{"rows":[
                      {"sku":"4710964232435","item_no":"1902226","quantity":3}
                    ]}}
                    """);
            }

            Contains("sku=4710964232411", request.RequestUri.Query, "BOM should be queried by the target parent SKU.");
            return Json("""
                {
                  "result":{"ok":true},
                  "data":{"rows":[
                    {"sku":"MAIN-BOM","item_no":"MB01","quantity":2,"cost_price":123}
                  ]}
                }
                """);
        }

        var body = JsonSerializer.Serialize(new
        {
            result = new
            {
                ok = true,
                rows = new object[]
                {
                    new Dictionary<string, object?>
                    {
                        ["order_no"] = "DECOY-ROWS",
                        ["status_code"] = "F",
                        ["products"] = """
                            [{"sku":"DECOY-COMBINE","type":"combine","name":"蛋黃酥假資料","qty":1}]
                            """
                    }
                }
            },
            data = new
            {
                maxpage = 1,
                rows = new object[]
                {
                    new Dictionary<string, object?>
                    {
                        ["order_no"] = "F001",
                        ["order_date"] = "2026/08/04 09:00:00",
                        ["arrival_date"] = "2026/08/12",
                        ["source"] = "flavor",
                        ["status_code"] = "F",
                        ["products"] = productsText
                    },
                    new Dictionary<string, object?>
                    {
                        ["order_no"] = "F-NON-PENDING",
                        ["arrival_date"] = null,
                        ["status_code"] = "W",
                        ["products"] = """
                            [{"sku":"SHOULD-NOT-BOM","type":"combine","name":"非待處理蛋黃酥組合","qty":1}]
                            """
                    }
                }
            }
        });
        return Json(body);
    });

    var client = new WmsApiClient(new HttpClient(handler));
    var result = await client.GetPendingOrdersSinceInitialDateAsync(
        new WmsSource("site2", "Flavor", new Uri("https://example.test")),
        new ApiCredentials("id", "key"),
        new DateOnly(2026, 8, 4));

    Equal(1, result.Rows.Count, "Flavor order containing target products should remain visible.");
    Equal(2, bomRequests, "Each combine product without items should use cached BOM fallback before filtering.");
    var products = result.Rows[0]["products"] ?? string.Empty;
    Contains("4710964232411", products, "First target SKU is missing.");
    Contains("4710964232565", products, "Second target SKU is missing.");
    Contains("4710964232435", products, "Third target SKU is missing.");
    Contains("MAIN-001", products, "Existing nested item was not preserved.");
    Contains("MAIN-BOM", products, "BOM component was not attached to the combine product.");
    Contains("EGG-EXTRA", products, "A Flavor product whose name contains egg-yolk pastry should remain visible.");
    Contains("BUNDLE-EGG", products, "A combine product whose BOM contains a target SKU should remain visible.");
    DoesNotContain("OTHER-999", products, "Non-target Flavor SKU was not filtered out.");
    DoesNotContain("不應顯示", result.SafeJson, "Product text parsing leaked an unapproved field.");
    DoesNotContain("cost_price", result.SafeJson, "BOM response leaked an unapproved field.");
}

static async Task FlavorBomBusinessFailureIsReported()
{
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath == "/api_v1/token/authorize.php")
        {
            return Json("{\"result\":{\"ok\":true,\"access_token\":\"token\"}}");
        }

        if (request.RequestUri.AbsolutePath == "/api_v1/product/bom.php")
        {
            return Json("{\"result\":{\"ok\":false,\"message\":\"BOM denied secret-key\"}}");
        }

        return Json("""
            {
              "result":{"ok":true,"maxpage":1,"rows":[{
                "order_no":"F002",
                "arrival_date":"2026/08/12",
                "status_code":"F",
                "products":[{"sku":"4710964232411","type":"combine","name":"目標禮盒","qty":1}]
              }]}
            }
            """);
    });
    var client = new WmsApiClient(new HttpClient(handler));

    try
    {
        await client.GetPendingOrdersSinceInitialDateAsync(
            new WmsSource("site2", "Flavor", new Uri("https://example.test")),
            new ApiCredentials("id", "secret-key"),
            new DateOnly(2026, 8, 4));
        throw new InvalidOperationException("Expected a BOM business failure.");
    }
    catch (WmsApiException exception)
    {
        Contains("BOM", exception.Message, "BOM failure should be understandable.");
        DoesNotContain("secret-key", exception.Message, "BOM failure leaked the API Key.");
    }
}

static async Task FlavorBomLookupLimitStopsIncompletePreview()
{
    var products = Enumerable.Range(1, 101)
        .Select(index => new Dictionary<string, object?>
        {
            ["sku"] = $"COMBINE-{index:000}",
            ["type"] = "combine",
            ["name"] = $"蛋黃酥組合 {index}",
            ["qty"] = 1
        })
        .ToArray();
    var orderBody = JsonSerializer.Serialize(new
    {
        result = new
        {
            ok = true,
            maxpage = 1,
            rows = new object[]
            {
                new Dictionary<string, object?>
                {
                    ["order_no"] = "F-LIMIT",
                    ["arrival_date"] = null,
                    ["status_code"] = "F",
                    ["products"] = products
                }
            }
        }
    });
    var bomRequests = 0;
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath == "/api_v1/token/authorize.php")
        {
            return Json("{\"result\":{\"ok\":true,\"access_token\":\"token\"}}");
        }

        if (request.RequestUri.AbsolutePath == "/api_v1/product/bom.php")
        {
            bomRequests++;
            return Json("{\"result\":{\"ok\":true},\"data\":{\"rows\":[]}}");
        }

        return Json(orderBody);
    });
    var client = new WmsApiClient(new HttpClient(handler));

    try
    {
        await client.GetPendingOrdersSinceInitialDateAsync(
            new WmsSource("site2", "Flavor", new Uri("https://example.test")),
            new ApiCredentials("id", "key"),
            new DateOnly(2026, 8, 4));
        throw new InvalidOperationException("Expected the BOM lookup ceiling to stop the preview.");
    }
    catch (WmsApiException exception)
    {
        Contains("100", exception.Message, "The BOM lookup ceiling should be understandable.");
        Equal(100, bomRequests, "The client should stop before sending request 101.");
    }
}

static Task PreviewCanBeExportedToExcel()
{
    var rows = new IReadOnlyDictionary<string, string?>[]
    {
        new Dictionary<string, string?>
        {
            ["order_no"] = "F001",
            ["order_date"] = "2026/08/04 09:00:00",
            ["arrival_date"] = null,
            ["source"] = "flavor",
            ["shop_name"] = "芝初測試賣場",
            ["ship_window_start"] = "2026-08-22",
            ["ship_window_end"] = "2026-08-28",
            ["derived_shipping_date"] = "2026-08-24",
            ["status_code"] = "F",
            ["products"] = """
                [{"SKU":"4710964232411","Type":"combine","Name":"中秋&禮盒","Spec":"6盒組","Qty":2,"Shipp_Qty":1,"Items":[{"SKU":"MAIN-001","Name":"主商品\u0001<一>","Qty":3,"Shipp_Qty":2}]}]
                """
        }
    };
    var result = new PreviewResult(
        true,
        "查詢成功",
        rows,
        "{}",
        new PreviewMetadata("Flavor", "/api_v1/order/order_query.php", DateTimeOffset.UtcNow, 200, 1, 1, true, "查詢成功"));
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-{Guid.NewGuid():N}.xlsx");

    try
    {
        PreviewWorkbookExporter.Export(path, result);
        True(File.Exists(path), "Excel file was not created.");
        using var archive = ZipFile.OpenRead(path);
        Equal(2, archive.Entries.Count(entry => entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal)),
            "Excel should contain order and product worksheets.");
        var productSheet = archive.GetEntry("xl/worksheets/sheet2.xml")
            ?? throw new InvalidOperationException("Product worksheet is missing.");
        using var reader = new StreamReader(productSheet.Open(), Encoding.UTF8);
        var xml = reader.ReadToEnd();
        Contains("4710964232411", xml, "Parent SKU is missing from Excel.");
        Contains("MAIN-001", xml, "Nested product SKU is missing from Excel.");
        Contains("中秋&amp;禮盒", xml, "Product name was not safely encoded in Excel XML.");
        Contains("主商品&lt;一&gt;", xml, "Nested product text was not safely encoded in Excel XML.");
        Contains("芝初測試賣場", xml, "Shop name is missing from Excel.");
        Contains("2026-08-24", xml, "Derived shipping date is missing from Excel.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static async Task CredentialsAreIsolatedAndKeyIsNeverDisplayed()
{
    var vault = new TestCredentialVault();
    var manager = new CredentialManager(vault);

    await manager.SaveAsync("site1", new ApiCredentials("id-1", "key-1"));
    await manager.SaveAsync("site2", new ApiCredentials("id-2", "key-2"));
    await manager.SaveAsync("site1", new ApiCredentials("id-1-new", "key-1-new"));

    var site1 = await manager.GetForUseAsync("site1");
    var site2 = await manager.GetForUseAsync("site2");
    var display = await manager.GetDisplayStatusAsync("site1");

    Equal("id-1-new", site1?.ApiId, "Site 1 API ID was not updated.");
    Equal("key-1-new", site1?.ApiKey, "Site 1 API Key was not updated.");
    Equal("id-2", site2?.ApiId, "Updating site 1 changed site 2 API ID.");
    Equal("key-2", site2?.ApiKey, "Updating site 1 changed site 2 API Key.");
    Equal("id-1-new", display.ApiId, "Display status should include API ID.");
    True(display.HasSavedKey, "Display status should report a saved Key.");
    DoesNotContain("key-1-new", display.ToString() ?? string.Empty, "Display status leaked API Key.");
}

static async Task ErrorMessageNeverEchoesApiKey()
{
    var basicCredential = Convert.ToBase64String(Encoding.UTF8.GetBytes("api-id:super-secret-key"));
    var handler = new RecordingHandler(_ => Json(
        $"{{\"result\":{{\"ok\":false,\"message\":\"invalid credential api-id super-secret-key Basic {basicCredential}\"}}}}"));
    var client = new WmsApiClient(new HttpClient(handler));

    try
    {
        await client.GetShopsAsync(
            new WmsSource("site1", "網站1", new Uri("https://example.test")),
            new ApiCredentials("api-id", "super-secret-key"));
        throw new InvalidOperationException("Expected WmsApiException was not thrown.");
    }
    catch (WmsApiException exception)
    {
        DoesNotContain("api-id", exception.Message, "WMS error message leaked the API ID.");
        DoesNotContain("super-secret-key", exception.Message, "WMS error message leaked API Key.");
        DoesNotContain(basicCredential, exception.Message, "WMS error message leaked Basic credentials.");
    }
}

static async Task SuccessMessageNeverLeaksSecrets()
{
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath.Contains("authorize", StringComparison.Ordinal))
        {
            return Json("{" + "\"result\":{\"ok\":true,\"access_token\":\"secret-token\"}}" );
        }

        return Json("""
            {"result":{"ok":true,"message":"api-key-value secret-token","rows":[]}}
            """);
    });
    var client = new WmsApiClient(new HttpClient(handler));
    var result = await client.GetShopsAsync(
        new WmsSource("site1", "網站1", new Uri("https://example.test")),
        new ApiCredentials("api-id", "api-key-value"));

    DoesNotContain("api-key-value", result.Message, "Success message leaked API Key.");
    DoesNotContain("secret-token", result.Message, "Success message leaked bearer token.");
    DoesNotContain("api-key-value", result.SafeJson, "Safe JSON leaked API Key from message.");
    DoesNotContain("secret-token", result.SafeJson, "Safe JSON leaked bearer token from message.");
}

static async Task InvalidJsonBecomesSafeWmsError()
{
    var handler = new RecordingHandler(_ => new HttpResponseMessage(HttpStatusCode.OK)
    {
        Content = new StringContent("<html>not json</html>", Encoding.UTF8, "text/html")
    });
    var client = new WmsApiClient(new HttpClient(handler));

    try
    {
        await client.GetShopsAsync(
            new WmsSource("site1", "網站1", new Uri("https://example.test")),
            new ApiCredentials("id", "key"));
        throw new InvalidOperationException("Expected WmsApiException was not thrown.");
    }
    catch (WmsApiException exception)
    {
        Contains("JSON", exception.Message, "Invalid JSON error is not understandable.");
        DoesNotContain("<html>", exception.Message, "Invalid response body leaked into error message.");
    }
}

static async Task TransientServerErrorsAreRetried()
{
    var authorizeAttempts = 0;
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath.Contains("authorize", StringComparison.Ordinal))
        {
            authorizeAttempts++;
            if (authorizeAttempts < 3)
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError)
                {
                    Content = new StringContent("{}", Encoding.UTF8, "application/json")
                };
            }

            return Json("{" + "\"result\":{\"ok\":true,\"access_token\":\"token\",\"expires_in\":3600}}" );
        }

        return Json("{" + "\"result\":{\"ok\":true,\"rows\":[]}}" );
    });
    var client = new WmsApiClient(new HttpClient(handler), (_, _) => Task.CompletedTask);

    var result = await client.GetShopsAsync(
        new WmsSource("site1", "網站1", new Uri("https://example.test")),
        new ApiCredentials("id", "key"));

    True(result.IsSuccess, "Request should succeed after transient retries.");
    Equal(3, authorizeAttempts, "Transient server errors were not retried exactly three attempts.");
}

static async Task ValidTokenIsCached()
{
    var authorizeRequests = 0;
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath.Contains("authorize", StringComparison.Ordinal))
        {
            authorizeRequests++;
            return Json("{" + "\"result\":{\"ok\":true,\"access_token\":\"cached-token\",\"expires_in\":3600}}" );
        }

        Equal("cached-token", request.Headers.Authorization?.Parameter, "Cached token was not used.");
        return Json("{" + "\"result\":{\"ok\":true,\"rows\":[]}}" );
    });
    var client = new WmsApiClient(new HttpClient(handler));
    var source = new WmsSource("site1", "網站1", new Uri("https://example.test"));
    var credentials = new ApiCredentials("id", "key");

    await client.GetShopsAsync(source, credentials);
    await client.GetShopsAsync(source, credentials);

    Equal(1, authorizeRequests, "A valid token should be reused instead of re-authorizing.");
}

static async Task RevokedTokenIsRefreshedOnce()
{
    var authorizeRequests = 0;
    var shopRequests = 0;
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath.Contains("authorize", StringComparison.Ordinal))
        {
            authorizeRequests++;
            var token = authorizeRequests == 1 ? "revoked-token" : "fresh-token";
            return Json("{" + $"\"result\":{{\"ok\":true,\"access_token\":\"{token}\",\"expires_in\":3600}}" + "}" );
        }

        shopRequests++;
        if (request.Headers.Authorization?.Parameter == "revoked-token")
        {
            return new HttpResponseMessage(HttpStatusCode.Unauthorized)
            {
                Content = new StringContent("{}", Encoding.UTF8, "application/json")
            };
        }

        Equal("fresh-token", request.Headers.Authorization?.Parameter, "Fresh token was not used after 401.");
        return Json("{" + "\"result\":{\"ok\":true,\"rows\":[]}}" );
    });
    var client = new WmsApiClient(new HttpClient(handler), (_, _) => Task.CompletedTask);

    var result = await client.GetShopsAsync(
        new WmsSource("site1", "網站1", new Uri("https://example.test")),
        new ApiCredentials("id", "key"));

    True(result.IsSuccess, "Request should succeed after refreshing a revoked token.");
    Equal(2, authorizeRequests, "401 should trigger exactly one re-authorization.");
    Equal(2, shopRequests, "Shop request should retry exactly once with the fresh token.");
}

static async Task EditedApiIdCannotReuseSavedKey()
{
    var vault = new TestCredentialVault();
    var manager = new CredentialManager(vault);
    await manager.SaveAsync("site1", new ApiCredentials("saved-id", "saved-key"));

    try
    {
        await manager.ResolveForUseAsync("site1", "edited-id", string.Empty);
        throw new InvalidOperationException("Expected edited API ID to require a new Key.");
    }
    catch (InvalidOperationException exception)
    {
        Contains("API ID", exception.Message, "Credential mismatch error is not understandable.");
    }

    var stored = await manager.ResolveForUseAsync("site1", "saved-id", string.Empty);
    Equal("saved-key", stored.ApiKey, "Matching saved API ID should reuse the saved Key.");
}

static async Task SyncFailurePreservesLastSuccessfulRun()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-sync-{Guid.NewGuid():N}");
    var path = Path.Combine(directory, "sync-status.json");
    var store = new SyncStatusStore(path);
    var startedAt = new DateTimeOffset(2026, 8, 4, 9, 0, 0, TimeSpan.FromHours(8));
    var succeededAt = startedAt.AddMinutes(2);

    try
    {
        await store.RecordSuccessAsync(new SyncRunSummary(
            "site1",
            "網站1－睿驛",
            new DateOnly(2026, 7, 1),
            new DateOnly(2026, 8, 4),
            startedAt,
            succeededAt,
            5,
            225));

        var successfulStatus = await new SyncStatusStore(path).GetAsync("site1")
            ?? throw new InvalidOperationException("Successful sync status was not persisted.");
        Equal("success", successfulStatus.Status, "The completed run should be marked successful.");
        Equal(5, successfulStatus.PageCount, "The completed page count was not persisted.");
        Equal(225, successfulStatus.RecordCount, "The completed record count was not persisted.");
        Equal(new DateOnly(2026, 8, 4), successfulStatus.QueryThrough, "The successful query end was not persisted.");

        await store.RecordFailureAsync(new SyncRunFailure(
            "site1",
            "網站1－睿驛",
            new DateOnly(2026, 7, 1),
            new DateOnly(2026, 8, 5),
            startedAt.AddDays(1),
            startedAt.AddDays(1).AddMinutes(1),
            "WMS 暫時無法連線"));

        var status = await new SyncStatusStore(path).GetAsync("site1")
            ?? throw new InvalidOperationException("Sync status was not persisted.");
        Equal("failed", status.Status, "The latest run should be marked failed.");
        Equal(succeededAt, status.LastSuccessAt, "A failed run must not erase the last successful time.");
        Equal(new DateOnly(2026, 7, 1), status.QueryFrom, "The query start was not preserved.");
        Equal(new DateOnly(2026, 8, 5), status.QueryThrough, "The failed run query end was not recorded.");
        Equal(0, status.PageCount, "A failed run must not claim completed pages.");
        Equal(0, status.RecordCount, "A failed run must not claim committed records.");
        Contains("WMS 暫時無法連線", status.ErrorSummary, "The safe failure summary was not recorded.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
}

static Task SyncQueryDateUsesTaipeiTimeZone()
{
    var utcInstant = new DateTimeOffset(2026, 8, 4, 16, 30, 0, TimeSpan.Zero);
    Equal(new DateOnly(2026, 8, 5), WmsQueryPolicy.GetTaipeiDate(utcInstant),
        "UTC evening should already be the next order-query date in Taipei.");
    return Task.CompletedTask;
}

static async Task OrderSnapshotReplacesOneSourceWithoutDuplicates()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-orders-{Guid.NewGuid():N}");
    var path = Path.Combine(directory, "order-snapshot.json");
    var store = new OrderSnapshotStore(path);
    var synchronizedAt = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    var site1Rows = new IReadOnlyDictionary<string, string?>[]
    {
        new Dictionary<string, string?>
        {
            ["order_no"] = "SAME-001",
            ["order_date"] = "2026/08/04 11:36:00",
            ["arrival_date"] = null,
            ["source_key"] = "shopee",
            ["shop_name"] = "芝初 SesaOle",
            ["derived_shipping_date"] = "2026-08-24",
            ["status_code"] = "F",
            ["products"] = """
                [{"sku":"BOX-001","item_no":"P01","name":"蛋黃酥禮盒","spec":"8/22-8/28當週出貨","qty":1,"shipp_qty":0,
                  "items":[{"sku":"4710964232565","item_no":"I01","name":"蛋黃酥9入","qty":9,"shipp_qty":0}]}]
                """
        }
    };
    var site2Rows = new IReadOnlyDictionary<string, string?>[]
    {
        new Dictionary<string, string?>
        {
            ["order_no"] = "SAME-001",
            ["source_key"] = "official",
            ["status_code"] = "F",
            ["products"] = """[{"sku":"4710964232411","name":"Flavor 蛋黃酥","qty":2,"shipp_qty":1}]"""
        }
    };

    try
    {
        await store.ReplaceSourceAsync("site1", site1Rows, synchronizedAt);
        await store.ReplaceSourceAsync("site2", site2Rows, synchronizedAt);

        var updatedSite1Rows = new IReadOnlyDictionary<string, string?>[]
        {
            new Dictionary<string, string?>
            {
                ["order_no"] = "SAME-001",
                ["source_key"] = "shopee",
                ["status_code"] = "F",
                ["products"] = """[{"sku":"BOX-001","item_no":"P01","name":"蛋黃酥禮盒","qty":3,"shipp_qty":1}]"""
            }
        };
        await new OrderSnapshotStore(path).ReplaceSourceAsync("site1", updatedSite1Rows, synchronizedAt.AddMinutes(5));

        var snapshot = await new OrderSnapshotStore(path).GetAllAsync();
        Equal(2, snapshot.Count, "Replacing site1 should remove its stale item without duplicating either source.");
        var site1 = snapshot.Single(line => line.SourceCode == "site1");
        var site2 = snapshot.Single(line => line.SourceCode == "site2");
        Equal("SAME-001", site1.ExternalOrderNo, "The site1 order number was not preserved.");
        Equal("parent:item_no:P01:1", site1.ExternalLineKey, "The stable parent line key is incorrect.");
        Equal(3m, site1.Quantity, "The repeated site1 sync did not update quantity.");
        Equal(1m, site1.ShippedQuantity, "The repeated site1 sync did not update shipped quantity.");
        Equal("4710964232411", site2.Sku, "Replacing site1 incorrectly changed site2.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
}

static HttpResponseMessage Json(string value) => new(HttpStatusCode.OK)
{
    Content = new StringContent(value, Encoding.UTF8, "application/json")
};

static void True(bool value, string message)
{
    if (!value) throw new InvalidOperationException(message);
}

static void Equal<T>(T expected, T actual, string message)
{
    if (!EqualityComparer<T>.Default.Equals(expected, actual))
        throw new InvalidOperationException($"{message} Expected: {expected}; Actual: {actual}");
}

static void Contains(string expected, string actual, string message)
{
    if (!actual.Contains(expected, StringComparison.Ordinal))
        throw new InvalidOperationException(message);
}

static void DoesNotContain(string unexpected, string actual, string message)
{
    if (actual.Contains(unexpected, StringComparison.OrdinalIgnoreCase))
        throw new InvalidOperationException(message);
}

sealed class RecordingHandler(Func<HttpRequestMessage, HttpResponseMessage> responseFactory) : HttpMessageHandler
{
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        var response = responseFactory(request);
        response.RequestMessage = request;
        return Task.FromResult(response);
    }
}

sealed class TestCredentialVault : ICredentialVault
{
    private readonly Dictionary<string, ApiCredentials> _items = new(StringComparer.OrdinalIgnoreCase);

    public Task SaveAsync(string sourceCode, ApiCredentials credentials, CancellationToken cancellationToken = default)
    {
        _items[sourceCode] = credentials;
        return Task.CompletedTask;
    }

    public Task<ApiCredentials?> LoadAsync(string sourceCode, CancellationToken cancellationToken = default)
    {
        _items.TryGetValue(sourceCode, out var credentials);
        return Task.FromResult(credentials);
    }
}
