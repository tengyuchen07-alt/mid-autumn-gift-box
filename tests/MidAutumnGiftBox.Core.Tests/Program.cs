using System.Net;
using System.Net.Http.Headers;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
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
    ("訂單快照遇到重複唯一鍵會拒絕整批資料", OrderSnapshotRejectsDuplicateUniqueKeys),
    ("訂單快照遇到無效數量會保留舊資料", OrderSnapshotRejectsInvalidQuantities),
    ("已提交快照會重建相同的畫面與 Excel 父子商品", CommittedSnapshotBuildsPreviewAndExcel),
    ("同步 checkpoint 不會縮短待處理訂單的完整成立日範圍", SyncWindowKeepsFullOrderDateRange),
    ("增量快照只取代來源的訂單成立日範圍", IncrementalSnapshotReplacesOnlyQueriedOrderDates),
    ("增量訂單 API 會送出明確的成立日起訖", ExplicitOrderRangeIsSentToWms),
    ("離開待處理的訂單可依訂單編號查回最終狀態", OrderStatusCanBeQueriedByOrderNumber),
    ("同步差異只以父商品數量產生追加紀錄", ParentQuantityChangesBecomeAppendOnlyEntries),
    ("離開待處理後只有取消退貨沖銷且刪除併單待人工確認", DepartedOrdersUseResolvedStatusRules),
    ("異動紀錄只追加且重試不會重複寫入", ChangeLedgerIsAppendOnlyAndIdempotent),
    ("Excel 追加新異動時會保留人工確認欄", ExcelAppendPreservesManualConfirmation),
    ("兩網站禮盒會依三六九入輸出且數量為 Excel 數字", ThreeGiftBoxSheetsCombineSourcesAndUseNumericQuantities),
    ("三種禮盒 Excel 追加時保留確認且不重複", ThreeGiftBoxWorkbookAppendPreservesConfirmation),
    ("舊異動名稱依蛋黃酥入數分流且不誤用芝麻粉倍數", GiftBoxSizeUsesEggYolkCountNotOtherMultipliers),
    ("同一來源的本機同步鎖同時間只能由一個程序取得", SourceSyncLockIsExclusive),
    ("最近同步失敗不會把最後提交快照標成失敗資料", FailedSyncKeepsCommittedPreviewValid),
    ("快照投影遇到孤立子商品會拒絕顯示", SnapshotPreviewRejectsOrphanItems),
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
                "order_no":"WEEKDAY-START",
                "order_date":"2026/09/15 09:00:00",
                "status_code":"F",
                "products":[{
                  "name":"(預購)【芝初】2026年黑芝麻Q潤蛋黃酥9入禮盒-ECTV",
                  "spec":"9/15~9/18出貨",
                  "qty":1
                }]
              },
              {
                "order_no":"WEEKEND-ONLY",
                "order_date":"2026/08/22 09:00:00",
                "status_code":"F",
                "products":[{
                  "name":"(預購)【芝初】2026年黑芝麻Q潤蛋黃酥9入禮盒",
                  "spec":"8/22~8/23出貨",
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

    Equal(8, result.Rows.Count, "All egg-yolk pastry orders should remain visible.");
    var fuzzy = result.Rows.Single(row => row["order_no"] == "FUZZY");
    Equal("原賣場", fuzzy["source"], "Unknown source_key should preserve the original source.");
    True(!fuzzy.ContainsKey("derived_shipping_date"), "Fuzzy month text must not derive an exact date.");
    var conflict = result.Rows.Single(row => row["order_no"] == "CONFLICT");
    True(!conflict.ContainsKey("shipping_date_status"),
        "Different parent products in one order must not create an order-level conflict.");
    True(!conflict.ContainsKey("derived_shipping_date"),
        "An order with multiple parent shipping dates must not expose one order-level date.");
    using (var conflictProducts = JsonDocument.Parse(conflict["products"]!))
    {
        var firstProduct = conflictProducts.RootElement[0];
        var secondProduct = conflictProducts.RootElement[1];
        Equal("2026-08-24", firstProduct.GetProperty("derived_shipping_date").GetString(),
            "The first parent product should keep its own shipping Monday.");
        Equal("2026-08-31", secondProduct.GetProperty("derived_shipping_date").GetString(),
            "The second parent product should keep its own shipping Monday.");
    }
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
    var weekdayStart = result.Rows.Single(row => row["order_no"] == "WEEKDAY-START");
    Equal("2026-09-15", weekdayStart["derived_shipping_date"],
        "A range without Monday should use its first in-range workday.");
    var weekendOnly = result.Rows.Single(row => row["order_no"] == "WEEKEND-ONLY");
    Equal("needs_review", weekendOnly["shipping_date_status"],
        "A weekend-only range should require review because it has no in-range workday.");
    True(!weekendOnly.ContainsKey("derived_shipping_date"),
        "A weekend-only range must not choose a date outside the range.");
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
                  "ship_window_start":"2026-09-05","ship_window_end":"2026-09-11",
                  "derived_shipping_date":"2026-09-07","shipping_date_source":"product.spec","shipping_date_status":"derived",
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
        var initialSnapshot = await store.GetAllAsync();
        Equal(3, initialSnapshot.Count, "Parent and nested item lines were not both expanded.");
        var nestedItem = initialSnapshot.Single(line => line.SourceCode == "site1" && line.LineLevel == "item");
        Equal("4710964232565", nestedItem.Sku, "The nested target SKU was not persisted.");
        Equal("parent:item_no:P01:1", nestedItem.ParentLineKey, "The nested item lost its parent line key.");
        Equal("2026-09-07", nestedItem.DerivedShippingDate,
            "A nested item should inherit its parent product shipping date.");
        var initialParent = initialSnapshot.Single(line => line.SourceCode == "site1" && line.LineLevel == "parent");
        Equal("2026-09-07", initialParent.DerivedShippingDate,
            "A parent line should prefer its own shipping date over the order-level fallback.");

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

static async Task OrderSnapshotRejectsDuplicateUniqueKeys()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-duplicate-{Guid.NewGuid():N}");
    var path = Path.Combine(directory, "order-snapshot.json");
    var store = new OrderSnapshotStore(path);
    IReadOnlyDictionary<string, string?> Row(string quantity) => new Dictionary<string, string?>
    {
        ["order_no"] = "DUP-001",
        ["status_code"] = "F",
        ["products"] = $"[{{\"sku\":\"4710964232411\",\"qty\":{quantity}}}]"
    };

    try
    {
        await store.ReplaceSourceAsync("site1", [Row("1")], DateTimeOffset.UtcNow);
        try
        {
            await store.ReplaceSourceAsync("site1", [Row("2"), Row("3")], DateTimeOffset.UtcNow);
            throw new InvalidOperationException("Expected duplicate unique keys to reject the snapshot.");
        }
        catch (InvalidDataException exception)
        {
            Contains("重複", exception.Message, "Duplicate-key failure should be understandable.");
        }

        var snapshot = await store.GetAllAsync();
        Equal(1, snapshot.Count, "A rejected duplicate snapshot changed the committed data.");
        Equal(1m, snapshot[0].Quantity, "A rejected duplicate snapshot replaced the prior quantity.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
}

static async Task OrderSnapshotRejectsInvalidQuantities()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-quantity-{Guid.NewGuid():N}");
    var path = Path.Combine(directory, "order-snapshot.json");
    var store = new OrderSnapshotStore(path);
    IReadOnlyDictionary<string, string?> Row(string products) => new Dictionary<string, string?>
    {
        ["order_no"] = "QTY-001",
        ["products"] = products
    };

    try
    {
        await store.ReplaceSourceAsync("site1",
            [Row("""[{"sku":"4710964232411","qty":2,"shipp_qty":1}]""")],
            DateTimeOffset.UtcNow);
        try
        {
            await store.ReplaceSourceAsync("site1",
                [Row("""[{"sku":"4710964232411","qty":"不是數字","shipp_qty":1}]""")],
                DateTimeOffset.UtcNow);
            throw new InvalidOperationException("Expected invalid quantity to reject the snapshot.");
        }
        catch (InvalidDataException exception)
        {
            Contains("數量", exception.Message, "Invalid-quantity failure should be understandable.");
        }

        var snapshot = await store.GetAllAsync();
        Equal(2m, snapshot.Single().Quantity, "An invalid quantity replaced the committed snapshot.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
}

static Task CommittedSnapshotBuildsPreviewAndExcel()
{
    var synchronizedAt = new DateTimeOffset(2026, 8, 5, 9, 10, 0, TimeSpan.FromHours(8));
    var lines = new OrderLineSnapshot[]
    {
        new("site1", "R001", "parent:item_no:P01:1", "parent", null,
            "shopee", "芝初 SesaOle", "2026/08/04 11:36:00", null, "2026-08-24",
            "BOX-001", "P01", "蛋黃酥禮盒", "8/22-8/28當週出貨", 1m, 0m, "F", synchronizedAt),
        new("site1", "R001", "item:parent:item_no:P01:1:item_no:I01:1", "item", "parent:item_no:P01:1",
            "shopee", "芝初 SesaOle", "2026/08/04 11:36:00", null, "2026-08-24",
            "4710964232565", "I01", "蛋黃酥9入", null, 9m, 0m, "F", synchronizedAt)
    };
    var status = new SyncSourceStatus(
        "site1", "網站1－睿驛", new DateOnly(2026, 7, 1), new DateOnly(2026, 8, 5),
        "success", synchronizedAt.AddMinutes(-2), synchronizedAt, 3, 1, synchronizedAt, string.Empty);

    var preview = OrderSnapshotPreviewBuilder.Build(
        new WmsSource("site1", "網站1－睿驛", new Uri("https://example.test")), lines, status);
    Equal(1, preview.Rows.Count, "Snapshot lines should rebuild one order row.");
    Equal("R001", preview.Rows[0]["order_no"], "The rebuilt order number is incorrect.");
    Contains("BOX-001", preview.Rows[0]["products"] ?? string.Empty, "The rebuilt parent product is missing.");
    Contains("4710964232565", preview.Rows[0]["products"] ?? string.Empty, "The rebuilt nested item is missing.");
    Contains("4710964232565", preview.SafeJson, "Safe JSON was not rebuilt from the committed snapshot.");

    var path = Path.Combine(Path.GetTempPath(), $"snapshot-preview-{Guid.NewGuid():N}.xlsx");
    try
    {
        PreviewWorkbookExporter.Export(path, preview);
        using var archive = ZipFile.OpenRead(path);
        var productSheet = archive.GetEntry("xl/worksheets/sheet2.xml")
            ?? throw new InvalidOperationException("Product worksheet is missing.");
        using var reader = new StreamReader(productSheet.Open(), Encoding.UTF8);
        var xml = reader.ReadToEnd();
        Contains("BOX-001", xml, "Excel was not generated from the committed parent product.");
        Contains("4710964232565", xml, "Excel was not generated from the committed nested item.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task SyncWindowKeepsFullOrderDateRange()
{
    var startedAt = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    var first = SyncWindowPolicy.Create(null, startedAt);
    Equal(new DateOnly(2026, 7, 1), first.FromDate, "First sync should use the initial order date.");
    Equal(new DateOnly(2026, 8, 5), first.ThroughDate, "Sync upper date should be fixed at run start.");

    var lastSuccess = new DateTimeOffset(2026, 8, 4, 16, 30, 0, TimeSpan.Zero);
    var refresh = SyncWindowPolicy.Create(lastSuccess, startedAt);
    Equal(new DateOnly(2026, 7, 1), refresh.FromDate,
        "Last success must not hide older orders whose pending status may have changed.");
    Equal(startedAt, refresh.StartedAt, "The run start instant should remain fixed.");
    return Task.CompletedTask;
}

static async Task IncrementalSnapshotReplacesOnlyQueriedOrderDates()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-range-{Guid.NewGuid():N}");
    var path = Path.Combine(directory, "order-snapshot.json");
    var store = new OrderSnapshotStore(path);
    IReadOnlyDictionary<string, string?> Row(string orderNo, string orderDate, int quantity) =>
        new Dictionary<string, string?>
        {
            ["order_no"] = orderNo,
            ["order_date"] = orderDate,
            ["products"] = $"[{{\"sku\":\"4710964232411\",\"qty\":{quantity}}}]"
        };

    try
    {
        await store.ReplaceSourceAsync("site1",
            [Row("OLD", "2026/07/10 08:00:00", 1), Row("REMOVE", "2026/08/04 08:00:00", 2)],
            DateTimeOffset.UtcNow);
        await store.ReplaceSourceAsync("site2", [Row("OTHER", "2026/08/04 08:00:00", 4)], DateTimeOffset.UtcNow);

        await new OrderSnapshotStore(path).ReplaceSourceRangeAsync(
            "site1",
            new DateOnly(2026, 8, 4),
            new DateOnly(2026, 8, 5),
            [Row("NEW", "2026/08/05 09:00:00", 3)],
            DateTimeOffset.UtcNow);

        var snapshot = await store.GetAllAsync();
        Equal(3, snapshot.Count, "Range replacement kept stale rows or removed out-of-range/source rows.");
        True(snapshot.Any(line => line.SourceCode == "site1" && line.ExternalOrderNo == "OLD"),
            "An older site1 order outside the queried range was removed.");
        True(snapshot.Any(line => line.SourceCode == "site1" && line.ExternalOrderNo == "NEW" && line.Quantity == 3m),
            "The new in-range site1 order was not inserted.");
        True(snapshot.Any(line => line.SourceCode == "site2" && line.ExternalOrderNo == "OTHER"),
            "Replacing site1 range changed site2.");
        True(snapshot.All(line => line.ExternalOrderNo != "REMOVE"),
            "A stale site1 order inside the queried range was not removed.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
}

static async Task ExplicitOrderRangeIsSentToWms()
{
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath == "/api_v1/token/authorize.php")
        {
            return Json("{\"result\":{\"ok\":true,\"access_token\":\"token\"}}");
        }

        Contains("date_min=2026-08-04", request.RequestUri.Query, "Incremental start date was not sent.");
        Contains("date_max=2026-08-05", request.RequestUri.Query, "Fixed run upper date was not sent.");
        Contains("status=F", request.RequestUri.Query, "Pending status was not sent.");
        return Json("{\"result\":{\"ok\":true},\"data\":{\"maxpage\":1,\"rows\":[]}}");
    });
    var client = new WmsApiClient(new HttpClient(handler));

    var result = await client.GetPendingOrdersAsync(
        new WmsSource("site1", "網站1", new Uri("https://example.test")),
        new ApiCredentials("id", "key"),
        new DateOnly(2026, 8, 4),
        new DateOnly(2026, 8, 5));
    True(result.IsSuccess, "Explicit incremental range query should succeed.");
}

static async Task OrderStatusCanBeQueriedByOrderNumber()
{
    var handler = new RecordingHandler(request =>
    {
        if (request.RequestUri!.AbsolutePath == "/api_v1/token/authorize.php")
        {
            return Json("{\"result\":{\"ok\":true,\"access_token\":\"token\"}}");
        }

        Contains("order_no=2607303JVTVQVW", request.RequestUri.Query,
            "Order-number status query did not send the exact order number.");
        DoesNotContain("status=F", request.RequestUri.Query,
            "Order-number status query must not hide cancelled orders behind the pending filter.");
        return Json("""
            {"result":{"ok":true},"data":{"total":1,"rows":[{
              "order_no":"2607303JVTVQVW","status_code":"N","status_name":"已取消",
              "order_date":"2026/07/30 23:09:33","products":[]
            }]}}
            """);
    });
    var client = new WmsApiClient(new HttpClient(handler));

    var result = await client.GetOrderByNumberAsync(
        new WmsSource("site1", "網站1", new Uri("https://example.test")),
        new ApiCredentials("id", "key"),
        "2607303JVTVQVW");

    True(result.IsSuccess, "Order-number status query should succeed.");
    Equal(1, result.Rows.Count, "Exact order-number query should return one order.");
    Equal("N", result.Rows[0]["status_code"], "Cancelled status was not preserved.");
}

static Task ParentQuantityChangesBecomeAppendOnlyEntries()
{
    var previousAt = new DateTimeOffset(2026, 8, 4, 9, 0, 0, TimeSpan.FromHours(8));
    var observedAt = previousAt.AddDays(1);
    OrderLineSnapshot Line(string orderNo, string lineKey, string sku, decimal quantity, string level = "parent") =>
        new("site1", orderNo, lineKey, level, level == "item" ? "parent:A" : null,
            "shopee", "賣場", "2026/08/04 08:00:00", null, "2026-08-24",
            sku, null, sku + " 商品", null, quantity, 0m, "F", previousAt,
            "蝦皮賣場");

    var previous = new[]
    {
        Line("UP", "parent:A", "SKU-UP", 2m),
        Line("DOWN", "parent:A", "SKU-DOWN", 5m),
        Line("REMOVE", "parent:A", "SKU-REMOVE", 4m),
        Line("REMOVE", "parent:B", "SKU-KEEP", 1m),
        Line("UP", "item:parent:A:I", "CHILD", 9m, "item")
    };
    var current = new[]
    {
        Line("UP", "parent:A", "SKU-UP", 5m),
        Line("DOWN", "parent:A", "SKU-DOWN", 2m),
        Line("REMOVE", "parent:B", "SKU-KEEP", 1m),
        Line("NEW", "parent:A", "SKU-NEW", 6m),
        Line("UP", "item:parent:A:I", "CHILD", 18m, "item")
    };

    var entries = OrderChangePlanner.Plan(
        "網站1－睿驛", previous, current,
        new Dictionary<string, OrderStatusResolution>(StringComparer.OrdinalIgnoreCase), observedAt);

    Equal(4, entries.Count, "Only four parent-product changes should be appended.");
    Equal(3m, entries.Single(entry => entry.ExternalOrderNo == "UP").QuantityChange,
        "Parent increase should append only the positive difference.");
    Equal(-3m, entries.Single(entry => entry.ExternalOrderNo == "DOWN").QuantityChange,
        "Parent decrease should append only the negative difference.");
    Equal(-4m, entries.Single(entry => entry.ExternalOrderNo == "REMOVE").QuantityChange,
        "Removing one product from a still-pending order should reverse that product quantity.");
    Equal(6m, entries.Single(entry => entry.ExternalOrderNo == "NEW").QuantityChange,
        "A new parent product should append its full quantity.");
    True(entries.All(entry => entry.Confirmation == string.Empty),
        "Program-created confirmation cells must remain blank.");
    return Task.CompletedTask;
}

static Task DepartedOrdersUseResolvedStatusRules()
{
    var synchronizedAt = new DateTimeOffset(2026, 8, 4, 9, 0, 0, TimeSpan.FromHours(8));
    OrderLineSnapshot Line(string orderNo, decimal quantity) =>
        new("site1", orderNo, "parent:A", "parent", null,
            "shopee", "賣場", "2026/08/04 08:00:00", null, "2026-08-24",
            "SKU-" + orderNo, null, orderNo + " 商品", null, quantity, 0m, "F", synchronizedAt,
            "蝦皮賣場");
    var previous = new[]
    {
        Line("CANCEL", 2m),
        Line("RETURN", 3m),
        Line("DELETE", 4m),
        Line("SHIPPED", 5m)
    };
    var statuses = new Dictionary<string, OrderStatusResolution>(StringComparer.OrdinalIgnoreCase)
    {
        ["CANCEL"] = new("CANCEL", "N", "已取消"),
        ["RETURN"] = new("RETURN", "R", "已退貨"),
        ["DELETE"] = new("DELETE", "D", "已封存"),
        ["SHIPPED"] = new("SHIPPED", "S", "已出貨")
    };

    var entries = OrderChangePlanner.Plan(
        "網站1－睿驛", previous, Array.Empty<OrderLineSnapshot>(), statuses,
        synchronizedAt.AddDays(1));

    Equal(3, entries.Count, "Shipped orders must not create a cancellation adjustment.");
    var cancelled = entries.Single(entry => entry.ExternalOrderNo == "CANCEL");
    Equal(-2m, cancelled.QuantityChange, "Cancelled order should reverse the remaining parent quantity.");
    Equal("取消", cancelled.ChangeType, "Cancelled order should be labelled as cancellation.");
    var returned = entries.Single(entry => entry.ExternalOrderNo == "RETURN");
    Equal(-3m, returned.QuantityChange, "Returned order should reverse the remaining parent quantity.");
    var deleted = entries.Single(entry => entry.ExternalOrderNo == "DELETE");
    Equal(0m, deleted.QuantityChange, "Deleted or merged order must not change quantity automatically.");
    True(deleted.NeedsReview, "Deleted or merged order should be marked for manual review.");
    return Task.CompletedTask;
}

static async Task ChangeLedgerIsAppendOnlyAndIdempotent()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-ledger-{Guid.NewGuid():N}");
    var store = new OrderChangeLedgerStore(Path.Combine(directory, "order-change-ledger.json"));
    var at = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(string id, string orderNo, decimal change) =>
        new(id, at, "site1", "網站1－睿驛", orderNo, "parent:A", "蝦皮賣場",
            "2026/08/04 08:00:00", "2026-08-24", "SKU", "蛋黃酥禮盒",
            change < 0 ? -change : 0m, change > 0 ? change : 0m, change,
            change < 0 ? "取消" : "新增", change < 0 ? "N" : "F",
            change < 0 ? "已取消" : "待處理", false, string.Empty);

    try
    {
        var first = Entry("E1", "ORDER-1", 2m);
        var second = Entry("E2", "ORDER-2", -1m);
        await store.AppendAsync([first]);
        await new OrderChangeLedgerStore(Path.Combine(directory, "order-change-ledger.json"))
            .AppendAsync([first, second]);

        var saved = await store.GetAllAsync();
        Equal(2, saved.Count, "Retrying the same change should not duplicate a ledger row.");
        Equal("E1", saved[0].EntryId, "Existing ledger history was overwritten.");
        Equal("E2", saved[1].EntryId, "New ledger entry was not appended after existing history.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
}

static Task ExcelAppendPreservesManualConfirmation()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-ledger-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    var preview = new PreviewResult(true, "ok", Array.Empty<IReadOnlyDictionary<string, string?>>(), "{}",
        new PreviewMetadata("網站1－睿驛", "/api_v1/order/order_query.php", at, 200, 1, 0, true, "ok"));
    OrderChangeEntry Entry(string id, string orderNo) =>
        new(id, at, "site1", "網站1－睿驛", orderNo, "parent:A", "蝦皮賣場",
            "2026/08/04 08:00:00", "2026-08-24", "SKU", "蛋黃酥禮盒",
            0m, 2m, 2m, "新增", "F", "待處理", false, string.Empty);

    try
    {
        var first = Entry("ENTRY-1", "ORDER-1");
        var second = Entry("ENTRY-2", "ORDER-2");
        PreviewWorkbookExporter.Export(path, preview, [first]);

        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = archive.GetEntry("xl/worksheets/sheet3.xml")
                ?? throw new InvalidOperationException("Change worksheet is missing.");
            XDocument document;
            using (var stream = entry.Open()) document = XDocument.Load(stream);
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var confirmationCell = document.Descendants(spreadsheet + "c")
                .Single(cell => (string?)cell.Attribute("r") == "P2");
            confirmationCell.Descendants(spreadsheet + "t").Single().Value = "已確認";
            entry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet3.xml");
            using var output = replacement.Open();
            document.Save(output);
        }

        PreviewWorkbookExporter.Export(path, preview, [first, second]);
        using var resultArchive = ZipFile.OpenRead(path);
        var resultEntry = resultArchive.GetEntry("xl/worksheets/sheet3.xml")
            ?? throw new InvalidOperationException("Change worksheet is missing after append.");
        using var reader = new StreamReader(resultEntry.Open(), Encoding.UTF8);
        var xml = reader.ReadToEnd();
        Contains("確認", xml, "Change worksheet does not contain the confirmation column.");
        Contains("已確認", xml, "Manual confirmation was overwritten during append.");
        Contains("ENTRY-2", xml, "New change row was not appended.");
        Equal(1, xml.Split("ENTRY-1", StringSplitOptions.None).Length - 1,
            "Existing change row was duplicated during append.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task ThreeGiftBoxSheetsCombineSourcesAndUseNumericQuantities()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-three-box-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(string id, string sourceCode, string sourceName, string orderNo, string sku,
        string productName, decimal quantity) =>
        new(id, at, sourceCode, sourceName, orderNo, "parent:A", "蝦皮賣場",
            "2026/08/04 08:00:00", "2026-08-24", sku, productName,
            0m, quantity, quantity, "新增", "F", "待處理", false, string.Empty);
    var entries = new[]
    {
        Entry("THREE", "site1", "網站1－睿驛", "R-3", "4710964232435", "蛋黃酥小禮盒", 2m),
        Entry("SIX", "site2", "網站2－Flavor", "F-6", "4710964232411", "蛋黃酥禮盒", 3m),
        Entry("NINE", "site1", "網站1－睿驛", "R-9", "4710964232565", "蛋黃酥9入禮盒", 4m),
        Entry("SINGLE", "site1", "網站1－睿驛", "R-1", "SINGLE", "蛋黃酥 單入包裝", 1m)
    };

    try
    {
        GiftBoxWorkbookExporter.Export(path, entries);
        using var archive = ZipFile.OpenRead(path);
        using (var reader = new StreamReader(
                   (archive.GetEntry("xl/workbook.xml")
                    ?? throw new InvalidOperationException("Workbook metadata is missing.")).Open(), Encoding.UTF8))
        {
            var workbookXml = reader.ReadToEnd();
            Contains("三入", workbookXml, "Three-piece worksheet is missing.");
            Contains("六入", workbookXml, "Six-piece worksheet is missing.");
            Contains("九入", workbookXml, "Nine-piece worksheet is missing.");
            DoesNotContain("訂單預覽", workbookXml, "Legacy preview worksheet should not remain.");
        }

        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XDocument Sheet(string name)
        {
            using var stream = (archive.GetEntry(name)
                                ?? throw new InvalidOperationException($"Worksheet {name} is missing.")).Open();
            return XDocument.Load(stream);
        }

        var three = Sheet("xl/worksheets/sheet1.xml");
        Contains("R-3", three.ToString(), "Site1 three-piece order is missing.");
        DoesNotContain("R-1", three.ToString(), "Single-piece item must not enter the three-piece worksheet.");
        var quantityCell = three.Descendants(spreadsheet + "c")
            .Single(cell => (string?)cell.Attribute("r") == "D2");
        Equal(null, (string?)quantityCell.Attribute("t"), "Quantity cell must not be stored as text.");
        Equal("2", quantityCell.Element(spreadsheet + "v")?.Value, "Numeric quantity value is incorrect.");

        var six = Sheet("xl/worksheets/sheet2.xml");
        Contains("F-6", six.ToString(), "Site2 six-piece order is missing from the same workbook.");
        var nine = Sheet("xl/worksheets/sheet3.xml");
        Contains("R-9", nine.ToString(), "Nine-piece order is missing.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task ThreeGiftBoxWorkbookAppendPreservesConfirmation()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-three-box-append-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(string id, string orderNo, decimal quantity) =>
        new(id, at, "site1", "網站1－睿驛", orderNo, "parent:A", "蝦皮賣場",
            "2026/08/04 08:00:00", "2026-08-24", "4710964232435", "蛋黃酥小禮盒",
            0m, quantity, quantity, "新增", "F", "待處理", false, string.Empty);

    try
    {
        var first = Entry("ID-1", "ORDER-1", 2m);
        var second = Entry("ID-2", "ORDER-2", -1m);
        GiftBoxWorkbookExporter.Export(path, [first]);
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var sheet = archive.GetEntry("xl/worksheets/sheet1.xml")
                ?? throw new InvalidOperationException("Three-piece worksheet is missing.");
            XDocument document;
            using (var input = sheet.Open()) document = XDocument.Load(input);
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            document.Descendants(spreadsheet + "c")
                .Single(cell => (string?)cell.Attribute("r") == "E2")
                .Descendants(spreadsheet + "t").Single().Value = "已確認";
            sheet.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var output = replacement.Open();
            document.Save(output);
        }

        GiftBoxWorkbookExporter.Export(path, [first, second]);
        using var resultArchive = ZipFile.OpenRead(path);
        using var stream = (resultArchive.GetEntry("xl/worksheets/sheet1.xml")
                            ?? throw new InvalidOperationException("Three-piece worksheet is missing after append.")).Open();
        var result = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Equal("已確認", result.Descendants(ns + "c")
            .Single(cell => (string?)cell.Attribute("r") == "E2")
            .Descendants(ns + "t").Single().Value, "Manual confirmation was overwritten.");
        Equal("-1", result.Descendants(ns + "c")
            .Single(cell => (string?)cell.Attribute("r") == "D3")
            .Element(ns + "v")?.Value, "Appended negative quantity is not numeric.");
        Equal(1, result.ToString().Split("ID-1", StringSplitOptions.None).Length - 1,
            "Existing row was duplicated during append.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task GiftBoxSizeUsesEggYolkCountNotOtherMultipliers()
{
    Equal(3, GiftBoxSizePolicy.Resolve("4710964232435", "芝初黑芝麻Q潤蛋黃酥小禮盒"),
        "Three-piece SKU mapping is incorrect.");
    Equal(6, GiftBoxSizePolicy.Resolve(null,
        "預購_黑芝麻Q潤蛋黃酥禮盒(每盒含蛋黃酥x6)_8/22出貨"),
        "Six-piece parent name was not recognized.");
    Equal(9, GiftBoxSizePolicy.Resolve(null,
        "黑芝麻Q潤蛋黃酥大禮盒(內含9入蛋黃酥+芝麻粉7g/包 x3)"),
        "Sesame-powder multiplier incorrectly overrode the nine-piece count.");
    Equal(null, GiftBoxSizePolicy.Resolve(null, "黑芝麻Q潤蛋黃酥 單入包裝 蛋奶素"),
        "Single-piece item should not enter the three gift-box sheets.");
    return Task.CompletedTask;
}

static Task SourceSyncLockIsExclusive()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-lock-{Guid.NewGuid():N}");
    Directory.CreateDirectory(directory);
    try
    {
        using var first = SourceSyncFileLock.TryAcquire(directory, "site1")
            ?? throw new InvalidOperationException("First source lock should be acquired.");
        var second = SourceSyncFileLock.TryAcquire(directory, "site1");
        True(second is null, "A second lock for the same source should not be acquired.");
        second?.Dispose();

        using var otherSource = SourceSyncFileLock.TryAcquire(directory, "site2")
            ?? throw new InvalidOperationException("A different source should have an independent lock.");
        first.Dispose();
        using var afterRelease = SourceSyncFileLock.TryAcquire(directory, "site1")
            ?? throw new InvalidOperationException("Released source lock should be acquirable again.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }

    return Task.CompletedTask;
}

static Task FailedSyncKeepsCommittedPreviewValid()
{
    var synchronizedAt = new DateTimeOffset(2026, 8, 4, 9, 0, 0, TimeSpan.FromHours(8));
    var lines = new[]
    {
        new OrderLineSnapshot("site1", "R001", "parent:sku:BOX:1", "parent", null,
            "shopee", "店舖", "2026/08/04 08:00:00", null, null,
            "BOX", null, "蛋黃酥禮盒", null, 1m, 0m, "F", synchronizedAt)
    };
    var failedAt = synchronizedAt.AddDays(1);
    var failedStatus = new SyncSourceStatus(
        "site1", "網站1", new DateOnly(2026, 7, 1), new DateOnly(2026, 8, 5),
        "failed", failedAt.AddMinutes(-1), failedAt, 0, 0, synchronizedAt, "網路失敗");

    var preview = OrderSnapshotPreviewBuilder.Build(
        new WmsSource("site1", "網站1", new Uri("https://example.test")), lines, failedStatus);
    True(preview.IsSuccess, "A committed snapshot should remain valid after a later sync failure.");
    True(preview.Metadata.ResultOk, "Snapshot metadata should describe the committed data, not the failed run.");
    Equal(synchronizedAt, preview.Metadata.TestedAt, "Snapshot time should come from committed data.");
    return Task.CompletedTask;
}

static Task SnapshotPreviewRejectsOrphanItems()
{
    var synchronizedAt = DateTimeOffset.UtcNow;
    var lines = new OrderLineSnapshot[]
    {
        new("site1", "R001", "parent:sku:BOX:1", "parent", null,
            "shopee", null, "2026/08/04", null, null, "BOX", null, "禮盒", null, 1m, 0m, "F", synchronizedAt),
        new("site1", "R001", "item:missing:sku:ITEM:1", "item", "parent:sku:MISSING:1",
            "shopee", null, "2026/08/04", null, null, "ITEM", null, "子商品", null, 1m, 0m, "F", synchronizedAt)
    };

    try
    {
        OrderSnapshotPreviewBuilder.Build(
            new WmsSource("site1", "網站1", new Uri("https://example.test")), lines, null);
        throw new InvalidOperationException("Expected orphan snapshot item to be rejected.");
    }
    catch (InvalidDataException exception)
    {
        Contains("子商品", exception.Message, "Orphan-item failure should be understandable.");
    }

    return Task.CompletedTask;
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
