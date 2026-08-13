using System.Net;
using System.Net.Http.Headers;
using System.IO.Compression;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using MidAutumnGiftBox.Core;

var tests = new (string Name, Func<Task> Run)[]
{
    ("Workbook consolidates all products into detail, pending review, and native-formula statistics", WorkbookUsesThreeConsolidatedWorksheetsAndNativeFormulas),
    ("Excel 2013 statistics remain structurally valid at 421 detail rows", WorkbookStatisticsSupports421DetailRows),
    ("Workbook output preserves past, current, and future order dates", WorkbookOutputPreservesAllOrderDates),
    ("Excel location uses the channel mapping only when the existing manual location is blank", WorkbookLocationUsesChannelMappingWithoutOverwritingManualValue),
    ("Flavor excludes explicit non-page logistics orders", FlavorExcludesExplicitNonPageLogisticsOrders),
    ("Committed Flavor preview preserves safe operational diagnostics", CommittedFlavorPreviewPreservesSafeOperationalDiagnostics),
    ("Flavor operational diagnostics expose only safe fields", FlavorPreviewShowsOnlySafeOperationalDiagnostics),
    ("店舖預覽會取得 Token、使用 Bearer 並移除敏感資料", ShopPreviewUsesBearerAndRemovesSensitiveData),
    ("訂單預覽依成立日範圍取得蛋黃酥資料且不以指定到貨日排除", PendingOrderPreviewUsesOrderDateRange),
    ("WMS 商品以 item_no 優先保留且名稱條件仍供異常預覽", WmsProductFilterUsesItemNumberAndKeepsReviewFallbacks),
    ("睿驛訂單以 source_key 顯示通路並從商品 spec 推導出貨星期一", ReyiChannelAndShippingWindowUseVerifiedFields),
    ("WMS 訂單備註可在既有日期空白時推導日期並完整進入 Excel", WmsOrderNoteFallbackPersistsAndExports),
    ("商品與 API 日期優先於備註且備註衝突不猜日期", ExistingDatesPrecedeOrderNoteAndNoteConflictsNeedReview),
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
    ("最新完整快照可直接產生正式訂單而不依賴 Ledger", CurrentSnapshotProjectsCurrentOrdersWithoutLedger),
    ("WMS 無日期單入會繼承同訂單最早禮盒日期且孤立單入不匯出", WmsDependentSinglesUseEarliestGiftDate),
    ("兩網站完整快照只有整批有效時才一起取代", FullSnapshotReplacementIsAtomicAcrossSources),
    ("同 SKU 不同日期明細重排後仍維持相同識別值", SameSkuDifferentDateLinesKeepIdentityWhenReordered),
    ("人工修改四欄含空白都永久優先且可整筆清除", ManualOverridesPreserveBlanksAndClearByDetailRow),
    ("人工覆寫檔保存最近 Excel 路徑與匯出列基準", ManualOverrideStatePersistsWorkbookPathAndRows),
    ("正式 Excel 可讀回四個人工欄位與明細識別值", FormalWorkbookReadsBackEditableRows),
    ("清除覆寫後同路徑匯出會恢復自動地點與確認", ClearedOverrideRestoresAutomaticFieldsInExistingWorkbook),
    ("訂單快照遇到重複唯一鍵會拒絕整批資料", OrderSnapshotRejectsDuplicateUniqueKeys),
    ("訂單快照遇到無效數量會保留舊資料", OrderSnapshotRejectsInvalidQuantities),
    ("已提交快照會重建相同的畫面與 Excel 父子商品", CommittedSnapshotBuildsPreviewAndExcel),
    ("同步 checkpoint 不會縮短待處理訂單的完整成立日範圍", SyncWindowKeepsFullOrderDateRange),
    ("增量快照只取代來源的訂單成立日範圍", IncrementalSnapshotReplacesOnlyQueriedOrderDates),
    ("增量訂單 API 會送出明確的成立日起訖", ExplicitOrderRangeIsSentToWms),
    ("離開待處理的訂單可依訂單編號查回最終狀態", OrderStatusCanBeQueriedByOrderNumber),
    ("同步差異只以父商品數量產生追加紀錄", ParentQuantityChangesBecomeAppendOnlyEntries),
    ("組合父商品可由唯一目標子商品 SKU 分類且只計父商品數量", ParentUsesUniqueTargetChildSkuForSize),
    ("Ledger 日期會前移且衝突商品不使用原始到貨日", LedgerUsesAdjustedDatesAndLeavesConflictsBlank),
    ("離開待處理後只有取消退貨沖銷且刪除併單待人工確認", DepartedOrdersUseResolvedStatusRules),
    ("異動紀錄只追加且重試不會重複寫入", ChangeLedgerIsAppendOnlyAndIdempotent),
    ("Excel 追加新異動時會保留人工確認欄", ExcelAppendPreservesManualConfirmation),
    ("兩網站禮盒會逐訂單分列且數量為 Excel 數字", ThreeGiftBoxSheetsCombineSourcesAndUseNumericQuantities),
    ("正式 Excel 逐訂單輸出且不保存收件資料", PerOrderWorkbookExcludesRecipientAndPreservesManualFields),
    ("舊彙總 Excel 不會被逐訂單格式原地改寫", LegacyAggregateWorkbookIsRejected),
    ("訂單明細重新匯出時保留逐訂單確認", ThreeGiftBoxWorkbookAppendPreservesConfirmation),
    ("日期補齊後會從待確認表移入日期彙總表", GiftBoxWorkbookSplitDateGroupsPreserveConfirmation),
    ("日期補齊後不同訂單仍分列並保留既有確認", GiftBoxWorkbookMergedDateGroupsPreserveConfirmation),
    ("Excel 只回填空白日期且保留人工確認與歷史欄位", GiftBoxWorkbookBackfillsBlankDatesFromSnapshot),
    ("Excel 不會用衝突快照的原始到貨日回填", GiftBoxWorkbookDoesNotBackfillConflictingDates),
    ("日期待確認工作表只列最終無法判定日期的訂單", PendingDateWorksheetShowsOnlyUnresolvedOrders),
    ("日期待確認只回寫指定到貨日並自動移入訂單明細", PendingDateWorkbookWritesBackSpecifiedDate),
    ("Excel 同時顯示原定到貨日、下單日並保留人工地點", WorkbookShowsOriginalAndOrderDatesAndPreservesLocation),
    ("Excel 人工欄位含公式時拒絕覆寫原檔", WorkbookRejectsManualFieldFormulaWithoutOverwrite),
    ("Excel 三張工作表使用標楷體且不建立自動篩選", WorkbookUsesDfkaiAndHasNoAutoFilters),
    ("更新舊版 Excel 會補標楷體並移除既有篩選", LegacyWorkbookUpdateAddsDfkaiAndRemovesFilters),
    ("ERP 與手打單 Excel 依欄名載入並產生可匯出資料", SpreadsheetSourcesMapColumnsAndTargetSkus),
    ("自動載入會從啟動資料夾解析三個固定 Excel 檔名", StartupSpreadsheetInputsResolveFixedFileNames),
    ("百貨 POS 預購報表依完整通路與非空白貨號安全匯入", PosSpreadsheetMapsExactChannelsDatesProductsAndQuantities),
    ("POS 無預購日固定首次載入排程且跨日重載不漂移", PosMissingPickupDateUsesPersistentFirstImportSchedule),
    ("手打單兩種單入品名進入日期待確認且不互相合併", ManualSingleItemsUseSeparateWorksheet),
    ("舊異動名稱依蛋黃酥入數分流且不誤用芝麻粉倍數", GiftBoxSizeUsesEggYolkCountNotOtherMultipliers),
    ("同一來源的本機同步鎖同時間只能由一個程序取得", SourceSyncLockIsExclusive),
    ("清除歷史會先備份三個資料檔且可在重抓失敗時復原", LocalHistoryResetIsRecoverable),
    ("手打單會保存稽核與登記狀態並安全加入 Excel 資料", ManualOrdersPersistAuditAndExportSafely),
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

static Task StartupSpreadsheetInputsResolveFixedFileNames()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-startup-{Guid.NewGuid():N}");
    Directory.CreateDirectory(directory);
    try
    {
        var erpPath = Path.Combine(directory, "ERP產出數量.xlsx");
        var posPath = Path.Combine(directory, "百貨專櫃門市預購報表.xlsx");
        File.WriteAllBytes(erpPath, []);
        File.WriteAllBytes(posPath, []);

        var inputs = StartupSpreadsheetInputs.Resolve(directory);

        Equal(Path.GetFullPath(erpPath), inputs.ErpPath, "ERP fixed input path is incorrect.");
        Equal(null, inputs.ManualPath, "A missing manual workbook must remain absent.");
        Equal(Path.GetFullPath(posPath), inputs.PosPath, "POS fixed input path is incorrect.");
        Equal(Path.Combine(Path.GetFullPath(directory), "中秋禮盒訂單統計.xlsx"), inputs.OutputPath,
            "Quick workflow output path is incorrect.");
        True(!inputs.HasAllRequiredInputs, "A missing fixed workbook must make the quick workflow incomplete.");
        True(inputs.MissingFileNames.SequenceEqual(["蛋黃酥-數量.xlsx"]),
            "Missing fixed input names are incorrect.");

        File.WriteAllBytes(Path.Combine(directory, "蛋黃酥-數量.xlsx"), []);
        True(StartupSpreadsheetInputs.Resolve(directory).HasAllRequiredInputs,
            "All three fixed workbooks must make the quick workflow complete.");
        return Task.CompletedTask;
    }
    finally
    {
        Directory.Delete(directory, recursive: true);
    }
}

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
                    "receiver_phone":"0911222333",
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
    Equal("2026-08-06", result.Rows[0]["derived_shipping_date"],
        "An API arrival date should move back at least five days to the preceding Tuesday or Thursday.");
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
    DoesNotContain("0911222333", result.SafeJson, "Safe JSON leaked receiver phone.");
    True(!result.Rows[0].ContainsKey("receiver_name"), "Canonical projection retained encrypted receiver name.");
    True(!result.Rows[0].ContainsKey("receiver_phone"), "Canonical projection retained encrypted receiver phone.");
    True(!result.Rows[0].ContainsKey("address"), "Canonical projection retained receiver address.");

    var synchronizedAt = new DateTimeOffset(2026, 8, 20, 9, 0, 0, TimeSpan.FromHours(8));
    var snapshots = OrderSnapshotStore.BuildSourceSnapshot("site1", result.Rows, synchronizedAt);
    var snapshotJson = JsonSerializer.Serialize(snapshots);
    DoesNotContain("不應顯示", snapshotJson, "Snapshot retained encrypted receiver name.");
    DoesNotContain("0911222333", snapshotJson, "Snapshot retained encrypted receiver phone.");
    DoesNotContain("台北市不應顯示", snapshotJson, "Snapshot retained receiver address.");
    var changes = OrderChangePlanner.Plan(
        "網站1", [], snapshots, new Dictionary<string, OrderStatusResolution>(), synchronizedAt);
    var changeJson = JsonSerializer.Serialize(changes);
    DoesNotContain("不應顯示", changeJson, "Ledger retained encrypted receiver name.");
    DoesNotContain("0911222333", changeJson, "Ledger retained encrypted receiver phone.");
    DoesNotContain("台北市不應顯示", changeJson, "Ledger retained receiver address.");
}

static Task WmsProductFilterUsesItemNumberAndKeepsReviewFallbacks()
{
    using var targetByItemNo = JsonDocument.Parse("""
        {"sku":"UNKNOWN","name":"一般商品","items":[
          {"sku":"UNKNOWN-CHILD","item_no":"1902248","name":"來源名稱已修改","qty":1}
        ]}
        """);
    True(GiftBoxProductPolicy.IsRelevantProduct(targetByItemNo.RootElement),
        "A target child item number must keep the WMS product even when SKU and name no longer match.");

    using var unrelated = JsonDocument.Parse("""
        {"sku":"UNKNOWN","item_no":"9999999","name":"一般商品","items":[
          {"sku":"UNKNOWN-CHILD","item_no":"8888888","name":"其他商品","qty":1}
        ]}
        """);
    True(!GiftBoxProductPolicy.IsRelevantProduct(unrelated.RootElement),
        "A non-target item number without SKU or name evidence must remain excluded.");

    using var reviewFallback = JsonDocument.Parse(
        "{\"sku\":\"UNKNOWN\",\"name\":\"待補 item_no 的蛋黃酥商品\",\"qty\":1}");
    True(GiftBoxProductPolicy.IsRelevantProduct(reviewFallback.RootElement),
        "The name fallback must keep an item-number anomaly visible in preview for correction.");
    return Task.CompletedTask;
}

static async Task ReyiChannelAndShippingWindowUseVerifiedFields()
{
    Equal(new DateOnly(2026, 8, 6), ShippingLeadTimePolicy.Adjust(new DateOnly(2026, 8, 15)),
        "The user's 8/15 example should adjust to Thursday 8/6.");
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
    Equal("2026-08-13", row["derived_shipping_date"],
        "The in-range workday should move back five workdays and then align to Tuesday or Thursday.");
    Equal("product.spec", row["shipping_date_source"], "Shipping-date source should identify product spec.");
    Contains("8/22-8/28當週出貨", row["products"] ?? string.Empty, "Product spec should remain visible.");
    Contains("蝦皮賣場", result.SafeJson, "Safe JSON should contain the channel display name.");
    Contains("2026-08-13", result.SafeJson, "Safe JSON should contain the adjusted shipping date.");
}

static async Task WmsOrderNoteFallbackPersistsAndExports()
{
    const string note = "預計到貨日 8/29~9/4";
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
                "order_no":"SS2608040003",
                "order_date":"2026/08/04 17:22:00",
                "arrival_date":null,
                "source":"美安",
                "source_key":"ysshop",
                "status_code":"F",
                "note":"預計到貨日 8/29~9/4",
                "products":[{
                  "sku":"1902247-1",
                  "type":"shop",
                  "name":"芝初黑芝麻蛋黃酥9入禮盒",
                  "spec":"蛋黃酥9入",
                  "qty":8,
                  "shipp_qty":0
                }]
              }]}
            }
            """);
    });
    var source = new WmsSource("site1", "睿驛", new Uri("https://example.test"));
    var client = new WmsApiClient(new HttpClient(handler));
    var result = await client.GetPendingOrdersSinceInitialDateAsync(
        source,
        new ApiCredentials("id", "key"),
        new DateOnly(2026, 8, 6));

    Equal(note, result.Rows.Single()["note"], "API order note was not projected.");
    Equal("2026-08-29", result.Rows.Single()["ship_window_start"], "Note window start is incorrect.");
    Equal("2026-09-04", result.Rows.Single()["ship_window_end"], "Note window end is incorrect.");
    Equal("2026-08-20", result.Rows.Single()["derived_shipping_date"],
        "Note fallback did not apply the five-workday order-date rule.");
    Equal("order.note", result.Rows.Single()["shipping_date_source"], "Note source was not identified.");
    Contains(note, result.SafeJson, "Safe preview JSON should show the requested order note.");

    var synchronizedAt = new DateTimeOffset(2026, 8, 6, 14, 0, 0, TimeSpan.FromHours(8));
    var snapshot = OrderSnapshotStore.BuildSourceSnapshot(source.Code, result.Rows, synchronizedAt);
    Equal(note, snapshot.Single(line => line.LineLevel == "parent").Note,
        "Committed snapshot did not preserve the order note.");
    var committedPreview = OrderSnapshotPreviewBuilder.Build(source, snapshot, null);
    Equal(note, committedPreview.Rows.Single()["note"], "Committed preview did not restore the order note.");
    var changes = OrderChangePlanner.Plan(
        source.Name,
        Array.Empty<OrderLineSnapshot>(),
        snapshot,
        new Dictionary<string, OrderStatusResolution>(),
        synchronizedAt);
    Equal(note, changes.Single().Note, "Ledger entry did not preserve the order note.");

    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-note-{Guid.NewGuid():N}.xlsx");
    try
    {
        GiftBoxWorkbookExporter.Export(path, changes, snapshot);
        using var archive = ZipFile.OpenRead(path);
        using var stream = (archive.GetEntry("xl/worksheets/sheet1.xml")
                            ?? throw new InvalidOperationException("Order-detail worksheet is missing.")).Open();
        var sheet = XDocument.Load(stream).ToString();
        Contains("備註", sheet, "Gift-box worksheet is missing the note column.");
        Contains(note, sheet, "Gift-box worksheet did not preserve the order note.");
        Contains(ExcelSerial("2026-08-27"), sheet, "Note range did not produce the expected adjusted delivery date.");
        Contains(ExcelSerial("2026-08-20"), sheet, "Note range did not produce the expected order date.");
        Contains("<v>8</v>", sheet, "Gift-box quantity must remain the parent qty.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }
}

static Task ExistingDatesPrecedeOrderNoteAndNoteConflictsNeedReview()
{
    JsonObject Product(string spec = "") => new()
    {
        ["sku"] = "4710964232565",
        ["name"] = "2026 黑芝麻蛋黃酥9入禮盒",
        ["spec"] = spec,
        ["qty"] = 1
    };
    JsonObject Order(string note, string? arrivalDate = null) => new()
    {
        ["order_date"] = "2026/08/04 09:00:00",
        ["arrival_date"] = arrivalDate,
        ["note"] = note
    };

    var productWins = Order("預計到貨日 9/12~9/18");
    OrderPreviewEnricher.Enrich(productWins, new JsonArray(Product("8/22-8/28當週出貨")));
    Equal("product.spec", productWins["shipping_date_source"]?.GetValue<string>(),
        "Product shipping window must take precedence over the order note.");
    Equal("2026-08-13", productWins["derived_shipping_date"]?.GetValue<string>(),
        "Product shipping window was unexpectedly replaced by the note.");

    var arrivalWins = Order("預計到貨日 9/12~9/18", "2026/09/07");
    OrderPreviewEnricher.Enrich(arrivalWins, new JsonArray(Product()));
    Equal("order.arrival_date", arrivalWins["shipping_date_source"]?.GetValue<string>(),
        "API arrival_date must take precedence over the order note.");
    Equal("2026-08-27", arrivalWins["derived_shipping_date"]?.GetValue<string>(),
        "API arrival_date was unexpectedly replaced by the note.");

    var conflict = Order("預計到貨日 8/29~9/4；備用 9/5~9/11");
    OrderPreviewEnricher.Enrich(conflict, new JsonArray(Product()));
    Equal("conflict", conflict["shipping_date_status"]?.GetValue<string>(),
        "Multiple note ranges must require review.");
    True(conflict["derived_shipping_date"] is null,
        "Multiple note ranges must not guess a delivery date.");

    var invalid = Order("預計到貨日 8/29~9/xx");
    OrderPreviewEnricher.Enrich(invalid, new JsonArray(Product()));
    Equal("needs_review", invalid["shipping_date_status"]?.GetValue<string>(),
        "An invalid explicit note range must require review.");
    True(invalid["derived_shipping_date"] is null,
        "An invalid note range must not guess a delivery date.");
    return Task.CompletedTask;
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
                "arrival_date":"2026/08/30",
                "status_code":"F",
                "products":[
                  {"name":"2026 蛋黃酥禮盒A","spec":"8/22-8/28當週出貨","qty":1},
                  {"name":"2026 蛋黃酥禮盒B","spec":"8/29-9/4當週出貨","qty":1}
                ]
              },
              {
                "order_no":"SPEC-AMBIGUOUS",
                "order_date":"2026/08/04 11:00:00",
                "arrival_date":"2026/08/30",
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
        Equal("2026-08-13", firstProduct.GetProperty("derived_shipping_date").GetString(),
            "The first parent product should keep its own adjusted date.");
        Equal("2026-08-20", secondProduct.GetProperty("derived_shipping_date").GetString(),
            "The second parent product should keep its own adjusted date.");
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
    Equal("2026-09-08", weekdayStart["derived_shipping_date"],
        "A range without Monday should adjust its first in-range workday by the lead-time rule.");
    var weekendOnly = result.Rows.Single(row => row["order_no"] == "WEEKEND-ONLY");
    Equal("needs_review", weekendOnly["shipping_date_status"],
        "A weekend-only range should require review because it has no in-range workday.");
    True(!weekendOnly.ContainsKey("derived_shipping_date"),
        "A weekend-only range must not choose a date outside the range.");
    var crossYear = result.Rows.Single(row => row["order_no"] == "CROSS-YEAR");
    Equal("2026-12-29", crossYear["ship_window_start"], "Cross-year start is incorrect.");
    Equal("2027-01-04", crossYear["ship_window_end"], "Cross-year end is incorrect.");
    Equal("2026-12-24", crossYear["derived_shipping_date"], "Cross-year adjusted date is incorrect.");
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

static async Task FlavorExcludesExplicitNonPageLogisticsOrders()
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
              "data":{"maxpage":1,"rows":[
                {
                  "order_no":"WEB-VISIBLE",
                  "order_date":"2026/08/05 09:00:00",
                  "status_code":"F",
                  "logistics_code":"cat",
                  "products":[{"sku":"4710964232565","type":"shop","name":"芝初蛋黃酥9入禮盒","qty":1}]
                },
                {
                  "order_no":"WEB-HIDDEN-PREORDER",
                  "order_date":"2026/08/05 10:00:00",
                  "status_code":"F",
                  "logistics_code":" none ",
                  "products":[{"sku":"4710964232565","type":"shop","name":"芝初蛋黃酥9入禮盒","qty":1}]
                },
                {
                  "order_no":"UNKNOWN-CODE-REVIEWABLE",
                  "order_date":"2026/08/05 11:00:00",
                  "status_code":"F",
                  "logistics_code":"future-code",
                  "products":[{"sku":"4710964232565","type":"shop","name":"芝初蛋黃酥9入禮盒","qty":1}]
                }
              ]}
            }
            """);
    });

    var client = new WmsApiClient(new HttpClient(handler));
    var result = await client.GetPendingOrdersSinceInitialDateAsync(
        new WmsSource("site2", "Flavor", new Uri("https://example.test")),
        new ApiCredentials("id", "key"),
        new DateOnly(2026, 8, 5));

    Equal(2, result.Rows.Count, "Flavor should exclude only explicit logistics_code=none orders.");
    Contains("WEB-VISIBLE", result.SafeJson, "A Flavor web-page order was excluded.");
    Contains("UNKNOWN-CODE-REVIEWABLE", result.SafeJson,
        "An unknown logistics code should remain visible for safe review.");
    DoesNotContain("WEB-HIDDEN-PREORDER", result.SafeJson,
        "A Flavor non-page preorder remained in the pending preview.");

    var reyiResult = await client.GetPendingOrdersSinceInitialDateAsync(
        new WmsSource("site1", "Reyi", new Uri("https://reyi.example.test")),
        new ApiCredentials("id", "key"),
        new DateOnly(2026, 8, 5));
    Equal(3, reyiResult.Rows.Count,
        "Reyi must not apply Flavor's logistics_code=none exclusion rule.");
}

static async Task FlavorPreviewShowsOnlySafeOperationalDiagnostics()
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
                "order_no":"FLAVOR-DIAGNOSTIC",
                "order_date":"2026/08/05 10:16:25",
                "arrival_date":null,
                "source":"pos",
                "status_code":"F",
                "transfer_date":"TRANSFER-TIME-MUST-NOT-LEAK",
                "shipping_method":"SHIPPING-METHOD-MUST-NOT-LEAK",
                "change_type":"CHANGE-TYPE-MUST-NOT-LEAK",
                "is_preorder":true,
                "finish_time":"2026-08-05 10:16:25",
                "logistics":"cat",
                "logistics_code":"STORE_PREORDER",
                "AllPayLogisticsID":"ALLPAY-ID-MUST-NOT-LEAK",
                "transfer_debug_marker":"UNKNOWN-VALUE-MUST-NOT-LEAK",
                "receiver_name":"CUSTOMER-MUST-NOT-LEAK",
                "receiver_phone":"0900000000",
                "address":"ADDRESS-MUST-NOT-LEAK",
                "access_token":"TOKEN-MUST-NOT-LEAK",
                "products":[{
                  "sku":"4710964232565",
                  "type":"shop",
                  "name":"芝初黑芝麻Q潤蛋黃酥禮盒(9入)",
                  "qty":1,
                  "shipp_qty":1
                }]
              }]}
            }
            """);
    });

    var client = new WmsApiClient(new HttpClient(handler));
    var result = await client.GetPendingOrdersSinceInitialDateAsync(
        new WmsSource("site2", "Flavor", new Uri("https://example.test")),
        new ApiCredentials("id", "key"),
        new DateOnly(2026, 8, 5));

    Equal(1, result.Rows.Count, "Flavor diagnostic order should remain visible.");
    var row = result.Rows[0];
    Contains("finish_time=2026-08-05 10:16:25", row["operational_diagnostics"] ?? string.Empty,
        "Safe operational diagnostics should show finish_time.");
    Contains("logistics=cat", row["operational_diagnostics"] ?? string.Empty,
        "Safe operational diagnostics should show logistics.");
    Contains("logistics_code=STORE_PREORDER", row["operational_diagnostics"] ?? string.Empty,
        "Safe operational diagnostics should show logistics_code.");
    Contains("transfer_date", row["operational_diagnostic_fields"] ?? string.Empty,
        "Unverified operational fields should expose names only.");
    Contains("shipping_method", row["operational_diagnostic_fields"] ?? string.Empty,
        "Unverified shipping fields should expose names only.");
    Contains("change_type", row["operational_diagnostic_fields"] ?? string.Empty,
        "Unverified status fields should expose names only.");
    DoesNotContain("TRANSFER-TIME-MUST-NOT-LEAK", result.SafeJson,
        "Unverified transfer values must remain masked.");
    DoesNotContain("SHIPPING-METHOD-MUST-NOT-LEAK", result.SafeJson,
        "Unverified shipping values must remain masked.");
    DoesNotContain("CHANGE-TYPE-MUST-NOT-LEAK", result.SafeJson,
        "Unverified status values must remain masked.");
    DoesNotContain("ALLPAY-ID-MUST-NOT-LEAK", result.SafeJson,
        "Safe JSON must not expose unique logistics identifiers.");
    Contains("transfer_debug_marker", row["operational_diagnostic_fields"] ?? string.Empty,
        "Unknown operational fields should expose their names for diagnosis.");
    DoesNotContain("UNKNOWN-VALUE-MUST-NOT-LEAK", result.SafeJson,
        "Unknown operational field values must remain masked.");
    DoesNotContain("CUSTOMER-MUST-NOT-LEAK", result.SafeJson, "Safe JSON leaked receiver_name.");
    DoesNotContain("0900000000", result.SafeJson, "Safe JSON leaked receiver_phone.");
    DoesNotContain("ADDRESS-MUST-NOT-LEAK", result.SafeJson, "Safe JSON leaked address.");
    DoesNotContain("TOKEN-MUST-NOT-LEAK", result.SafeJson, "Safe JSON leaked access_token.");
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
        True(nestedItem.ParentLineKey?.StartsWith(
                 "parent:item_no:P01:detail:", StringComparison.OrdinalIgnoreCase) == true,
            "The nested item lost its stable parent line key.");
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
        True(site1.ExternalLineKey.StartsWith(
                "parent:item_no:P01:detail:", StringComparison.OrdinalIgnoreCase),
            "The stable parent line key is incorrect.");
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

static Task CommittedFlavorPreviewPreservesSafeOperationalDiagnostics()
{
    var synchronizedAt = new DateTimeOffset(2026, 8, 5, 15, 30, 0, TimeSpan.FromHours(8));
    IReadOnlyDictionary<string, string?> row = new Dictionary<string, string?>
    {
        ["order_no"] = "FLAVOR-DIAGNOSTIC",
        ["order_date"] = "2026/08/05 10:16:25",
        ["status_code"] = "F",
        ["operational_diagnostic_fields"] = "shipping_method, transfer_date, transfer_debug_marker",
        ["operational_diagnostics"] = "finish_time=null; logistics=cat; logistics_code=cat",
        ["products"] = """[{"sku":"4710964232565","type":"shop","name":"芝初黑芝麻Q潤蛋黃酥禮盒(9入)","qty":1,"shipp_qty":1}]"""
    };

    var lines = OrderSnapshotStore.BuildSourceSnapshot("site2", [row], synchronizedAt);
    var preview = OrderSnapshotPreviewBuilder.Build(
        new WmsSource("site2", "Flavor", new Uri("https://example.test")), lines, null);

    Equal(1, preview.Rows.Count, "Committed diagnostics should rebuild one order row.");
    Contains("transfer_debug_marker", preview.Rows[0]["operational_diagnostic_fields"] ?? string.Empty,
        "Committed preview lost diagnostic field names.");
    Contains("logistics_code=cat", preview.Rows[0]["operational_diagnostics"] ?? string.Empty,
        "Committed preview lost safe operational values.");
    Contains("logistics_code=cat", preview.SafeJson,
        "Committed safe JSON lost operational diagnostics.");
    return Task.CompletedTask;
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

static Task ParentUsesUniqueTargetChildSkuForSize()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-child-size-{Guid.NewGuid():N}.xlsx");
    var synchronizedAt = new DateTimeOffset(2026, 8, 6, 9, 0, 0, TimeSpan.FromHours(8));
    OrderLineSnapshot Parent(string orderNo, decimal quantity) =>
        new("site1", orderNo, "parent:A", "parent", null,
            "shopee", "賣場", "2026/08/04 08:00:00", "2026/08/18", null,
            "COMBINE-UNKNOWN", null, "中秋組合商品", null, quantity, 0m, "F", synchronizedAt,
            "蝦皮賣場");
    OrderLineSnapshot Child(string orderNo, string lineKey, string sku, decimal quantity) =>
        new("site1", orderNo, lineKey, "item", "parent:A",
            "shopee", "賣場", "2026/08/04 08:00:00", "2026/08/18", null,
            sku, null, "蛋黃酥主商品", null, quantity, 0m, "F", synchronizedAt,
            "蝦皮賣場");
    var current = new[]
    {
        Parent("UNIQUE-CHILD", 2m),
        Child("UNIQUE-CHILD", "item:parent:A:target", "4710964232411", 12m),
        Parent("AMBIGUOUS-CHILDREN", 3m),
        Child("AMBIGUOUS-CHILDREN", "item:parent:A:three", "4710964232435", 9m),
        Child("AMBIGUOUS-CHILDREN", "item:parent:A:nine", "4710964232565", 27m)
    };
    var plannedEntries = OrderChangePlanner.Plan(
        "網站1－睿驛",
        Array.Empty<OrderLineSnapshot>(),
        current,
        new Dictionary<string, OrderStatusResolution>(),
        synchronizedAt);
    var entries = plannedEntries
        .Select(entry => entry with { GiftBoxSize = null })
        .ToArray();

    try
    {
        GiftBoxWorkbookExporter.Export(path, entries, current);
        using var archive = ZipFile.OpenRead(path);
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using var stream = (archive.GetEntry("xl/worksheets/sheet1.xml")
                            ?? throw new InvalidOperationException("Order-detail worksheet is missing.")).Open();
        var six = XDocument.Load(stream);
        Equal(2, six.Descendants(spreadsheet + "row").Count(),
            "Only the parent with one unique target child SKU should enter the six-piece sheet.");
        Contains("UNIQUE-CHILD", six.ToString(), "Unique child SKU did not classify its parent.");
        DoesNotContain("AMBIGUOUS-CHILDREN", six.ToString(),
            "Different target child SKUs must not guess a gift-box size.");
        var quantity = six.Descendants(spreadsheet + "c")
            .Single(cell => (string?)cell.Attribute("r") == "E2")
            .Element(spreadsheet + "v")?.Value;
        Equal("2", quantity, "Gift-box quantity must use parent qty, not child qty.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task LedgerUsesAdjustedDatesAndLeavesConflictsBlank()
{
    var synchronizedAt = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    OrderLineSnapshot Line(
        string orderNo,
        string? shippingDateStatus,
        string? derivedShippingDate = null,
        string arrivalDate = "2026/08/15") =>
        new("site1", orderNo, "parent:A", "parent", null,
            "shopee", "賣場", "2026/08/04 08:00:00", arrivalDate, derivedShippingDate,
            "4710964232435", null, "蛋黃酥3入禮盒", null, 1m, 0m, "F", synchronizedAt,
            "蝦皮賣場", ShippingDateStatus: shippingDateStatus);
    var current = new[]
    {
        Line("ARRIVAL-FALLBACK", null),
        Line("FIVE-WORKDAYS", null, arrivalDate: "2026/08/18"),
        Line("CONFLICT", "conflict", "2099-01-01"),
        Line("NEEDS-REVIEW", "needs_review")
    };

    var entries = OrderChangePlanner.Plan(
        "網站1－睿驛",
        Array.Empty<OrderLineSnapshot>(),
        current,
        new Dictionary<string, OrderStatusResolution>(),
        synchronizedAt);

    Equal("2026-08-06", entries.Single(entry => entry.ExternalOrderNo == "ARRIVAL-FALLBACK").DeliveryDate,
        "A raw API arrival date must be adjusted before entering the ledger.");
    Equal("2026-08-13", entries.Single(entry => entry.ExternalOrderNo == "ARRIVAL-FALLBACK").OriginalDeliveryDate,
        "The API arrival date must move back to the nearest Tuesday or Thursday.");
    Equal("2026-08-11", entries.Single(entry => entry.ExternalOrderNo == "FIVE-WORKDAYS").DeliveryDate,
        "Order date must move back five weekdays before aligning to Tuesday or Thursday.");
    Equal(null, entries.Single(entry => entry.ExternalOrderNo == "CONFLICT").DeliveryDate,
        "A conflicting product date must not fall back to the raw API arrival date.");
    Equal(null, entries.Single(entry => entry.ExternalOrderNo == "NEEDS-REVIEW").DeliveryDate,
        "A date needing review must not fall back to the raw API arrival date.");
    return Task.CompletedTask;
}

static Task DepartedOrdersUseResolvedStatusRules()
{
    const string diagnosticMarker = "DIAGNOSTIC-MUST-NOT-ENTER-LEDGER";
    var synchronizedAt = new DateTimeOffset(2026, 8, 4, 9, 0, 0, TimeSpan.FromHours(8));
    OrderLineSnapshot Line(string orderNo, decimal quantity, string? operationalDiagnostics = null) =>
        new("site1", orderNo, "parent:A", "parent", null,
            "shopee", "賣場", "2026/08/04 08:00:00", null, "2026-08-24",
            "SKU-" + orderNo, null, orderNo + " 商品", null, quantity, 0m, "F", synchronizedAt,
            "蝦皮賣場", OperationalDiagnostics: operationalDiagnostics);
    var previous = new[]
    {
        Line("CANCEL", 2m),
        Line("RETURN", 3m),
        Line("DELETE", 4m),
        Line("SHIPPED", 5m),
        Line("FLAVOR-NONE", 6m, diagnosticMarker)
    };
    var statuses = new Dictionary<string, OrderStatusResolution>(StringComparer.OrdinalIgnoreCase)
    {
        ["CANCEL"] = new("CANCEL", "N", "已取消"),
        ["RETURN"] = new("RETURN", "R", "已退貨"),
        ["DELETE"] = new("DELETE", "D", "已封存"),
        ["SHIPPED"] = new("SHIPPED", "S", "已出貨"),
        ["FLAVOR-NONE"] = new("FLAVOR-NONE", "F", "待處理")
    };

    var entries = OrderChangePlanner.Plan(
        "網站1－睿驛", previous, Array.Empty<OrderLineSnapshot>(), statuses,
        synchronizedAt.AddDays(1));

    Equal(4, entries.Count, "Shipped orders must not create a cancellation adjustment.");
    var cancelled = entries.Single(entry => entry.ExternalOrderNo == "CANCEL");
    Equal(-2m, cancelled.QuantityChange, "Cancelled order should reverse the remaining parent quantity.");
    Equal("取消", cancelled.ChangeType, "Cancelled order should be labelled as cancellation.");
    var returned = entries.Single(entry => entry.ExternalOrderNo == "RETURN");
    Equal(-3m, returned.QuantityChange, "Returned order should reverse the remaining parent quantity.");
    var deleted = entries.Single(entry => entry.ExternalOrderNo == "DELETE");
    Equal(0m, deleted.QuantityChange, "Deleted or merged order must not change quantity automatically.");
    True(deleted.NeedsReview, "Deleted or merged order should be marked for manual review.");
    var filteredFlavor = entries.Single(entry => entry.ExternalOrderNo == "FLAVOR-NONE");
    Equal(-6m, filteredFlavor.QuantityChange,
        "A previously committed Flavor none order should append a reversing entry after exclusion.");
    Equal("品項移除", filteredFlavor.ChangeType,
        "A still-F Flavor none order should be labelled as a removed item.");
    DoesNotContain(diagnosticMarker, JsonSerializer.Serialize(entries),
        "Operational diagnostics must not enter the append-only ledger projection.");
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
        string productName, decimal quantity, string channel = "蝦皮賣場", string date = "2026-08-18") =>
        new(id, at, sourceCode, sourceName, orderNo, "parent:A", channel,
            "2026/08/04 08:00:00", date, sku, productName,
            0m, quantity, quantity, "新增", "F", "待處理", false, string.Empty);
    var entries = new[]
    {
        Entry("THREE", "site1", "網站1－睿驛", "R-3", "4710964232435", "蛋黃酥小禮盒", 2m),
        Entry("THREE-2", "site2", "網站2－Flavor", "F-3", "4710964232435", "蛋黃酥小禮盒", 3m,
            date: "2026/08/18"),
        Entry("THREE-CANCEL", "site1", "網站1－睿驛", "R-3-CANCEL", "4710964232435", "蛋黃酥小禮盒", -1m),
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
            Contains("訂單明細", workbookXml, "Order-detail worksheet is missing.");
            Contains("日期待確認", workbookXml, "Pending-review worksheet is missing.");
            Contains("統計", workbookXml, "Statistics worksheet is missing.");
            DoesNotContain("訂單預覽", workbookXml, "Legacy preview worksheet should not remain.");
        }

        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XDocument Sheet(string name)
        {
            using var stream = (archive.GetEntry(name)
                                ?? throw new InvalidOperationException($"Worksheet {name} is missing.")).Open();
            return XDocument.Load(stream);
        }

        var details = Sheet("xl/worksheets/sheet1.xml");
        Equal(6, details.Descendants(spreadsheet + "row").Count(),
            "Every supported gift-box order must remain a separate detail row.");
        Contains("R-3", details.ToString(), "Site1 order number is missing from per-order output.");
        Contains("F-3", details.ToString(), "Site2 order number is missing from per-order output.");
        DoesNotContain("R-1", details.ToString(), "An unsupported single item must not enter order details.");
        var quantityCell = details.Descendants(spreadsheet + "c")
            .Single(cell => (string?)cell.Attribute("r") == "E2");
        Equal(null, (string?)quantityCell.Attribute("t"), "Quantity cell must not be stored as text.");
        True(details.Descendants(spreadsheet + "c")
                .Where(cell => ((string?)cell.Attribute("r"))?.StartsWith("E", StringComparison.Ordinal) == true &&
                               (string?)cell.Attribute("r") != "E1")
                .Select(cell => cell.Element(spreadsheet + "v")?.Value)
                .ToHashSet().SetEquals(["2", "3", "-1", "4"]),
            "Per-order numeric quantities are incorrect.");
        Contains("3入禮盒紙盒(黑金袋)", details.ToString(), "Three-piece name is missing.");
        Contains("6入鐵盒2026年鐵盒(黑金袋)", details.ToString(), "Six-piece name is missing.");
        Contains("9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)", details.ToString(), "Nine-piece name is missing.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task PerOrderWorkbookExcludesRecipientAndPreservesManualFields()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-per-order-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 6, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(
        string id, string source, string orderNo, string lineKey, decimal quantity,
        string originalDate) =>
        new(id, at, source, source, orderNo, lineKey, "蝦皮賣場",
            "2026/08/04 08:00:00", ShippingLeadTimePolicy.Adjust(
                DateOnly.ParseExact(originalDate, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture))
                .ToString("yyyy-MM-dd"),
            "4710964232435", "蛋黃酥三入禮盒", 0m, quantity, quantity,
            "新增", "F", "待處理", false, string.Empty, originalDate);
    var entries = new[]
    {
        Entry("A1", "site1", "ORDER-001", "parent:A", 2m, "2026-08-18"),
        Entry("A2", "site1", "ORDER-001", "parent:B", 3m, "2026-08-18"),
        Entry("B1", "site2", "ORDER-001", "parent:A", 4m, "2026-08-18"),
        Entry("C1", "site1", "ORDER-001", "parent:C", 1m, "2026-08-25")
    };

    try
    {
        GiftBoxWorkbookExporter.Export(path, entries);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using (var archive = ZipFile.OpenRead(path))
        using (var input = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open())
        {
            var sheet = XDocument.Load(input);
            var rows = sheet.Descendants(ns + "row").ToArray();
            Equal(4, rows.Length, "Source + order number + original delivery date must define one visible row.");
            var expectedHeaders = new[]
            {
                "通路別", "品名", "指定到貨日", "下單日", "數量", "確認", "地點", "訂單編號", "備註"
            };
            var actualHeaders = rows[0].Elements(ns + "c").Take(9)
                .Select(cell => string.Concat(cell.Descendants(ns + "t").Select(text => text.Value)))
                .ToArray();
            True(expectedHeaders.SequenceEqual(actualHeaders), "Gift-box visible columns are not in the agreed order.");
            var quantityCells = rows.Skip(1).Select(row => row.Elements(ns + "c")
                .Single(cell => ((string?)cell.Attribute("r"))?.StartsWith("E", StringComparison.Ordinal) == true))
                .ToArray();
            True(quantityCells.All(cell => (string?)cell.Attribute("t") is null),
                "Per-order quantities must be Excel numbers.");
            True(quantityCells.Select(cell => cell.Element(ns + "v")?.Value).ToHashSet().SetEquals(["5", "4", "1"]),
                "Per-order quantities were not grouped correctly.");
            Contains("ORDER-001", sheet.ToString(), "Order number is missing from gift-box output.");
            DoesNotContain("收件人", sheet.ToString(), "Recipient column must not remain in gift-box output.");
        }

        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument sheet;
            using (var input = entry.Open()) sheet = XDocument.Load(input);
            var row = sheet.Descendants(ns + "row").Single(item => (string?)item.Attribute("r") == "2");
            void SetInline(string column, string value)
            {
                var cell = row.Elements(ns + "c").Single(item =>
                    ((string?)item.Attribute("r"))?.StartsWith(column, StringComparison.Ordinal) == true);
                cell.SetAttributeValue("t", "inlineStr");
                cell.Element(ns + "v")?.Remove();
                cell.Add(new XElement(ns + "is", new XElement(ns + "t", value)));
            }
            SetInline("F", "已確認");
            SetInline("G", "台中廠");
            entry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var output = replacement.Open();
            sheet.Save(output);
        }

        GiftBoxWorkbookExporter.Export(path, entries);
        using var updated = ZipFile.OpenRead(path);
        using var updatedInput = updated.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var updatedSheet = XDocument.Load(updatedInput).ToString();
        Contains("已確認", updatedSheet, "Manual confirmation was not preserved for the same order row.");
        Contains("台中廠", updatedSheet, "Manual location was not preserved for the same order row.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task LegacyAggregateWorkbookIsRejected()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-legacy-aggregate-{Guid.NewGuid():N}.xlsx");
    var entry = new OrderChangeEntry(
        "LEGACY", DateTimeOffset.Now, "site1", "網站1", "ORDER", "parent:A", "蝦皮賣場", null,
        "2026-08-06", "4710964232435", "蛋黃酥3入禮盒", 0m, 1m, 1m,
        "新增", "F", "待處理", false, OriginalDeliveryDate: "2026-08-15");
    try
    {
        GiftBoxWorkbookExporter.Export(path, [entry]);
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var workbookEntry = archive.GetEntry("xl/workbook.xml")!;
            XDocument document;
            using (var input = workbookEntry.Open()) document = XDocument.Load(input);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var sheets = document.Descendants(ns + "sheets").Single();
            sheets.ReplaceNodes(
                new XElement(ns + "sheet", new XAttribute("name", "三入"), new XAttribute("sheetId", "1")),
                new XElement(ns + "sheet", new XAttribute("name", "六入"), new XAttribute("sheetId", "2")),
                new XElement(ns + "sheet", new XAttribute("name", "九入"), new XAttribute("sheetId", "3")),
                new XElement(ns + "sheet", new XAttribute("name", "日期待確認"), new XAttribute("sheetId", "4")),
                new XElement(ns + "sheet", new XAttribute("name", "單入"), new XAttribute("sheetId", "5")));
            workbookEntry.Delete();
            var replacement = archive.CreateEntry("xl/workbook.xml");
            using var output = replacement.Open();
            document.Save(output);
        }
        var before = File.ReadAllBytes(path);

        try
        {
            GiftBoxWorkbookExporter.Export(path, [entry]);
            throw new InvalidOperationException("Expected a legacy aggregate workbook to be rejected.");
        }
        catch (InvalidDataException exception)
        {
            Contains("另存新檔", exception.Message, "Legacy workbook rejection should tell the user to save a new file.");
        }

        True(before.SequenceEqual(File.ReadAllBytes(path)), "Rejecting a legacy workbook must leave it unchanged.");
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
            "2026/08/04 08:00:00", "2026-08-18", "4710964232435", "蛋黃酥小禮盒",
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
                .Single(cell => (string?)cell.Attribute("r") == "F2")
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
            .Single(cell => (string?)cell.Attribute("r") == "F2")
            .Descendants(ns + "t").Single().Value, "Manual confirmation was overwritten.");
        Equal("2", result.Descendants(ns + "c")
            .Single(cell => (string?)cell.Attribute("r") == "E2")
            .Element(ns + "v")?.Value, "The first order quantity changed unexpectedly.");
        Equal(3, result.Descendants(ns + "row").Count(),
            "A different order must be written as another row.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task GiftBoxWorkbookSplitDateGroupsPreserveConfirmation()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-date-split-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(string id, string orderNo, decimal quantity) =>
        new(id, at, "site1", "網站1－睿驛", orderNo, "parent:A", "蝦皮賣場",
            "2026/08/04 08:00:00", null, "4710964232435", "蛋黃酥3入禮盒",
            0m, quantity, quantity, "新增", "F", "待處理", false, string.Empty);
    OrderLineSnapshot Snapshot(string orderNo, string deliveryDate) =>
        new("site1", orderNo, "parent:A", "parent", null, "shopee", null,
            "2026/08/04 08:00:00", null, deliveryDate, "4710964232435", null,
            "蛋黃酥3入禮盒", null, 1m, 0m, "F", at.AddHours(1));

    try
    {
        var entries = new[] { Entry("SPLIT-1", "ORDER-1", 1m), Entry("SPLIT-2", "ORDER-2", 2m) };
        GiftBoxWorkbookExporter.Export(path, entries);
        using (var archive = ZipFile.OpenRead(path))
        {
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            using var datedInput = (archive.GetEntry("xl/worksheets/sheet1.xml")
                                    ?? throw new InvalidOperationException("Three-piece worksheet is missing.")).Open();
            Equal(1, XDocument.Load(datedInput).Descendants(ns + "row").Count(),
                "Undated orders must not enter a dated gift-box sheet.");
            using var pendingInput = (archive.GetEntry("xl/worksheets/sheet2.xml")
                                      ?? throw new InvalidOperationException("Pending-date worksheet is missing.")).Open();
            Equal(3, XDocument.Load(pendingInput).Descendants(ns + "row").Count(),
                "Both undated orders should begin in the pending-date worksheet.");
        }

        GiftBoxWorkbookExporter.Export(
            path,
            [Entry("RESET-SPLIT-1", "ORDER-1", 1m), Entry("RESET-SPLIT-2", "ORDER-2", 2m)],
            [Snapshot("ORDER-1", "2026-08-18"), Snapshot("ORDER-2", "2026-08-20")]);

        using var resultArchive = ZipFile.OpenRead(path);
        using var stream = (resultArchive.GetEntry("xl/worksheets/sheet1.xml")
                            ?? throw new InvalidOperationException("Three-piece worksheet is missing after split.")).Open();
        var result = XDocument.Load(stream);
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var dataRows = result.Descendants(spreadsheet + "row").Skip(1).ToArray();
        Equal(2, dataRows.Length, "Two resolved delivery dates should split the former blank-date aggregate.");
        True(dataRows.All(row => row.Elements(spreadsheet + "c")
                .Where(cell => ((string?)cell.Attribute("r"))?.StartsWith("F", StringComparison.OrdinalIgnoreCase) == true)
                .SelectMany(cell => cell.Descendants(spreadsheet + "t"))
                .All(text => string.IsNullOrEmpty(text.Value))),
            "Newly resolved aggregates should start with blank confirmation.");
        using var pendingResultStream = (resultArchive.GetEntry("xl/worksheets/sheet2.xml")
                                         ?? throw new InvalidOperationException("Pending-date worksheet is missing after resolution.")).Open();
        Equal(1, XDocument.Load(pendingResultStream).Descendants(spreadsheet + "row").Count(),
            "Resolved orders should leave the pending-date worksheet.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task GiftBoxWorkbookMergedDateGroupsPreserveConfirmation()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-date-merge-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(string id, string orderNo, string? deliveryDate) =>
        new(id, at, "site1", "網站1－睿驛", orderNo, "parent:A", "蝦皮賣場",
            "2026/08/04 08:00:00", deliveryDate, "4710964232435", "蛋黃酥3入禮盒",
            0m, 1m, 1m, "新增", "F", "待處理", false, string.Empty);
    OrderLineSnapshot Snapshot(string orderNo, string deliveryDate) =>
        new("site1", orderNo, "parent:A", "parent", null, "shopee", null,
            "2026/08/04 08:00:00", null, deliveryDate, "4710964232435", null,
            "蛋黃酥3入禮盒", null, 1m, 0m, "F", at.AddHours(1));

    try
    {
        var entries = new[]
        {
            Entry("MERGE-DATED", "ORDER-DATED", "2026-08-18"),
            Entry("MERGE-BLANK", "ORDER-BLANK", null)
        };
        GiftBoxWorkbookExporter.Export(path, entries);
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var sheet = archive.GetEntry("xl/worksheets/sheet1.xml")
                ?? throw new InvalidOperationException("Three-piece worksheet is missing.");
            XDocument document;
            using (var input = sheet.Open()) document = XDocument.Load(input);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            document.Descendants(ns + "c")
                .Single(cell => (string?)cell.Attribute("r") == "F2")
                .Descendants(ns + "t").Single().Value = "已確認";
            sheet.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var output = replacement.Open();
            document.Save(output);
        }

        GiftBoxWorkbookExporter.Export(
            path,
            entries,
            [Snapshot("ORDER-BLANK", "2026-08-18")]);

        using var resultArchive = ZipFile.OpenRead(path);
        using var stream = (resultArchive.GetEntry("xl/worksheets/sheet1.xml")
                            ?? throw new InvalidOperationException("Three-piece worksheet is missing after merge.")).Open();
        var result = XDocument.Load(stream);
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var dataRows = result.Descendants(spreadsheet + "row").Skip(1).ToArray();
        Equal(2, dataRows.Length, "Different orders must remain separate after resolving to the same date.");
        Contains("已確認", dataRows.Single(row => row.ToString().Contains("ORDER-DATED", StringComparison.Ordinal)).ToString(),
            "Resolving another order must preserve the existing order confirmation.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task GiftBoxWorkbookBackfillsBlankDatesFromSnapshot()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-date-backfill-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(string id, string orderNo, string? deliveryDate) =>
        new(id, at, "site1", "網站1：睿驛", orderNo, "parent:A", "蝦皮賣場",
            "2026/08/04 08:00:00", deliveryDate, "4710964232435", "蛋黃酥3入禮盒",
            0m, 2m, 2m, "新增", "F", "待處理", false, string.Empty);
    OrderLineSnapshot Snapshot(string orderNo, string arrivalDate, string deliveryDate) =>
        new("site1", orderNo, "parent:A", "parent", null, "shopee", null,
            "2026/08/04 08:00:00", arrivalDate, deliveryDate, "4710964232435", null,
            "蛋黃酥3入禮盒", null, 2m, 0m, "F", at.AddHours(1),
            OperationalDiagnostics: "DIAGNOSTIC-MUST-NOT-ENTER-EXCEL");

    try
    {
        var blankDate = Entry("BACKFILL-1", "ORDER-1", null);
        var existingDate = Entry("BACKFILL-2", "ORDER-2", "2026-08-10");
        GiftBoxWorkbookExporter.Export(path, [blankDate, existingDate]);

        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var sheet = archive.GetEntry("xl/worksheets/sheet1.xml")
                ?? throw new InvalidOperationException("Three-piece worksheet is missing.");
            XDocument document;
            using (var input = sheet.Open()) document = XDocument.Load(input);
            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            document.Descendants(spreadsheet + "c")
                .Single(cell => (string?)cell.Attribute("r") == "F2")
                .Descendants(spreadsheet + "t").Single().Value = "已確認";
            sheet.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var output = replacement.Open();
            document.Save(output);
        }

        GiftBoxWorkbookExporter.Export(
            path,
            [blankDate, existingDate],
            [Snapshot("ORDER-1", "2026/08/31", "2026-08-20"), Snapshot("ORDER-2", "2026/09/14", "2026-09-03")]);

        using var resultArchive = ZipFile.OpenRead(path);
        using var stream = (resultArchive.GetEntry("xl/worksheets/sheet1.xml")
                            ?? throw new InvalidOperationException("Three-piece worksheet is missing after backfill.")).Open();
        var result = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        string Text(string reference) => result.Descendants(ns + "c")
            .Single(cell => (string?)cell.Attribute("r") == reference)
            .Descendants(ns + "t").FirstOrDefault()?.Value ?? string.Empty;
        string CellText(XElement row, string column) => string.Concat(row.Elements(ns + "c")
            .Single(cell => ((string?)cell.Attribute("r"))?.StartsWith(column, StringComparison.Ordinal) == true)
            .Descendants(ns + "t").Select(text => text.Value));
        var backfilledRow = result.Descendants(ns + "row").Single(row => CellText(row, "H") == "ORDER-1");
        var backfilledRowNumber = (string?)backfilledRow.Attribute("r")
            ?? throw new InvalidOperationException("Backfilled aggregate row has no row number.");
        var existingRow = result.Descendants(ns + "row").Single(row => CellText(row, "H") == "ORDER-2");
        var existingRowNumber = (string?)existingRow.Attribute("r")
            ?? throw new InvalidOperationException("Existing aggregate row has no row number.");
        Equal(ExcelSerial("2026-09-03"), existingRow.Elements(ns + "c")
            .Single(cell => (string?)cell.Attribute("r") == $"D{existingRowNumber}").Element(ns + "v")?.Value,
            "The new date policy must replace a stale ledger order date from the latest snapshot.");
        Equal("已確認", Text($"F{existingRowNumber}"),
            "The existing dated order confirmation was overwritten while applying the new date policy.");
        Equal(string.Empty, Text($"F{backfilledRowNumber}"),
            "A newly resolved order should begin with blank confirmation.");
        Equal("2", result.Descendants(ns + "c").Single(cell =>
                (string?)cell.Attribute("r") == $"E{backfilledRowNumber}")
            .Element(ns + "v")?.Value, "Quantity changed while backfilling a date.");
        Contains("ORDER-1", result.ToString(), "Individual order number is missing from per-order Excel.");
        DoesNotContain("DIAGNOSTIC-MUST-NOT-ENTER-EXCEL", result.ToString(),
            "Operational diagnostics must not enter the formal gift-box workbook.");
        Equal(null, blankDate.DeliveryDate, "Backfilling Excel must not mutate the append-only ledger entry.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task GiftBoxWorkbookDoesNotBackfillConflictingDates()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-conflict-date-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    var entry = new OrderChangeEntry(
        "CONFLICT-ENTRY", at, "site1", "網站1－睿驛", "CONFLICT-ORDER", "parent:A", "蝦皮賣場",
        "2026/08/04 08:00:00", null, "4710964232435", "蛋黃酥3入禮盒",
        0m, 1m, 1m, "新增", "F", "待處理", false, string.Empty);
    var snapshot = new OrderLineSnapshot(
        "site1", "CONFLICT-ORDER", "parent:A", "parent", null, "shopee", null,
        "2026/08/04 08:00:00", "2026/08/15", "2099-01-01", "4710964232435", null,
        "蛋黃酥3入禮盒", null, 1m, 0m, "F", at.AddHours(1),
        ShippingDateStatus: "conflict");

    try
    {
        GiftBoxWorkbookExporter.Export(path, [entry], [snapshot]);
        using var archive = ZipFile.OpenRead(path);
        using var stream = (archive.GetEntry("xl/worksheets/sheet1.xml")
                            ?? throw new InvalidOperationException("Three-piece worksheet is missing.")).Open();
        var xml = XDocument.Load(stream).ToString();
        DoesNotContain("2026-08-06", xml,
            "A conflicting snapshot must not adjust and backfill its raw API arrival date.");
        DoesNotContain("2099-01-01", xml,
            "A conflicting snapshot must not use a stale derived date.");
        Equal(null, entry.DeliveryDate, "Excel conflict handling must not mutate the ledger entry.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task PendingDateWorksheetShowsOnlyUnresolvedOrders()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-pending-date-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(
        string id,
        string orderNo,
        decimal quantity,
        string? deliveryDate = null,
        string? note = null) =>
        new(id, at, "site1", "網站1－睿驛", orderNo, $"parent:{id}", "蝦皮賣場",
            "2026/08/04 08:00:00", deliveryDate, "4710964232435", "蛋黃酥3入禮盒",
            0m, quantity, quantity, "新增", "F", "待處理", false, string.Empty,
            Note: note);
    OrderLineSnapshot Snapshot(
        string orderNo,
        string lineKey,
        string? arrivalDate,
        string? derivedDate,
        string? status = null) =>
        new("site1", orderNo, lineKey, "parent", null, "shopee", null,
            "2026/08/04 08:00:00", arrivalDate, derivedDate, "4710964232435", null,
            "蛋黃酥3入禮盒", null, 1m, 0m, "F", at.AddHours(1),
            ShippingDateStatus: status);
    var entries = new[]
    {
        Entry("U1-A", "UNRESOLVED-1", 2m, note: "待確認日期備註"),
        Entry("U1-B", "UNRESOLVED-1", -1m, note: "待確認日期備註"),
        Entry("DATED", "DATED-ORDER", 3m, "2026-08-18"),
        Entry("FILLED", "FILLED-BY-SNAPSHOT", 4m),
        Entry("CONFLICT", "CONFLICT-ORDER", 5m, "2026-08-18"),
        Entry("ZERO-A", "ZERO-ORDER", 1m),
        Entry("ZERO-B", "ZERO-ORDER", -1m)
    };
    var snapshots = new[]
    {
        Snapshot("FILLED-BY-SNAPSHOT", "parent:FILLED", null, "2026-08-20"),
        Snapshot("CONFLICT-ORDER", "parent:CONFLICT", "2026/08/15", "2099-01-01", "conflict")
    };

    try
    {
        GiftBoxWorkbookExporter.Export(path, entries, snapshots);
        using var archive = ZipFile.OpenRead(path);
        using (var reader = new StreamReader(
                   (archive.GetEntry("xl/workbook.xml")
                    ?? throw new InvalidOperationException("Workbook metadata is missing.")).Open(), Encoding.UTF8))
        {
            Contains("日期待確認", reader.ReadToEnd(), "The pending-date worksheet is missing.");
        }

        using var stream = (archive.GetEntry("xl/worksheets/sheet2.xml")
                            ?? throw new InvalidOperationException("Pending-date worksheet is missing.")).Open();
        var sheet = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Equal(3, sheet.Descendants(ns + "row").Count(),
            "Only two unresolved channel/order aggregates should be listed.");
        var xml = sheet.ToString();
        Contains("UNRESOLVED-1", xml, "An unresolved order is missing from the guard worksheet.");
        Contains("備註", xml, "Pending-date worksheet is missing the note column.");
        Contains("待確認日期備註", xml, "Pending-date worksheet did not preserve the order note.");
        Contains("CONFLICT-ORDER", xml, "A conflicting order should remain visible for review.");
        DoesNotContain("DATED-ORDER", xml, "A dated order must not enter the guard worksheet.");
        DoesNotContain("FILLED-BY-SNAPSHOT", xml, "A snapshot-resolved order must leave the guard worksheet.");
        DoesNotContain("ZERO-ORDER", xml, "A zero-quantity unresolved aggregate should be omitted.");
        var numericValues = sheet.Descendants(ns + "c")
            .Where(cell => ((string?)cell.Attribute("r"))?.StartsWith("D", StringComparison.OrdinalIgnoreCase) == true &&
                           (string?)cell.Attribute("r") != "D1")
            .Select(cell => ((string?)cell.Attribute("t"), cell.Element(ns + "v")?.Value))
            .ToArray();
        True(numericValues.All(value => value.Item1 is null),
            "Pending-date quantities must be stored as Excel numbers.");
        True(numericValues.Select(value => value.Item2).ToHashSet().SetEquals(["1", "5"]),
            "Pending-date quantities were not aggregated correctly.");
        using var giftSheetStream = (archive.GetEntry("xl/worksheets/sheet1.xml")
                                     ?? throw new InvalidOperationException("Three-piece worksheet is missing.")).Open();
        var giftSheetXml = XDocument.Load(giftSheetStream).ToString();
        DoesNotContain("UNRESOLVED-1", giftSheetXml,
            "An unresolved order must not also enter a dated gift-box sheet.");
        DoesNotContain("CONFLICT-ORDER", giftSheetXml,
            "A stale ledger date overridden by a conflicting snapshot must leave the dated sheet.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task PendingDateWorkbookWritesBackSpecifiedDate()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-pending-writeback-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 12, 9, 0, 0, TimeSpan.FromHours(8));
    var entry = new OrderChangeEntry(
        "pending-entry", at, "site1", "網站1", "PENDING-WRITEBACK", "parent:pending", "蝦皮賣場",
        "2026/08/04 08:00:00", null, "1902248", "六入", 0m, 2m, 2m,
        "新增", "F", "待處理", true, string.Empty,
        OriginalDeliveryDate: null, GiftBoxSize: 6, Note: "網站無法修改日期");
    try
    {
        GiftBoxWorkbookExporter.Export(path, [entry]);
        var baseline = GiftBoxWorkbookExporter.ReadEditableRows(path).Single();
        Equal("pending", baseline.Sheet, "The unresolved order must start in the pending worksheet.");
        Equal(string.Empty, baseline.OriginalDeliveryDate,
            "The pending specified-delivery cell must initially be blank.");

        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var sheetEntry = archive.GetEntry("xl/worksheets/sheet2.xml")!;
            XDocument sheet;
            using (var input = sheetEntry.Open()) sheet = XDocument.Load(input);
            var headers = sheet.Descendants(ns + "row").First()
                .Descendants(ns + "t").Select(text => text.Value).ToArray();
            True(headers.SequenceEqual([
                    "通路別", "品名", "指定到貨日", "數量", "確認", "地點", "訂單編號", "備註", "訂單鍵", "來源明細指紋"]),
                "The pending worksheet must expose only specified delivery date and never an order-date column.");
            DoesNotContain("下單日", string.Join("|", headers),
                "The pending worksheet must not invite users to enter an order date.");
            var row = sheet.Descendants(ns + "row").Single(item => (string?)item.Attribute("r") == "2");
            var dateCell = row.Elements(ns + "c")
                .Single(item => ((string?)item.Attribute("r"))?.StartsWith("C", StringComparison.Ordinal) == true);
            dateCell.RemoveNodes();
            dateCell.Attribute("t")?.Remove();
            dateCell.SetAttributeValue("s", "2");
            dateCell.Add(new XElement(ns + "v", "46254"));
            sheetEntry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet2.xml");
            using var output = replacement.Open();
            sheet.Save(output);
        }

        var edited = GiftBoxWorkbookExporter.ReadEditableRows(path).Single();
        Equal("2026-08-20", edited.OriginalDeliveryDate,
            "The pending specified-delivery date must be read back from Excel.");
        Equal(string.Empty, edited.OrderDate,
            "The pending worksheet must not expose or read an order date.");
        var overrides = ManualOverrideWorkflow.CaptureChanges([baseline], [edited], [], at.AddMinutes(1));
        var applied = ManualOverrideWorkflow.Apply([entry], overrides).Single();
        Equal("2026-08-20", applied.OriginalDeliveryDate,
            "The Excel-entered specified delivery date must become a permanent override.");
        Equal("2026-08-13", applied.DeliveryDate,
            "The order date must be derived automatically as seven days before delivery.");

        GiftBoxWorkbookExporter.Export(path, [applied]);
        using (var updatedArchive = ZipFile.OpenRead(path))
        {
            using var detailInput = updatedArchive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            var detailXml = XDocument.Load(detailInput).ToString();
            using var pendingInput = updatedArchive.GetEntry("xl/worksheets/sheet2.xml")!.Open();
            var pendingXml = XDocument.Load(pendingInput).ToString();
            Contains("PENDING-WRITEBACK", detailXml,
                "The resolved order must move into order details on re-export.");
            DoesNotContain("PENDING-WRITEBACK", pendingXml,
                "The resolved order must leave pending review on re-export.");
        }

        var invalidCurrent = edited with { OriginalDeliveryDate = "2026-08-19" };
        try
        {
            ManualOverrideWorkflow.CaptureChanges(
                [baseline], [invalidCurrent], [], at.AddMinutes(2));
            throw new InvalidOperationException(
                "A pending date that is not Tuesday or Thursday must be rejected instead of adjusted silently.");
        }
        catch (InvalidDataException)
        {
            // Expected: pending dates are business inputs, not silently adjusted guesses.
        }

        GiftBoxWorkbookExporter.Export(path, [entry]);
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var sheetEntry = archive.GetEntry("xl/worksheets/sheet2.xml")!;
            XDocument sheet;
            using (var input = sheetEntry.Open()) sheet = XDocument.Load(input);
            var row = sheet.Descendants(ns + "row").Single(item => (string?)item.Attribute("r") == "2");
            var dateCell = row.Elements(ns + "c")
                .Single(item => ((string?)item.Attribute("r"))?.StartsWith("C", StringComparison.Ordinal) == true);
            dateCell.RemoveNodes();
            dateCell.SetAttributeValue("t", "inlineStr");
            dateCell.Add(new XElement(ns + "is", new XElement(ns + "t", "2026-08-20")));
            sheetEntry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet2.xml");
            using var output = replacement.Open();
            sheet.Save(output);
        }
        try
        {
            GiftBoxWorkbookExporter.ReadEditableRows(path);
            throw new InvalidOperationException("A text-formatted pending date must be rejected.");
        }
        catch (InvalidDataException)
        {
            // Expected: only an Excel numeric date is accepted.
        }
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task CurrentSnapshotProjectsCurrentOrdersWithoutLedger()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-main-item-projection-{Guid.NewGuid():N}.xlsx");
    var synchronizedAt = new DateTimeOffset(2026, 8, 7, 9, 0, 0, TimeSpan.FromHours(8));
    var snapshots = new[]
    {
        new OrderLineSnapshot(
            "site1", "ORDER-1", "parent:A", "parent", null, "shopee", "賣場",
            "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "PARENT", null,
            "蛋黃酥組合禮盒", null, 2m, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "ORDER-1", "item:parent:A:sku:1", "item", "parent:A", "shopee", "賣場",
            "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "4710964232435", "1902226",
            "芝麻蛋黃酥3入", null, 6m, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site2", "ORDER-2", "parent:B", "parent", null, "momo", "賣場2",
            "2026/08/05 08:00:00", "2026/08/20", "2026-08-13", "4710964232411", "1902248",
            "芝麻蛋黃酥6入", null, 4m, 0m, "F", synchronizedAt,
            ChannelName: "MOMO", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "ORDER-MULTI", "parent:C", "parent", null, "shopee", "賣場",
            "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "BUNDLE", null,
            "多主商品組合", null, 1m, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理",
            Note: "共享備註"),
        new OrderLineSnapshot(
            "site1", "ORDER-MULTI", "item:parent:C:single:1", "item", "parent:C", "shopee", "賣場",
            "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "BARCODE-SINGLE", "1902072",
            "來源單顆A", null, 5m, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "ORDER-MULTI", "item:parent:C:single:2", "item", "parent:C", "shopee", "賣場",
            "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "BARCODE-SINGLE", "1902072",
            "來源單顆B", null, 7m, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "ORDER-MULTI", "item:parent:C:six", "item", "parent:C", "shopee", "賣場",
            "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "BARCODE-SIX", "1902248",
            "來源六入", null, 2m, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "SS2607220002", "parent:three", "parent", null, "shopee", "美安",
            "2026/07/22 08:00:00", null, "2026-08-13", "SELLER-THREE", null,
            "芝初8倍細黑芝麻蛋黃酥禮盒(小)", "3入蛋黃酥", 1m, 0m, "F", synchronizedAt,
            ChannelName: "美安", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "SS2607220002", "item:parent:three", "item", "parent:three", "shopee", "美安",
            "2026/07/22 08:00:00", null, "2026-08-13", "4710964232435", "1902226",
            "芝麻蛋黃酥3入", null, 1m, 0m, "F", synchronizedAt,
            ChannelName: "美安", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "SS2607220002", "parent:six", "parent", null, "shopee", "美安",
            "2026/07/22 08:00:00", null, "2026-08-13", "SELLER-SIX", null,
            "芝初8倍細黑芝麻蛋黃酥禮盒", "蛋黃酥6入", 1m, 0m, "F", synchronizedAt,
            ChannelName: "美安", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "SS2607220002", "item:parent:six", "item", "parent:six", "shopee", "美安",
            "2026/07/22 08:00:00", null, "2026-08-13", "4710964232701", "1902248",
            "芝初8倍細黑芝麻Q潤蛋黃酥禮盒-2026版", null, 1m, 0m, "F", synchronizedAt,
            ChannelName: "美安", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site2", "ORDER-THREE-PACK", "parent:three-pack", "parent", null, "pos", "新光A4",
            "2026/07/24 14:39:01", "2026/08/31", "2026-08-20", "1902072*3", null,
            "芝麻蛋黃酥3顆組", null, 40m, 0m, "F", synchronizedAt,
            ChannelName: "正-新光A4", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site2", "ORDER-THREE-PACK", "item:parent:three-pack", "item", "parent:three-pack", "pos", "新光A4",
            "2026/07/24 14:39:01", "2026/08/31", "2026-08-20", "4710964232107", "1902072",
            "芝麻蛋黃酥(1入)組", null, 120m, 0m, "F", synchronizedAt,
            ChannelName: "正-新光A4", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "ORDER-NON-TARGET", "parent:non-target", "parent", null, "shopee", "賣場",
            "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "UNKNOWN", null,
            "待確認蛋黃酥", null, 9m, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理"),
        new OrderLineSnapshot(
            "site1", "ORDER-NON-TARGET", "item:parent:non-target", "item", "parent:non-target", "shopee", "賣場",
            "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "UNKNOWN", "9999999",
            "待確認商品", null, 9m, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理")
    };

    var projected = CurrentSnapshotOrderProjector.Project(
        snapshots,
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["site1"] = "網站1",
            ["site2"] = "網站2"
        },
        synchronizedAt.AddMinutes(1));

    Equal(7, projected.Count, "Each distinct target item number must enter the formal order projection.");
    var first = projected.Single(entry => entry.ExternalOrderNo == "ORDER-1");
    Equal(6m, first.QuantityChange, "Current projection must use the target child quantity.");
    Equal("1902226", first.Sku, "Current projection must classify by the target child item number.");
    Equal("parent:A", first.ExternalLineKey, "The stable source detail key must be preserved.");
    Equal(3, first.GiftBoxSize, "A unique target child SKU must classify an ambiguous parent.");
    Equal("2026-08-11", first.DeliveryDate, "The current resolved order date must be preserved.");
    True(projected.All(entry => entry.ChangeType == "目前訂單" && entry.PreviousQuantity == 0m),
        "Current projection must not describe historical differences.");
    var leaf = projected.Single(entry => entry.ExternalOrderNo == "ORDER-2");
    Equal("1902248", leaf.Sku, "A leaf parent must use its own target item number.");
    Equal(4m, leaf.QuantityChange, "A leaf parent without children must keep the parent quantity.");

    var multi = projected.Where(entry => entry.ExternalOrderNo == "ORDER-MULTI").ToArray();
    Equal(2, multi.Length, "Different target child item numbers under one parent must become separate rows.");
    Equal(12m, multi.Single(entry => entry.Sku == "1902072").QuantityChange,
        "Repeated child rows with the same item number must sum child quantities.");
    Equal(2m, multi.Single(entry => entry.Sku == "1902248").QuantityChange,
        "Each distinct child item number must keep its own child quantity.");
    True(multi.Select(entry => entry.ExternalLineKey).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 2,
        "Split item-number rows must have distinct stable source-line keys.");
    True(multi.All(entry => entry.DeliveryDate == "2026-08-11" && entry.Note == "共享備註"),
        "Split item-number rows must preserve their shared parent date and note.");
    var reorderedSnapshots = snapshots
        .OrderByDescending(snapshot => snapshot.ExternalLineKey, StringComparer.OrdinalIgnoreCase)
        .ToArray();
    var reorderedMulti = CurrentSnapshotOrderProjector.Project(
            reorderedSnapshots,
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["site1"] = "網站1",
                ["site2"] = "網站2"
            },
            synchronizedAt.AddMinutes(1))
        .Where(entry => entry.ExternalOrderNo == "ORDER-MULTI")
        .ToDictionary(entry => entry.Sku!, entry => entry.ExternalLineKey, StringComparer.Ordinal);
    True(multi.ToDictionary(entry => entry.Sku!, entry => entry.ExternalLineKey, StringComparer.Ordinal)
            .SequenceEqual(reorderedMulti),
        "Child-array reordering must not change the item-number to source-line-key mapping.");
    var threePack = projected.Single(entry => entry.ExternalOrderNo == "ORDER-THREE-PACK");
    Equal("1902072", threePack.Sku, "A three-single bundle must classify as the single main item.");
    Equal(120m, threePack.QuantityChange,
        "A parent quantity of 40 with child quantity 120 must export 120 single items.");
    True(projected.All(entry => entry.ExternalOrderNo != "ORDER-NON-TARGET"),
        "A parent and child without a target item number must not enter the formal projection.");

    var stableParent = new OrderLineSnapshot(
        "site1", "ORDER-STABLE", "parent:stable", "parent", null, "shopee", "賣場",
        "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "4710964232411", null,
        "六入主商品", null, 1m, 0m, "F", synchronizedAt,
        ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理");
    var stableSix = new OrderLineSnapshot(
        "site1", "ORDER-STABLE", "item:stable:six", "item", "parent:stable", "shopee", "賣場",
        "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "4710964232701", "1902248",
        "六入主商品", null, 1m, 0m, "F", synchronizedAt,
        ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理");
    var stableSingle = new OrderLineSnapshot(
        "site1", "ORDER-STABLE", "item:stable:single", "item", "parent:stable", "shopee", "賣場",
        "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "4710964232107", "1902072",
        "新增單顆", null, 2m, 0m, "F", synchronizedAt,
        ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理");
    var sourceNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["site1"] = "網站1" };
    var singleton = CurrentSnapshotOrderProjector.Project([stableParent, stableSix], sourceNames, synchronizedAt)
        .Single();
    var expanded = CurrentSnapshotOrderProjector.Project(
        [stableParent, stableSix, stableSingle], sourceNames, synchronizedAt);
    True(expanded.All(entry => entry.ExternalLineKey != singleton.ExternalLineKey),
        "An ambiguous multi-item expansion must retire the bare parent key instead of guessing its product.");
    var legacyFingerprint = ManualOverrideWorkflow.BuildEntryFingerprint(singleton);
    var legacyOverride = new OrderRowManualOverride(
        ManualOverrideWorkflow.BuildRowKey([legacyFingerprint]),
        [legacyFingerprint],
        ManualFieldOverride.None,
        ManualFieldOverride.None,
        ManualFieldOverride.None,
        new ManualFieldOverride(true, "已確認"),
        synchronizedAt);
    var expandedWithOverride = ManualOverrideWorkflow.Apply(expanded, [legacyOverride]);
    True(expandedWithOverride.All(entry => string.IsNullOrEmpty(entry.Confirmation)),
        "A legacy override must not be guessed onto either product after an ambiguous split.");

    var conflictingParent = stableParent with
    {
        ExternalOrderNo = "ORDER-CONFLICT",
        ExternalLineKey = "parent:conflict",
        ItemNo = "1902248"
    };
    var conflictingSingle = stableSingle with
    {
        ExternalOrderNo = "ORDER-CONFLICT",
        ExternalLineKey = "item:conflict:single",
        ParentLineKey = "parent:conflict"
    };
    var conflictingSix = stableSix with
    {
        ExternalOrderNo = "ORDER-CONFLICT",
        ExternalLineKey = "item:conflict:six",
        ParentLineKey = "parent:conflict"
    };
    var conflictingSingleton = CurrentSnapshotOrderProjector.Project(
            [conflictingParent, conflictingSingle], sourceNames, synchronizedAt)
        .Single();
    var conflictingExpanded = CurrentSnapshotOrderProjector.Project(
        [conflictingParent, conflictingSingle, conflictingSix], sourceNames, synchronizedAt);
    True(conflictingExpanded.All(entry => entry.ExternalLineKey != conflictingSingleton.ExternalLineKey),
        "A parent item number must not steal a legacy child key when a second target item appears.");
    var conflictingFingerprint = ManualOverrideWorkflow.BuildEntryFingerprint(conflictingSingleton);
    var conflictingOverride = legacyOverride with
    {
        RowKey = ManualOverrideWorkflow.BuildRowKey([conflictingFingerprint]),
        SourceLineFingerprints = [conflictingFingerprint]
    };
    True(ManualOverrideWorkflow.Apply(conflictingExpanded, [conflictingOverride])
            .All(entry => string.IsNullOrEmpty(entry.Confirmation)),
        "A conflicting parent item number must never redirect an old override to the new product.");

    try
    {
        GiftBoxWorkbookExporter.Export(path, projected, snapshots);
        using var archive = ZipFile.OpenRead(path);
        using var input = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var detailDocument = XDocument.Load(input);
        var xml = detailDocument.ToString();
        Equal(2, xml.Split("ORDER-MULTI", StringSplitOptions.None).Length - 1,
            "Both target item-number rows must enter the formal Excel workbook.");
        Contains("單顆(有盒)(霧面袋)", xml,
            "The single-item child item number must map to the fixed Excel product name.");
        Contains("6入鐵盒2026年鐵盒(黑金袋)", xml,
            "The six-piece child item number must map to the fixed Excel product name.");
        Equal(2, xml.Split("SS2607220002", StringSplitOptions.None).Length - 1,
            "The real ReYi order must export both its three-piece and six-piece parent products.");
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var threePackRow = detailDocument.Descendants(spreadsheet + "row")
            .Single(row => row.Descendants(spreadsheet + "t")
                .Any(text => text.Value == "ORDER-THREE-PACK"));
        Equal("120", threePackRow.Elements(spreadsheet + "c")
                .Single(cell => ((string?)cell.Attribute("r"))?.StartsWith("E", StringComparison.Ordinal) == true)
                .Element(spreadsheet + "v")?.Value,
            "The formal Excel quantity must be the numeric child quantity 120.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task WmsDependentSinglesUseEarliestGiftDate()
{
    var synchronizedAt = new DateTimeOffset(2026, 8, 12, 9, 0, 0, TimeSpan.FromHours(8));
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-dependent-single-{Guid.NewGuid():N}.xlsx");

    OrderLineSnapshot Parent(
        string orderNo,
        string lineKey,
        string? arrivalDate,
        string? derivedDate,
        decimal quantity = 1m) =>
        new(
            "site1", orderNo, lineKey, "parent", null, "shopee", "測試賣場",
            "2026/08/01 08:00:00", arrivalDate, derivedDate, "SELLER-SKU", null,
            "測試賣場商品", null, quantity, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理");

    OrderLineSnapshot Child(
        string orderNo,
        string lineKey,
        string parentLineKey,
        string itemNo,
        decimal quantity = 1m) =>
        new(
            "site1", orderNo, lineKey, "item", parentLineKey, "shopee", "測試賣場",
            "2026/08/01 08:00:00", null, null, $"BARCODE-{itemNo}", itemNo,
            $"主商品-{itemNo}", null, quantity, 0m, "F", synchronizedAt,
            ChannelName: "蝦皮賣場", ShippingDateStatus: "resolved", StatusName: "待處理");

    var snapshots = new[]
    {
        Parent("2608017BB8ENDM", "parent:late-gift", "2026/09/10", "2026-09-03", 3m),
        Child("2608017BB8ENDM", "item:late-gift", "parent:late-gift", "1902248", 3m),
        Parent("2608017BB8ENDM", "parent:early-gift", "2026/09/03", "2026-08-27", 2m),
        Child("2608017BB8ENDM", "item:early-gift", "parent:early-gift", "1902226", 2m),
        Parent("2608017BB8ENDM", "parent:single", null, null, 2m),
        Child("2608017BB8ENDM", "item:single", "parent:single", "1902072", 2m),

        Parent("SOLO-UNDATED", "parent:only-single", null, null, 4m),
        Child("SOLO-UNDATED", "item:only-single", "parent:only-single", "1902072", 4m),

        Parent("SOLO-DATED", "parent:dated-only-single", "2026/09/17", "2026-09-10", 5m),
        Child("SOLO-DATED", "item:dated-only-single", "parent:dated-only-single", "1902072", 5m),

        Parent("GIFT-DATE-UNKNOWN", "parent:unknown-gift", null, null),
        Child("GIFT-DATE-UNKNOWN", "item:unknown-gift", "parent:unknown-gift", "1902247"),
        Parent("GIFT-DATE-UNKNOWN", "parent:unknown-single", null, null),
        Child("GIFT-DATE-UNKNOWN", "item:unknown-single", "parent:unknown-single", "1902072"),

        Parent("DATED-SINGLE", "parent:dated-gift", "2026/09/03", "2026-08-27"),
        Child("DATED-SINGLE", "item:dated-gift", "parent:dated-gift", "1902248"),
        Parent("DATED-SINGLE", "parent:dated-single", "2026/09/17", "2026-09-10"),
        Child("DATED-SINGLE", "item:dated-single", "parent:dated-single", "1902072"),

        Parent("CROSS-SOURCE", "parent:site1-gift", "2026/09/03", "2026-08-27"),
        Child("CROSS-SOURCE", "item:site1-gift", "parent:site1-gift", "1902248"),
        Parent("CROSS-SOURCE", "parent:site1-single", null, null),
        Child("CROSS-SOURCE", "item:site1-single", "parent:site1-single", "1902072"),
        Parent("CROSS-SOURCE", "parent:site2-gift", "2026/08/27", "2026-08-20") with
        {
            SourceCode = "site2"
        },
        Child("CROSS-SOURCE", "item:site2-gift", "parent:site2-gift", "1902247") with
        {
            SourceCode = "site2"
        }
    };
    var projected = CurrentSnapshotOrderProjector.Project(
        snapshots,
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["site1"] = "網站1",
            ["site2"] = "網站2"
        },
        synchronizedAt);

    var inheritedSingle = projected.Single(entry =>
        entry.ExternalOrderNo == "2608017BB8ENDM" && entry.Sku == "1902072");
    Equal("2026-09-03", inheritedSingle.OriginalDeliveryDate,
        "An undated single item must inherit the earliest gift delivery date in the same order.");
    Equal("2026-08-27", inheritedSingle.DeliveryDate,
        "An inherited single item must use the matching gift order date.");
    True(projected.All(entry => entry.ExternalOrderNo != "SOLO-UNDATED"),
        "An undated WMS order containing only dependent single items must not enter the formal projection.");
    var datedOnlySingle = projected.Single(entry => entry.ExternalOrderNo == "SOLO-DATED");
    Equal("2026-09-17", datedOnlySingle.OriginalDeliveryDate,
        "A dated single item must remain exportable even when the order has no gift box.");
    Equal("2026-09-10", datedOnlySingle.DeliveryDate,
        "A dated single item without a gift box must keep its own order date.");
    var unresolvedSingle = projected.Single(entry =>
        entry.ExternalOrderNo == "GIFT-DATE-UNKNOWN" && entry.Sku == "1902072");
    True(string.IsNullOrWhiteSpace(unresolvedSingle.OriginalDeliveryDate) &&
         string.IsNullOrWhiteSpace(unresolvedSingle.DeliveryDate),
        "A single item with an undated gift must remain pending review instead of guessing a date.");
    var independentlyDatedSingle = projected.Single(entry =>
        entry.ExternalOrderNo == "DATED-SINGLE" && entry.Sku == "1902072");
    Equal("2026-09-17", independentlyDatedSingle.OriginalDeliveryDate,
        "A single item with its own valid date must keep that date.");
    Equal("2026-09-10", independentlyDatedSingle.DeliveryDate,
        "A single item with its own valid order date must not be overwritten by the gift date.");
    var sourceScopedSingle = projected.Single(entry =>
        entry.SourceCode == "site1" && entry.ExternalOrderNo == "CROSS-SOURCE" && entry.Sku == "1902072");
    Equal("2026-09-03", sourceScopedSingle.OriginalDeliveryDate,
        "A single item must inherit only from a gift in the same website source.");
    Equal("2026-08-27", sourceScopedSingle.DeliveryDate,
        "An earlier gift date from another website must never cross the source boundary.");

    try
    {
        GiftBoxWorkbookExporter.Export(path, projected, snapshots);
        using var archive = ZipFile.OpenRead(path);
        using var detailInput = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var detailXml = XDocument.Load(detailInput).ToString();
        using var pendingInput = archive.GetEntry("xl/worksheets/sheet2.xml")!.Open();
        var pendingXml = XDocument.Load(pendingInput).ToString();
        Contains("2608017BB8ENDM", detailXml,
            "The inherited single item must enter the dated order-detail worksheet.");
        DoesNotContain("SOLO-UNDATED", detailXml + pendingXml,
            "An order containing only dependent single items must not enter either formal worksheet.");
        Contains("GIFT-DATE-UNKNOWN", pendingXml,
            "A single item with an existing but undated gift must remain visible for date review.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static async Task FullSnapshotReplacementIsAtomicAcrossSources()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-atomic-snapshot-{Guid.NewGuid():N}");
    var path = Path.Combine(directory, "order-snapshot.json");
    Directory.CreateDirectory(directory);
    var synchronizedAt = new DateTimeOffset(2026, 8, 7, 9, 0, 0, TimeSpan.FromHours(8));

    OrderLineSnapshot Row(string source, string orderNo, string lineKey, decimal quantity) => new(
        source, orderNo, lineKey, "parent", null, "channel", "shop",
        "2026/08/04 08:00:00", "2026/08/18", "2026-08-11", "4710964232411", null,
        "gift box", null, quantity, 0m, "F", synchronizedAt,
        ChannelName: "channel", ShippingDateStatus: "resolved", StatusName: "pending");

    try
    {
        var store = new OrderSnapshotStore(path);
        await store.ReplaceAllAsync([
            Row("site1", "OLD-1", "parent:old-1", 1m),
            Row("site2", "OLD-2", "parent:old-2", 2m)
        ]);

        var invalid = new[]
        {
            Row("site1", "NEW-1", "parent:new-1", 3m),
            Row("site2", "DUPLICATE", "parent:same", 4m),
            Row("site2", "DUPLICATE", "parent:same", 5m)
        };
        var rejected = false;
        try
        {
            await store.ReplaceAllAsync(invalid);
        }
        catch (InvalidDataException)
        {
            rejected = true;
        }

        True(rejected, "An invalid complete snapshot must be rejected before replacing stored data.");
        var afterFailure = await store.GetAllAsync();
        True(afterFailure.Select(line => line.ExternalOrderNo).ToHashSet().SetEquals(["OLD-1", "OLD-2"]),
            "A failed complete reload must preserve both previous source snapshots.");

        await store.ReplaceAllAsync([
            Row("site1", "NEW-1", "parent:new-1", 3m),
            Row("site2", "NEW-2", "parent:new-2", 4m)
        ]);
        var replaced = await store.GetAllAsync();
        True(replaced.Select(line => line.ExternalOrderNo).ToHashSet().SetEquals(["NEW-1", "NEW-2"]),
            "A valid complete reload must replace both source snapshots together.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
    }
}

static Task SameSkuDifferentDateLinesKeepIdentityWhenReordered()
{
    var synchronizedAt = DateTimeOffset.UtcNow;
    IReadOnlyDictionary<string, string?> Row(string products) => new Dictionary<string, string?>
    {
        ["order_no"] = "ORDER-DATES",
        ["status_code"] = "F",
        ["products"] = products
    };
    const string firstProduct =
        "{\"sku\":\"SAME-SKU\",\"name\":\"蛋黃酥禮盒\",\"spec\":\"8/18~8/21出貨\",\"qty\":1}";
    const string secondProduct =
        "{\"sku\":\"SAME-SKU\",\"name\":\"蛋黃酥禮盒\",\"spec\":\"8/25~8/28出貨\",\"qty\":2}";
    var original = OrderSnapshotStore.BuildSourceSnapshot(
        "site1", [Row($"[{firstProduct},{secondProduct}]")], synchronizedAt);
    var reordered = OrderSnapshotStore.BuildSourceSnapshot(
        "site1", [Row($"[{secondProduct},{firstProduct}]")], synchronizedAt.AddMinutes(1));

    var originalBySpec = original.ToDictionary(line => line.Spec!, line => line.ExternalLineKey);
    var reorderedBySpec = reordered.ToDictionary(line => line.Spec!, line => line.ExternalLineKey);
    Equal(originalBySpec["8/18~8/21出貨"], reorderedBySpec["8/18~8/21出貨"],
        "The first shipping-date detail changed identity after API reordering.");
    Equal(originalBySpec["8/25~8/28出貨"], reorderedBySpec["8/25~8/28出貨"],
        "The second shipping-date detail changed identity after API reordering.");
    True(originalBySpec.Values.Distinct(StringComparer.OrdinalIgnoreCase).Count() == 2,
        "Different shipping-date details must not share one identity.");
    return Task.CompletedTask;
}

static Task ManualOverridesPreserveBlanksAndClearByDetailRow()
{
    var entry = new OrderChangeEntry(
        "entry", DateTimeOffset.UtcNow, "site1", "site1", "ORDER-1", "parent:1", "MOMO",
        "2026/08/01", "2026-08-06", "4710964232411", "six", 0m, 2m, 2m,
        "current", "F", "pending", false,
        OriginalDeliveryDate: "2026-08-15", GiftBoxSize: 6, SourceLocation: "彰廠");
    var fingerprint = ManualOverrideWorkflow.BuildEntryFingerprint(entry);
    var rowKey = ManualOverrideWorkflow.BuildRowKey([fingerprint]);
    var baseline = new WorkbookEditableRow(
        rowKey, "detail", [fingerprint], "2026-08-15", "2026-08-06", "彰廠", "",
        "MOMO", "ORDER-1", "六入");
    var edited = baseline with
    {
        OriginalDeliveryDate = "",
        OrderDate = "2026-08-04",
        Location = "",
        Confirmation = "已確認"
    };

    var overrides = ManualOverrideWorkflow.CaptureChanges(
        [baseline], [edited], [], new DateTimeOffset(2026, 8, 7, 10, 0, 0, TimeSpan.FromHours(8)));
    Equal(1, overrides.Count, "One edited detail row should produce one persistent override.");
    var manual = overrides.Single();
    True(manual.OriginalDeliveryDate.IsOverridden && manual.OriginalDeliveryDate.Value == "",
        "An intentionally blank specified-delivery date must remain an explicit override.");
    True(manual.Location.IsOverridden && manual.Location.Value == "",
        "An intentionally blank location must remain an explicit override.");

    var applied = ManualOverrideWorkflow.Apply([entry], overrides).Single();
    Equal(null, applied.OriginalDeliveryDate, "A blank specified-delivery override must not be regenerated.");
    Equal("2026-08-04", applied.DeliveryDate, "The manually edited order date must win.");
    Equal("", applied.SourceLocation, "A blank location override must win over automatic mapping.");
    True(applied.HasManualLocation, "The exporter must be able to distinguish a manual blank location.");
    Equal("已確認", applied.Confirmation, "The manually edited confirmation must win.");
    True(applied.HasManualConfirmation, "The exporter must preserve an explicit confirmation override.");

    var cleared = ManualOverrideWorkflow.Clear(overrides, rowKey);
    Equal(0, cleared.Count, "Clearing a selected detail row must remove all four manual fields together.");
    var restored = ManualOverrideWorkflow.Apply([entry], cleared).Single();
    Equal("2026-08-15", restored.OriginalDeliveryDate,
        "After clearing, the automatic specified-delivery date must be restored.");
    Equal("彰廠", restored.SourceLocation, "After clearing, the automatic location must be restored.");

    var baselineWithAggregateId = baseline with { AggregateId = "AGGREGATE-1" };
    var legacyEdited = edited with
    {
        RowKey = ManualOverrideWorkflow.BuildRowKey(["LEGACY-FINGERPRINT"]),
        SourceLineFingerprints = ["LEGACY-FINGERPRINT"],
        AggregateId = "AGGREGATE-1"
    };
    var migrated = ManualOverrideWorkflow.CaptureChanges(
        [baselineWithAggregateId], [legacyEdited], [], DateTimeOffset.UtcNow).Single();
    True(migrated.SourceLineFingerprints.SequenceEqual([fingerprint]),
        "A legacy workbook row matched by aggregate ID must migrate to the current stable fingerprint.");

    var stableManualBaseline = baseline with
    {
        ExternalOrderNo = "手打檔-NEWIDENTITY-7",
        AggregateId = "NEW-MANUAL-AGGREGATE"
    };
    var legacyManualEdited = edited with
    {
        RowKey = ManualOverrideWorkflow.BuildRowKey(["LEGACY-MANUAL-FINGERPRINT"]),
        SourceLineFingerprints = ["LEGACY-MANUAL-FINGERPRINT"],
        ExternalOrderNo = "手打檔-OLDFILEHASH-7",
        AggregateId = "OLD-MANUAL-AGGREGATE"
    };
    var migratedManual = ManualOverrideWorkflow.CaptureChanges(
        [stableManualBaseline], [legacyManualEdited], [], DateTimeOffset.UtcNow).Single();
    Equal(stableManualBaseline.RowKey, migratedManual.RowKey,
        "A legacy manual row must migrate by source row when its old whole-file hash changes.");

    return Task.CompletedTask;
}

static async Task ManualOverrideStatePersistsWorkbookPathAndRows()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-manual-state-{Guid.NewGuid():N}");
    var statePath = Path.Combine(directory, "manual-overrides.json");
    var workbookPath = Path.Combine(directory, "formal.xlsx");
    Directory.CreateDirectory(directory);
    try
    {
        var row = new WorkbookEditableRow(
            ManualOverrideWorkflow.BuildRowKey(["ABC123"]),
            "detail", ["ABC123"], "2026-08-15", "2026-08-06", "彰廠", "已確認",
            "MOMO", "ORDER-1", "六入");
        var manual = new OrderRowManualOverride(
            row.RowKey,
            row.SourceLineFingerprints,
            new ManualFieldOverride(true, ""),
            ManualFieldOverride.None,
            new ManualFieldOverride(true, ""),
            new ManualFieldOverride(true, "已確認"),
            DateTimeOffset.UtcNow);
        var store = new ManualOverrideStore(statePath);
        await store.SaveAsync(new ManualOverrideState(
            workbookPath, [row], [manual], [new WorkbookRowAutomaticReset(row.RowKey, row.SourceLineFingerprints)]));

        var loaded = await new ManualOverrideStore(statePath).LoadAsync();
        Equal(Path.GetFullPath(workbookPath), loaded.LastWorkbookPath,
            "The last formal workbook path must survive application restarts.");
        Equal(1, loaded.LastExportedRows.Count, "The last exported row baseline must be persisted.");
        Equal(1, loaded.Overrides.Count, "Manual overrides must be persisted independently from WMS history.");
        Equal(1, loaded.PendingAutomaticResets?.Count ?? 0,
            "A pending clear operation must survive until the next successful workbook export.");
        True(loaded.Overrides[0].OriginalDeliveryDate.IsOverridden &&
             loaded.Overrides[0].OriginalDeliveryDate.Value == "",
            "A persistent blank override must round-trip through JSON.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
    }
}

static Task FormalWorkbookReadsBackEditableRows()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-editable-rows-{Guid.NewGuid():N}.xlsx");
    try
    {
        var entry = new OrderChangeEntry(
            "entry", new DateTimeOffset(2026, 8, 7, 9, 0, 0, TimeSpan.FromHours(8)),
            "site1", "site1", "ORDER-READ", "parent:read", "MOMO",
            "2026/08/01", "2026-08-06", "4710964232411", "six", 0m, 2m, 2m,
            "current", "F", "pending", false,
            Confirmation: "已確認", OriginalDeliveryDate: "2026-08-15", GiftBoxSize: 6,
            SourceLocation: "彰廠", HasManualLocation: true, HasManualConfirmation: true);
        GiftBoxWorkbookExporter.Export(path, [entry]);

        var rows = GiftBoxWorkbookExporter.ReadEditableRows(path);
        Equal(1, rows.Count, "The detail worksheet should expose one editable row.");
        var row = rows.Single();
        Equal("detail", row.Sheet, "The editable row must retain its worksheet kind.");
        Equal("2026-08-15", row.OriginalDeliveryDate, "Excel date cells must be read as ISO dates.");
        Equal("2026-08-06", row.OrderDate, "The order-date cell must be read as an ISO date.");
        Equal("彰廠", row.Location, "The location must be read back.");
        Equal("已確認", row.Confirmation, "The confirmation must be read back.");
        Equal("ORDER-READ", row.ExternalOrderNo, "The display order number must be available for clearing UI.");
        Equal(ManualOverrideWorkflow.BuildRowKey(row.SourceLineFingerprints), row.RowKey,
            "The persistent row key must be derived from stable source-line fingerprints.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task ClearedOverrideRestoresAutomaticFieldsInExistingWorkbook()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-clear-existing-{Guid.NewGuid():N}.xlsx");
    var entry = new OrderChangeEntry(
        "entry", DateTimeOffset.UtcNow, "site1", "site1", "ORDER-CLEAR", "parent:clear", "蝦皮賣場",
        "2026/08/01", "2026-08-11", "4710964232411", "six", 0m, 1m, 1m,
        "current", "F", "pending", false,
        OriginalDeliveryDate: "2026-08-18", GiftBoxSize: 6);
    try
    {
        GiftBoxWorkbookExporter.Export(path, [entry]);
        var baseline = GiftBoxWorkbookExporter.ReadEditableRows(path);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument sheet;
            using (var input = sheetEntry.Open()) sheet = XDocument.Load(input);
            var row = sheet.Descendants(ns + "row").Single(item => (string?)item.Attribute("r") == "2");
            SetText("C", "2026-08-20");
            SetText("D", "2026-08-13");
            SetText("F", "人工確認");
            SetText("G", "人工地點");
            sheetEntry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var output = replacement.Open();
            sheet.Save(output);

            void SetText(string column, string value)
            {
                var cell = row.Elements(ns + "c")
                    .Single(item => ((string?)item.Attribute("r"))?.StartsWith(column,
                        StringComparison.Ordinal) == true);
                cell.RemoveNodes();
                cell.SetAttributeValue("t", "inlineStr");
                cell.Add(new XElement(ns + "is", new XElement(ns + "t", value)));
            }
        }

        var editedWorkbook = GiftBoxWorkbookExporter.ReadEditableRows(path);
        var overrides = ManualOverrideWorkflow.CaptureChanges(
            baseline, editedWorkbook, [], DateTimeOffset.UtcNow);
        var manuallyApplied = ManualOverrideWorkflow.Apply([entry], overrides);
        GiftBoxWorkbookExporter.Export(path, manuallyApplied);
        var preserved = GiftBoxWorkbookExporter.ReadEditableRows(path).Single();
        Equal("2026-08-20", preserved.OriginalDeliveryDate,
            "A workbook-edited specified-delivery date must survive re-export.");
        Equal("2026-08-13", preserved.OrderDate,
            "A workbook-edited order date must survive re-export.");
        Equal("人工地點", preserved.Location, "A workbook-edited location must survive re-export.");
        Equal("人工確認", preserved.Confirmation, "A workbook-edited confirmation must survive re-export.");

        var reset = new WorkbookRowAutomaticReset(
            overrides.Single().RowKey,
            overrides.Single().SourceLineFingerprints);
        var automatic = ManualOverrideWorkflow.ApplyAutomaticResets([entry], [reset]);
        GiftBoxWorkbookExporter.Export(path, automatic);
        var restored = GiftBoxWorkbookExporter.ReadEditableRows(path).Single();
        Equal("2026-08-18", restored.OriginalDeliveryDate,
            "Clearing an override must restore the automatic specified-delivery date.");
        Equal("2026-08-11", restored.OrderDate,
            "Clearing an override must restore the automatic order date.");
        Equal("創鵬", restored.Location,
            "Clearing an override must ignore the old workbook value and restore channel mapping.");
        Equal("", restored.Confirmation,
            "Clearing an override must ignore and remove the old workbook confirmation.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task WorkbookUsesThreeConsolidatedWorksheetsAndNativeFormulas()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-three-sheets-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 7, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(
        string id, string source, string channel, string orderNo, string sku,
        int? size, decimal quantity, string? originalDate, string? orderDate) => new(
        EntryId: id,
        ObservedAt: at,
        SourceCode: source,
        SourceName: source,
        ExternalOrderNo: orderNo,
        ExternalLineKey: $"parent:{id}",
        ChannelName: channel,
        OrderDate: "2026/08/06 08:00:00",
        DeliveryDate: orderDate,
        Sku: sku,
        ProductName: "來源商品名稱",
        PreviousQuantity: 0m,
        NewQuantity: quantity,
        QuantityChange: quantity,
        ChangeType: "新增",
        StatusCode: "F",
        StatusName: "待處理",
        NeedsReview: originalDate is null,
        OriginalDeliveryDate: originalDate,
        GiftBoxSize: size,
        Note: originalDate is null ? "日期格式待確認" : "");
    var entries = new[]
    {
        Entry("ONE", "manual_excel", "MOMO", "ORDER-1", "1902072", null, 2m, "2026-08-15", "2026-08-06"),
        Entry("THREE", "site1", "蝦皮賣場", "ORDER-3", "4710964232435", 3, 3m, "2026-08-15", "2026-08-06"),
        Entry("SIX", "site1", "美安", "ORDER-6", "4710964232411", 6, 4m, "2026-08-22", "2026-08-13"),
        Entry("NINE", "site2", "好的文創", "ORDER-9", "4710964232565", 9, 5m, "2026-08-22", "2026-08-13"),
        Entry("UNKNOWN", "manual_excel", "未設定通路", "ORDER-U", "1902226", 3, 7m, "2026-08-15", "2026-08-06"),
        Entry("PENDING", "site1", "美安", "ORDER-P", "4710964232565", 9, 6m, null, null)
    };

    try
    {
        GiftBoxWorkbookExporter.Export(path, entries);
        using var archive = ZipFile.OpenRead(path);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using (var workbookInput = archive.GetEntry("xl/workbook.xml")!.Open())
        {
            var workbook = XDocument.Load(workbookInput);
            var names = workbook.Descendants(ns + "sheet")
                .Select(sheet => (string?)sheet.Attribute("name"))
                .ToArray();
            True(names.SequenceEqual(["訂單明細", "日期待確認", "統計"]),
                "The workbook must contain exactly the three agreed worksheets in order.");
            var calculation = workbook.Descendants(ns + "calcPr").Single();
            Equal("1", (string?)calculation.Attribute("fullCalcOnLoad"),
                "The workbook must request a full formula recalculation when opened.");
            Equal("1", (string?)calculation.Attribute("forceFullCalc"),
                "The workbook must force recalculation of the statistics formulas.");
            Equal("auto", (string?)calculation.Attribute("calcMode"),
                "The workbook must explicitly enable automatic calculation for Excel 2013.");
        }

        using (var detailInput = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open())
        {
            var detail = XDocument.Load(detailInput);
            var headers = detail.Descendants(ns + "row").First().Elements(ns + "c")
                .Take(9).Select(cell => string.Concat(cell.Descendants(ns + "t").Select(text => text.Value)))
                .ToArray();
            True(headers.SequenceEqual([
                    "通路別", "品名", "指定到貨日", "下單日", "數量", "確認", "地點", "訂單編號", "備註"]),
                "Order-detail headers do not match the agreed order.");
            var xml = detail.ToString();
            Contains("單顆(有盒)(霧面袋)", xml, "The mapped one-piece product name is missing.");
            Contains("3入禮盒紙盒(黑金袋)", xml, "The three-piece product name is missing.");
            Contains("6入鐵盒2026年鐵盒(黑金袋)", xml, "The six-piece product name is missing.");
            Contains("9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)", xml, "The nine-piece product name is missing.");
            DoesNotContain("ORDER-P", xml, "An unresolved order must not enter order details.");
        }

        using (var pendingInput = archive.GetEntry("xl/worksheets/sheet2.xml")!.Open())
        {
            var pending = XDocument.Load(pendingInput);
            var headers = pending.Descendants(ns + "row").First().Elements(ns + "c")
                .Take(8).Select(cell => string.Concat(cell.Descendants(ns + "t").Select(text => text.Value)))
                .ToArray();
            True(headers.SequenceEqual(["通路別", "品名", "指定到貨日", "數量", "確認", "地點", "訂單編號", "備註"]),
                "Pending-review headers do not match the agreed order.");
            Contains("ORDER-P", pending.ToString(), "The unresolved order is missing from pending review.");
        }

        using var statisticsInput = archive.GetEntry("xl/worksheets/sheet3.xml")!.Open();
        var statistics = XDocument.Load(statisticsInput);
        foreach (var row in statistics.Descendants(ns + "row"))
        {
            var cellReferences = row.Elements(ns + "c")
                .Select(cell => (string?)cell.Attribute("r") ?? string.Empty)
                .ToArray();
            True(cellReferences.SequenceEqual(cellReferences.OrderBy(reference => reference, StringComparer.Ordinal)),
                $"Statistics row {row.Attribute("r")?.Value} cells must be written in ascending reference order.");
        }
        var statisticHeaders = statistics.Descendants(ns + "row").First().Elements(ns + "c")
            .Select(cell => string.Concat(cell.Descendants(ns + "t").Select(text => text.Value)))
            .ToArray();
        True(statisticHeaders.SequenceEqual([
                "下單日", "到貨日", "地點",
                "單顆(無盒)(透明袋)", "單顆(無盒)(霧面袋)", "單顆(有盒)(霧面袋)",
                "3入條裝(無盒)(霧面袋)", "3入禮盒紙盒(黑金袋)",
                "6入鐵盒2024年鐵盒(黑金袋)", "6入鐵盒2026年鐵盒(黑金袋)",
                "9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)"]),
            "Statistics headers do not match the agreed order.");
        var formulas = statistics.Descendants(ns + "f").Select(formula => formula.Value).ToArray();
        Equal(65, formulas.Length,
            "Five formula-capacity rows must each contain three displayed keys, eight SUMIFS, and two helpers.");
        var sumifs = formulas.Where(formula => formula.Contains("SUMIFS(", StringComparison.OrdinalIgnoreCase)).ToArray();
        Equal(40, sumifs.Length, "Each formula-capacity row must use eight native SUMIFS formulas.");
        True(sumifs.All(formula => formula.Contains("SUMIFS(", StringComparison.OrdinalIgnoreCase)),
            "Statistics quantities must use native SUMIFS formulas.");
        True(sumifs.All(formula => formula.Contains("'訂單明細'!$E$2:$E$", StringComparison.Ordinal)),
            "Statistics formulas must sum the numeric quantity column in order details.");
        True(sumifs.All(formula => formula.Contains("'訂單明細'!$D$2:$D$", StringComparison.Ordinal)),
            "Statistics formulas must match the order-date column in order details.");
        True(sumifs.All(formula => formula.Contains("'訂單明細'!$C$2:$C$", StringComparison.Ordinal)),
            "Statistics formulas must match the arrival-date column in order details.");
        True(sumifs.All(formula => formula.Contains("'訂單明細'!$G$2:$G$", StringComparison.Ordinal)),
            "Statistics formulas must match the location column in order details.");
        True(sumifs.All(formula => formula.Contains("'訂單明細'!$B$2:$B$", StringComparison.Ordinal)),
            "Statistics formulas must match the product-name column in order details.");
        Contains("未設定", statistics.ToString(), "A blank detail location must have an explicit statistics row.");
        True(sumifs.All(formula => formula.Contains("IF($C", StringComparison.Ordinal)),
            "The 未設定 statistics row must query blank detail locations rather than the display label.");
        True(formulas.Any(formula => formula.Contains("_xlfn.AGGREGATE(15,6", StringComparison.OrdinalIgnoreCase)),
            "Post-Excel-2007 AGGREGATE formulas must use the OOXML compatibility prefix.");
        True(formulas.Where(formula => formula.Contains("_xlfn.AGGREGATE(", StringComparison.OrdinalIgnoreCase))
                .All(formula => formula.Contains("ROW()-1", StringComparison.OrdinalIgnoreCase) &&
                                !formula.Contains("ROWS($M$2:", StringComparison.OrdinalIgnoreCase)),
            "The compact-index helper must not reference its own M cell range.");
        True(formulas.Any(formula => formula.Contains("COUNTIFS(", StringComparison.OrdinalIgnoreCase)),
            "Excel formulas must identify unique date and location combinations.");
        True(formulas.All(formula => !formula.Contains("TODAY()", StringComparison.OrdinalIgnoreCase)),
            "Excel formulas must not exclude past order dates.");
        True(formulas.Where(formula => formula.Contains("COUNTIFS(", StringComparison.OrdinalIgnoreCase))
                .All(formula => !formula.Contains(":'訂單明細'!", StringComparison.Ordinal)),
            "Excel 2013 range endpoints must not repeat the worksheet qualifier after the colon.");
        True(statistics.Descendants(ns + "c").Where(cell => cell.Element(ns + "f") is not null)
            .All(cell => cell.Element(ns + "v") is { Value.Length: 0 }),
            "Excel 2013 formula cells must contain an empty value marker without a cached result.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task WorkbookStatisticsSupports421DetailRows()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-statistics-421-{Guid.NewGuid():N}.xlsx");
    var observedAt = new DateTimeOffset(2026, 8, 7, 9, 0, 0, TimeSpan.FromHours(8));
    var entries = Enumerable.Range(0, 421)
        .Select(index =>
        {
            var orderDate = new DateOnly(2026, 8, 11).AddDays(index);
            return new OrderChangeEntry(
                $"STRESS-{index}", observedAt, "site1", "網站1", $"ORDER-{index:000}",
                $"parent:{index}", "蝦皮賣場", "2026/08/01 08:00:00",
                orderDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                "4710964232435", "蛋黃酥三入禮盒", 0m, 1m, 1m,
                "新增", "F", "待處理", false,
                OriginalDeliveryDate: orderDate.AddDays(7).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        })
        .ToArray();

    try
    {
        GiftBoxWorkbookExporter.Export(path, entries);
        using var archive = ZipFile.OpenRead(path);
        using var input = archive.GetEntry("xl/worksheets/sheet3.xml")!.Open();
        var statistics = XDocument.Load(input);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var rows = statistics.Descendants(ns + "row").ToArray();
        Equal(422, rows.Length, "A 421-detail workbook must provide one formula-capacity row per detail row.");
        Equal(421 * 13, statistics.Descendants(ns + "f").Count(),
            "Each large statistics row must retain eleven visible formulas and two helper formulas.");
        Equal(421 * 13, statistics.Descendants(ns + "c")
                .Count(cell => cell.Element(ns + "f") is not null && cell.Element(ns + "v") is { Value.Length: 0 }),
            "Every large-workbook formula cell must have an empty Excel 2013 value marker.");
        foreach (var row in rows.Skip(1))
        {
            var rowNumber = (string?)row.Attribute("r") ?? string.Empty;
            var expected = Enumerable.Range(1, 13).Select(column => $"{ColumnNameForTest(column)}{rowNumber}");
            var actual = row.Elements(ns + "c").Select(cell => (string?)cell.Attribute("r") ?? string.Empty);
            True(actual.SequenceEqual(expected), $"Large statistics row {rowNumber} has invalid cell ordering.");
        }
        DoesNotContain(":'訂單明細'!", statistics.ToString(),
            "Large statistics formulas contain an invalid repeated worksheet qualifier.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static string ColumnNameForTest(int column)
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

static Task WorkbookOutputPreservesAllOrderDates()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-all-dates-output-{Guid.NewGuid():N}.xlsx");
    var observedAt = new DateTimeOffset(2026, 8, 7, 9, 0, 0, TimeSpan.FromHours(8));
    var testToday = new DateOnly(2026, 8, 7);
    OrderChangeEntry Entry(string id, string orderNo, string? orderDate) =>
        new(id, observedAt, "site1", "網站1", orderNo, $"parent:{orderNo}", "蝦皮賣場",
            "2026/08/01 08:00:00", orderDate, "4710964232435", "蛋黃酥三入禮盒",
            0m, 1m, 1m, "新增", "F", "待處理", false, OriginalDeliveryDate: "2026-08-18");

    var entries = new[]
    {
        Entry("PAST", "ORDER-PAST", testToday.AddDays(-1).ToString("yyyy-MM-dd")),
        Entry("TODAY", "ORDER-TODAY", testToday.ToString("yyyy-MM-dd")),
        Entry("FUTURE", "ORDER-FUTURE", testToday.AddDays(4).ToString("yyyy-MM-dd")),
        Entry("PENDING", "ORDER-PENDING", null)
    };

    try
    {
        GiftBoxWorkbookExporter.Export(path, entries);

        using var archive = ZipFile.OpenRead(path);
        using var detailReader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml")!.Open(), Encoding.UTF8);
        var detail = detailReader.ReadToEnd();
        Contains("ORDER-PAST", detail, "A past order date must remain in order details.");
        Contains("ORDER-TODAY", detail, "An order dated today must remain in order details.");
        Contains("ORDER-FUTURE", detail, "A future order must remain in order details.");

        using var pendingReader = new StreamReader(archive.GetEntry("xl/worksheets/sheet2.xml")!.Open(), Encoding.UTF8);
        var pending = pendingReader.ReadToEnd();
        Contains("ORDER-PENDING", pending, "An unresolved date must remain in pending review.");

        Equal(4, entries.Length, "Exporting all dates must not mutate source entries.");
        Equal("2026-08-06", entries[0].DeliveryDate,
            "Exporting all dates must not mutate the historical order date.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task WorkbookLocationUsesChannelMappingWithoutOverwritingManualValue()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-location-map-{Guid.NewGuid():N}.xlsx");
    var nonWebsitePath = Path.Combine(Path.GetTempPath(), $"mid-autumn-non-website-location-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 6, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(
        string id,
        string channel,
        string orderNo,
        string sourceCode = "site1",
        string? sourceLocation = null) => new(
        EntryId: id,
        ObservedAt: at,
        SourceCode: sourceCode,
        SourceName: sourceCode == "pos_excel" ? "POS機" : "網站1－睿驛",
        ExternalOrderNo: orderNo,
        ExternalLineKey: $"parent:{id}",
        ChannelName: channel,
        OrderDate: "2026/08/04 08:00:00",
        DeliveryDate: "2026-08-06",
        Sku: sourceCode == "pos_excel" ? "1902247" : "4710964232435",
        ProductName: sourceCode == "pos_excel" ? "芝初黑芝麻Q潤蛋黃酥禮盒(9入)" : "芝初黑芝麻蛋黃酥3入禮盒",
        PreviousQuantity: 0m,
        NewQuantity: 1m,
        QuantityChange: 1m,
        ChangeType: "新增",
        StatusCode: "F",
        StatusName: "待處理",
        NeedsReview: false,
        OriginalDeliveryDate: "2026-08-15",
        GiftBoxSize: 3,
        SourceLocation: sourceLocation);
    var entries = new[]
    {
        Entry("A", "蝦皮賣場", "ORDER-A"),
        Entry("B", " 正-新光A4 ", "ORDER-B"),
        Entry("C", "電子商務-二聯客戶", "ORDER-C"),
        Entry("D", "正-夢時代", "ORDER-D"),
        Entry("E", "MOMO", "ORDER-E"),
        Entry("F", "正-京站", "ORDER-F"),
        Entry("G", "好的文創", "ORDER-G"),
        Entry("H", "美安", "ORDER-H"),
        Entry("I", "未設定通路", "ORDER-I"),
        Entry("J", "正-台中高鐵", "ORDER-J"),
        Entry("K", "臨新竹大全", "ORDER-K"),
        Entry("L", "POS 正-京站", "ORDER-L", "pos_excel"),
        Entry("N", "臨新竹大全聯", "ORDER-N", "site2"),
        Entry("Q", "臨新莊宏匯", "ORDER-Q"),
        Entry("R", "臨桃園大江", "ORDER-R"),
        Entry("S", "臨大葉高島屋", "ORDER-S"),
        Entry("T", "臨新光三越台南新天地", "ORDER-T"),
        Entry("V", "YAHOO", "ORDER-V"),
        Entry("W", "阿瘦", "ORDER-W"),
        Entry("M", "官網", "ORDER-M", "manual_excel", "來源自訂地點"),
        Entry("P", "美安", "ORDER-P") with { DeliveryDate = null, OriginalDeliveryDate = null }
    };

    try
    {
        GiftBoxWorkbookExporter.Export(path, entries);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        static string CellText(XElement row, XNamespace ns, string column) =>
            string.Concat(row.Elements(ns + "c")
                .Single(cell => ((string?)cell.Attribute("r"))?.StartsWith(column, StringComparison.Ordinal) == true)
                .Descendants(ns + "t")
                .Select(text => text.Value));

        using (var archive = ZipFile.OpenRead(path))
        using (var input = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open())
        {
            var rows = XDocument.Load(input).Descendants(ns + "row").Skip(1).ToArray();
            var locations = rows.ToDictionary(row => CellText(row, ns, "A"), row => CellText(row, ns, "G"));
            var expectedLocations = new Dictionary<string, string>
            {
                ["蝦皮賣場"] = "創鵬",
                ["電子商務-二聯客戶"] = "彰廠",
                ["正-新光A4"] = "彰廠",
                ["正-夢時代"] = "彰廠",
                ["MOMO"] = "創鵬",
                ["正-京站"] = "彰廠",
                ["好的文創"] = "彰廠",
                ["美安"] = "創鵬",
                ["正-台中高鐵"] = "彰廠",
                ["臨新竹大全聯"] = "彰廠",
                ["臨新莊宏匯"] = "彰廠",
                ["臨桃園大江"] = "彰廠",
                ["臨大葉高島屋"] = "彰廠",
                ["臨新光三越台南新天地"] = "彰廠",
                ["YAHOO"] = "彰廠",
                ["阿瘦"] = "彰廠",
                ["POS 正-京站"] = "躉泰",
                ["官網"] = "來源自訂地點"
            };
            foreach (var expected in expectedLocations)
            {
                Equal(expected.Value, locations[expected.Key], $"Location mapping failed for {expected.Key}.");
            }
            Equal(string.Empty, locations["未設定通路"], "An unmapped channel must keep a blank location.");
            Equal(string.Empty, locations["臨新竹大全"],
                "The nonexistent channel 臨新竹大全 must not remain in the website mapping.");
        }
        using (var archive = ZipFile.OpenRead(path))
        using (var input = archive.GetEntry("xl/worksheets/sheet2.xml")!.Open())
        {
            var pendingRow = XDocument.Load(input).Descendants(ns + "row").Skip(1).Single();
            Equal("創鵬", CellText(pendingRow, ns, "F"),
                "Pending-review rows must use the same location mapping.");
        }

        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument sheet;
            using (var input = sheetEntry.Open()) sheet = XDocument.Load(input);
            var targetRow = sheet.Descendants(ns + "row").Skip(1)
                .Single(row => CellText(row, ns, "A") == "蝦皮賣場");
            var locationCell = targetRow.Elements(ns + "c")
                .Single(cell => ((string?)cell.Attribute("r"))?.StartsWith("G", StringComparison.Ordinal) == true);
            locationCell.RemoveNodes();
            locationCell.SetAttributeValue("t", "inlineStr");
            locationCell.Add(new XElement(ns + "is", new XElement(ns + "t", "人工地點")));
            var manualSourceRow = sheet.Descendants(ns + "row").Skip(1)
                .Single(row => CellText(row, ns, "A") == "官網");
            var manualSourceLocationCell = manualSourceRow.Elements(ns + "c")
                .Single(cell => ((string?)cell.Attribute("r"))?.StartsWith("G", StringComparison.Ordinal) == true);
            manualSourceLocationCell.RemoveNodes();
            manualSourceLocationCell.SetAttributeValue("t", "inlineStr");
            manualSourceLocationCell.Add(new XElement(ns + "is", new XElement(ns + "t", "人工覆寫地點")));
            sheetEntry.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var output = replacement.Open();
            sheet.Save(output);
        }

        GiftBoxWorkbookExporter.Export(path, entries);
        using var updated = ZipFile.OpenRead(path);
        using var updatedInput = updated.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var updatedRows = XDocument.Load(updatedInput).Descendants(ns + "row").Skip(1).ToArray();
        var updatedShopee = updatedRows.Single(row => CellText(row, ns, "A") == "蝦皮賣場");
        Equal("人工地點", CellText(updatedShopee, ns, "G"),
            "An existing manual location must not be overwritten by the channel mapping.");
        var updatedManualSource = updatedRows.Single(row => CellText(row, ns, "A") == "官網");
        Equal("人工覆寫地點", CellText(updatedManualSource, ns, "G"),
            "An existing workbook location must not be overwritten by a custom spreadsheet source location.");

        GiftBoxWorkbookExporter.Export(
            nonWebsitePath,
            [
                Entry("U", "正-新光A4", "ORDER-U", "erp"),
                Entry("X", "ERP自訂", "ORDER-X", "erp", "人工ERP地點")
            ]);
        using var nonWebsite = ZipFile.OpenRead(nonWebsitePath);
        using var nonWebsiteInput = nonWebsite.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var nonWebsiteRows = XDocument.Load(nonWebsiteInput).Descendants(ns + "row").Skip(1).ToArray();
        Equal("彰廠", CellText(nonWebsiteRows.Single(row => CellText(row, ns, "A") == "正-新光A4"), ns, "G"),
            "ERP rows without a source location must default to 彰廠.");
        Equal("人工ERP地點", CellText(nonWebsiteRows.Single(row => CellText(row, ns, "A") == "ERP自訂"), ns, "G"),
            "An explicit ERP location must remain higher priority than the default.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
        if (File.Exists(nonWebsitePath)) File.Delete(nonWebsitePath);
    }

    return Task.CompletedTask;
}

static Task WorkbookShowsOriginalAndOrderDatesAndPreservesLocation()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-original-order-date-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 6, 9, 0, 0, TimeSpan.FromHours(8));
    var entry = new OrderChangeEntry(
        "DATE-COLUMNS", at, "site1", "網站1－睿驛", "ORDER-1", "parent:A", "蝦皮賣場",
        "2026/08/04 08:00:00", "2026-08-06", "4710964232435", "蛋黃酥3入禮盒",
        0m, 2m, 2m, "新增", "F", "待處理", false, string.Empty,
        OriginalDeliveryDate: "2026-08-15");

    try
    {
        GiftBoxWorkbookExporter.Export(path, [entry]);
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var sheet = archive.GetEntry("xl/worksheets/sheet1.xml")
                ?? throw new InvalidOperationException("Three-piece worksheet is missing.");
            XDocument document;
            using (var input = sheet.Open()) document = XDocument.Load(input);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var xml = document.ToString();
            Contains("指定到貨日", xml, "The original arrival date column is missing.");
            Contains("下單日", xml, "The adjusted order date column is missing.");
            Contains("地點", xml, "The manual location column is missing.");
            Contains(ExcelSerial("2026-08-15"), xml, "The original arrival date is missing.");
            Contains(ExcelSerial("2026-08-06"), xml, "The adjusted order date is missing.");
            var locationCell = document.Descendants(ns + "c")
                .Single(cell => (string?)cell.Attribute("r") == "G2");
            locationCell.Attribute("t")?.Remove();
            locationCell.RemoveNodes();
            locationCell.Add(new XElement(ns + "v", "123"));
            sheet.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var output = replacement.Open();
            document.Save(output);
        }

        GiftBoxWorkbookExporter.Export(path, [entry]);
        using var resultArchive = ZipFile.OpenRead(path);
        using var resultInput = (resultArchive.GetEntry("xl/worksheets/sheet1.xml")
                                 ?? throw new InvalidOperationException("Three-piece worksheet is missing.")).Open();
        Contains("123", XDocument.Load(resultInput).ToString(),
            "Re-exporting the aggregate must preserve a numeric manually entered location.");
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }

    return Task.CompletedTask;
}

static Task WorkbookRejectsManualFieldFormulaWithoutOverwrite()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-manual-formula-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 6, 9, 0, 0, TimeSpan.FromHours(8));
    var entry = new OrderChangeEntry(
        "FORMULA", at, "site1", "網站1", "ORDER", "parent:A", "蝦皮賣場", null,
        "2026-08-06", "4710964232435", "蛋黃酥3入禮盒", 0m, 1m, 1m,
        "新增", "F", "待處理", false, OriginalDeliveryDate: "2026-08-15");
    try
    {
        GiftBoxWorkbookExporter.Export(path, [entry]);
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            var sheet = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            XDocument document;
            using (var input = sheet.Open()) document = XDocument.Load(input);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var location = document.Descendants(ns + "c").Single(cell => (string?)cell.Attribute("r") == "G2");
            location.Attribute("t")?.Remove();
            location.RemoveNodes();
            location.Add(new XElement(ns + "f", "1+1"), new XElement(ns + "v", "2"));
            sheet.Delete();
            var replacement = archive.CreateEntry("xl/worksheets/sheet1.xml");
            using var output = replacement.Open();
            document.Save(output);
        }

        try
        {
            GiftBoxWorkbookExporter.Export(path, [entry]);
            throw new InvalidOperationException("Expected a formula in a manual field to reject the update.");
        }
        catch (InvalidDataException)
        {
            using var archive = ZipFile.OpenRead(path);
            using var input = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            Contains("<f>1+1</f>", XDocument.Load(input).ToString(),
                "Rejecting the update must leave the original formula intact.");
        }
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }
    return Task.CompletedTask;
}

static Task WorkbookUsesDfkaiAndHasNoAutoFilters()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-font-filter-{Guid.NewGuid():N}.xlsx");
    var entry = new OrderChangeEntry(
        "FONT", DateTimeOffset.Now, "site1", "網站1", "ORDER", "parent:A", "蝦皮賣場", null,
        "2026-08-06", "4710964232435", "蛋黃酥3入禮盒", 0m, 1m, 1m,
        "新增", "F", "待處理", false, OriginalDeliveryDate: "2026-08-15");
    try
    {
        GiftBoxWorkbookExporter.Export(path, [entry]);
        using var archive = ZipFile.OpenRead(path);
        string Read(string entryPath)
        {
            using var reader = new StreamReader(
                (archive.GetEntry(entryPath) ?? throw new InvalidOperationException($"Missing {entryPath}.")).Open(),
                Encoding.UTF8);
            return reader.ReadToEnd();
        }

        var styles = Read("xl/styles.xml");
        Contains("標楷體", styles, "The workbook font is not DFKai.");
        var stylesDocument = XDocument.Parse(styles);
        XNamespace styleNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var fonts = stylesDocument.Descendants(styleNs + "fonts").Elements(styleNs + "font").ToArray();
        Equal("標楷體", (string?)fonts[1].Element(styleNs + "name")?.Attribute("val"),
            "Font index 1 is not DFKai.");
        var cellFormats = stylesDocument.Descendants(styleNs + "cellXfs").Elements(styleNs + "xf").ToArray();
        Equal("1", (string?)cellFormats[1].Attribute("fontId"),
            "Cell style index 1 does not point to the DFKai font.");
        Contains("styles", Read("[Content_Types].xml"), "The style content type is missing.");
        Contains("relationships/styles", Read("xl/_rels/workbook.xml.rels"),
            "The workbook style relationship is missing.");
        foreach (var sheetPath in Enumerable.Range(1, 3).Select(index => $"xl/worksheets/sheet{index}.xml"))
        {
            var sheet = Read(sheetPath);
            DoesNotContain("autoFilter", sheet, $"{sheetPath} must not contain an automatic filter.");
            var document = XDocument.Parse(sheet);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            True(document.Descendants(ns + "c").All(cell => (string?)cell.Attribute("s") is "1" or "2"),
                $"Every output cell in {sheetPath} must use the DFKai style.");
        }
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }
    return Task.CompletedTask;
}

static Task LegacyWorkbookUpdateAddsDfkaiAndRemovesFilters()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-legacy-style-{Guid.NewGuid():N}.xlsx");
    var entry = new OrderChangeEntry(
        "LEGACY", DateTimeOffset.Now, "site1", "網站1", "ORDER", "parent:A", "蝦皮賣場", null,
        "2026-08-06", "4710964232435", "蛋黃酥3入禮盒", 0m, 1m, 1m,
        "新增", "F", "待處理", false, OriginalDeliveryDate: "2026-08-15");
    try
    {
        GiftBoxWorkbookExporter.Export(path, [entry]);
        using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
        {
            archive.GetEntry("xl/styles.xml")?.Delete();
            void RewriteDocument(string entryPath, Action<XDocument> change)
            {
                var source = archive.GetEntry(entryPath)!;
                XDocument document;
                using (var input = source.Open()) document = XDocument.Load(input);
                change(document);
                source.Delete();
                var replacement = archive.CreateEntry(entryPath);
                using var output = replacement.Open();
                document.Save(output);
            }
            RewriteDocument("[Content_Types].xml", document =>
                document.Descendants().Where(element =>
                    (string?)element.Attribute("PartName") == "/xl/styles.xml").Remove());
            RewriteDocument("xl/_rels/workbook.xml.rels", document =>
                document.Descendants().Where(element =>
                    ((string?)element.Attribute("Type"))?.EndsWith("/styles", StringComparison.Ordinal) == true).Remove());
            RewriteDocument("xl/workbook.xml", document =>
            {
                var calculation = document.Descendants()
                    .Single(element => element.Name.LocalName == "calcPr");
                calculation.Attribute("fullCalcOnLoad")?.Remove();
                calculation.Attribute("forceFullCalc")?.Remove();
            });
            for (var index = 1; index <= 3; index++)
            {
                RewriteDocument($"xl/worksheets/sheet{index}.xml", document =>
                {
                    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                    foreach (var cell in document.Descendants(ns + "c")) cell.Attribute("s")?.Remove();
                    if (index == 3)
                    {
                        var headerRow = document.Descendants(ns + "row").First();
                        headerRow.RemoveNodes();
                        var legacyHeaders = new[] { "下單日", "地點", "一入", "三入", "六入", "九入", "合計" };
                        for (var headerIndex = 0; headerIndex < legacyHeaders.Length; headerIndex++)
                        {
                            headerRow.Add(new XElement(ns + "c",
                                new XAttribute("r", $"{(char)('A' + headerIndex)}1"),
                                new XAttribute("t", "inlineStr"),
                                new XElement(ns + "is", new XElement(ns + "t", legacyHeaders[headerIndex]))));
                        }
                    }
                    document.Root?.Add(new XElement(ns + "autoFilter", new XAttribute("ref", "A1:A1")));
                });
            }
        }

        GiftBoxWorkbookExporter.Export(path, [entry]);
        using var updated = ZipFile.OpenRead(path);
        True(updated.GetEntry("xl/styles.xml") is not null, "Updating a legacy workbook must restore styles.xml.");
        using var styleReader = new StreamReader(updated.GetEntry("xl/styles.xml")!.Open(), Encoding.UTF8);
        Contains("標楷體", styleReader.ReadToEnd(), "Updating a legacy workbook did not restore DFKai.");
        using (var workbookInput = updated.GetEntry("xl/workbook.xml")!.Open())
        {
            var workbook = XDocument.Load(workbookInput);
            var calculation = workbook.Descendants().Single(element => element.Name.LocalName == "calcPr");
            Equal("1", (string?)calculation.Attribute("fullCalcOnLoad"),
                "Updating the workbook must restore full recalculation on open.");
            Equal("1", (string?)calculation.Attribute("forceFullCalc"),
                "Updating the workbook must restore forced formula recalculation.");
            Equal("auto", (string?)calculation.Attribute("calcMode"),
                "Updating the workbook must restore automatic Excel 2013 calculation.");
        }
        for (var index = 1; index <= 3; index++)
        {
            using var input = updated.GetEntry($"xl/worksheets/sheet{index}.xml")!.Open();
            var document = XDocument.Load(input);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            True(!document.Descendants(ns + "autoFilter").Any(), "A legacy auto-filter was not removed.");
            True(document.Descendants(ns + "c").All(cell => (string?)cell.Attribute("s") is "1" or "2"),
                "A legacy workbook cell did not receive the DFKai style.");
            if (index == 3)
            {
                var headers = document.Descendants(ns + "row").First().Elements(ns + "c")
                    .Select(cell => string.Concat(cell.Descendants(ns + "t").Select(text => text.Value)))
                    .ToArray();
                True(headers.SequenceEqual([
                        "下單日", "到貨日", "地點",
                        "單顆(無盒)(透明袋)", "單顆(無盒)(霧面袋)", "單顆(有盒)(霧面袋)",
                        "3入條裝(無盒)(霧面袋)", "3入禮盒紙盒(黑金袋)",
                        "6入鐵盒2024年鐵盒(黑金袋)", "6入鐵盒2026年鐵盒(黑金袋)",
                        "9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)"]),
                    "A legacy statistics sheet must upgrade to the fixed-product format.");
            }
        }
    }
    finally
    {
        if (File.Exists(path)) File.Delete(path);
    }
    return Task.CompletedTask;
}

static Task SpreadsheetSourcesMapColumnsAndTargetSkus()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-import-{Guid.NewGuid():N}");
    Directory.CreateDirectory(directory);
    var erpPath = Path.Combine(directory, "erp.xlsx");
    var erpAllSkusPath = Path.Combine(directory, "erp-all-skus.xlsx");
    var manualPath = Path.Combine(directory, "manual.xlsx");
    try
    {
        CreateSimpleWorkbook(erpPath,
        [
            ["ERP 報表"],
            [],
            ["單據日期", "銷貨單號", "客戶簡稱", "收貨人", "品號", "品名", "銷貨數量", "備註"],
            ["2026/09/07", "2321-20260900007", "客戶A", "王小明", "1902248", "六入禮盒", 6m, "ERP備註-附發票"],
            ["2026/09/07", "2321-20260900007", "客戶A", "王小明", "1902094", "芝麻粉禮盒", 6m, "非目標備註"],
            ["2026/09/07", "2321-20260900007", "客戶A", "王小明", "1902226", "三入禮盒", 1m, "ERP備註-三入"],
            ["2026/09/08", "2321-20260900008", "客戶B", "李小華", "1902226", "三入禮盒", 2m, "ERP備註-客戶B"]
        ], prependBlankSheet: true);
        var erp = SpreadsheetOrderImporter.ReadErp(erpPath, new DateTimeOffset(2026, 8, 6, 9, 0, 0, TimeSpan.FromHours(8)));
        Equal(3, erp.Rows.Count, "ERP target rows were not filtered correctly.");
        Equal(1, erp.ExcludedRowCount, "ERP excluded-row count is incorrect.");
        Equal("2321-20260900007", erp.Rows[0].ExternalOrderNo, "ERP sales order number mapping is incorrect.");
        Equal("1902248", erp.Rows[0].Sku, "ERP item number mapping is incorrect.");
        Equal("2026-09-03", erp.Rows[0].OriginalDeliveryDate, "ERP document date mapping is incorrect.");
        Equal("2026-08-27", erp.Rows[0].OrderDate, "ERP lead-time calculation is incorrect.");
        Equal(6m, erp.Rows[0].Quantity, "ERP sales quantity must remain numeric.");
        Equal("ERP備註-附發票", erp.Rows[0].Note, "ERP note mapping is incorrect.");
        Equal("ERP備註-附發票", erp.ExportEntries[0].Note, "ERP note must enter the export projection.");
        DoesNotContain("王小明", JsonSerializer.Serialize(erp),
            "ERP recipient must not enter preview or export data.");
        True(erp.ExportEntries.Select(entry => $"{entry.ExternalOrderNo}|{entry.ExternalLineKey}").Distinct().Count() == 3,
            "ERP rows need stable distinct detail keys.");
        Equal(2, erp.Rows.Count(row => row.ExternalOrderNo == "2321-20260900007"),
            "Two target products on the same ERP order must remain separate.");
        var repeatedErp = SpreadsheetOrderImporter.ReadErp(
            erpPath, new DateTimeOffset(2026, 8, 6, 9, 5, 0, TimeSpan.FromHours(8)));
        True(erp.ExportEntries.Select(entry => entry.EntryId).SequenceEqual(
                repeatedErp.ExportEntries.Select(entry => entry.EntryId)),
            "Reloading the same ERP file must produce the same stable entry identities.");

        var erpSkus = new[]
        {
            "1902199", "192199", "1902072", "1902042",
            "1902226", "1902197", "1902248", "1902247"
        };
        var erpAllSkuRows = new List<object?[]>
        {
            new object?[] { "單據日期", "銷貨單號", "客戶簡稱", "品號", "品名", "銷貨數量" }
        };
        erpAllSkuRows.AddRange(erpSkus.Select((sku, index) => new object?[]
        {
            "2026/09/07", $"ERP-ALL-{index + 1}", "ERP完整料號測試", sku, $"品名-{sku}", index + 1
        }));
        erpAllSkuRows.Add(new object?[]
            { "2026/09/07", "ERP-EXCLUDED", "ERP完整料號測試", "1902094", "非目標商品", 99m });
        CreateSimpleWorkbook(erpAllSkusPath, erpAllSkuRows);
        var erpAllSkus = SpreadsheetOrderImporter.ReadErp(
            erpAllSkusPath, new DateTimeOffset(2026, 8, 6, 9, 10, 0, TimeSpan.FromHours(8)));
        Equal(8, erpAllSkus.ExportEntries.Count, "ERP must accept all eight agreed exact item numbers.");
        True(erpAllSkus.ExportEntries.Select(entry => entry.Sku).OrderBy(value => value)
                .SequenceEqual(erpSkus.OrderBy(value => value)),
            "ERP accepted item numbers do not exactly match the agreed eight-item list.");
        Equal(1, erpAllSkus.ExcludedRowCount, "ERP must exclude a non-target item number.");

        CreateSimpleWorkbook(manualPath,
        [
            ["通路", "抵達日", "料號", "數量", "地點", "備註", "下單(計算,5工作日且為二或四)", "出貨(計算)"],
            ["官網", "2026/09/07", "1902247", 3m, "自訂倉庫", "手打明細備註", "", ""],
            ["官網", "", "1902199", 4m, "創鵬", "手打待確認備註", "", ""],
            ["官網", "", "1902072", 5m, "彰廠", "", "", ""],
            ["官網", "2026/09/08", "192199", 6m, "彰廠", "", "", ""],
            ["官網", "2026/09/08", "1902042", 7m, "彰廠", "", "", ""],
            ["官網", "2026/09/08", "1902197", 8m, "彰廠", "", "", ""],
            ["官網", "", "1902248", "", "創鵬", "", "", ""]
        ]);
        var manual = SpreadsheetOrderImporter.ReadManual(
            manualPath, new DateTimeOffset(2026, 8, 6, 9, 1, 0, TimeSpan.FromHours(8)));
        Equal(7, manual.Rows.Count, "Manual spreadsheet preview must retain all eight-category target rows, including blank quantity.");
        Equal(6, manual.ExportEntries.Count, "A blank manual quantity must not enter the export projection.");
        True(manual.ExportEntries.Any(entry => entry.Sku == "1902199"), "Unboxed single item is missing.");
        True(manual.ExportEntries.Any(entry => entry.Sku == "1902072"), "Boxed single item is missing.");
        True(manual.ExportEntries.Any(entry => entry.Sku == "192199"), "Matte-bag unboxed single item is missing.");
        True(manual.ExportEntries.Any(entry => entry.Sku == "1902042"), "Three-piece strip item is missing.");
        True(manual.ExportEntries.Any(entry => entry.Sku == "1902197"), "2024 six-piece tin item is missing.");
        Equal("自訂倉庫", manual.Rows.Single(row => row.Sku == "1902247").SourceLocation,
            "Manual spreadsheet preview must preserve a custom source location.");
        Equal("彰廠", manual.ExportEntries.Single(entry => entry.Sku == "1902072").SourceLocation,
            "Manual spreadsheet export entries must preserve their exact source location.");
        Equal(null, manual.Rows.Single(row => row.Sku == "1902199").OrderDate,
            "A blank manual arrival date must remain blank.");
        Equal("手打明細備註", manual.ExportEntries.Single(entry => entry.Sku == "1902247").Note,
            "Manual note must enter the dated export projection.");
        Equal("手打待確認備註", manual.ExportEntries.Single(entry => entry.Sku == "1902199").Note,
            "Manual note must enter the pending-date export projection.");

        var originalManualEntry = manual.ExportEntries.Single(entry => entry.Sku == "1902247");
        File.Delete(manualPath);
        CreateSimpleWorkbook(manualPath,
        [
            ["通路", "抵達日", "料號", "數量", "地點", "備註"],
            ["官網", "2026/09/07", "1902247", 3m, "自訂倉庫", "手打備註已修改"]
        ]);
        var manualAfterNoteEdit = SpreadsheetOrderImporter.ReadManual(
            manualPath, new DateTimeOffset(2026, 8, 6, 9, 2, 0, TimeSpan.FromHours(8)));
        var editedManualEntry = manualAfterNoteEdit.ExportEntries.Single();
        Equal(originalManualEntry.ExternalOrderNo, editedManualEntry.ExternalOrderNo,
            "Editing only a manual note must not change its synthetic order number.");
        Equal(originalManualEntry.ExternalLineKey, editedManualEntry.ExternalLineKey,
            "Editing only a manual note must not change its detail key.");
        Equal(originalManualEntry.EntryId, editedManualEntry.EntryId,
            "Editing only a manual note must not change its stable export identity.");
        Equal("手打備註已修改", editedManualEntry.Note,
            "The edited manual note must still enter the export projection.");
        DoesNotContain("Recipient", JsonSerializer.Serialize(manual),
            "Manual spreadsheet data must not define recipient fields.");

        var outputPath = Path.Combine(directory, "combined.xlsx");
        GiftBoxWorkbookExporter.Export(outputPath, erp.ExportEntries.Concat(manual.ExportEntries).ToArray());
        using var output = ZipFile.OpenRead(outputPath);
        string Sheet(string name)
        {
            using var reader = new StreamReader(
                (output.GetEntry(name) ?? throw new InvalidOperationException($"Missing worksheet {name}.")).Open(),
                Encoding.UTF8);
            return reader.ReadToEnd();
        }
        Contains("客戶B", Sheet("xl/worksheets/sheet1.xml"), "ERP three-piece row did not enter the workbook.");
        Contains("客戶A", Sheet("xl/worksheets/sheet1.xml"), "ERP six-piece row did not enter the workbook.");
        Contains("官網", Sheet("xl/worksheets/sheet1.xml"), "Manual nine-piece row did not enter the workbook.");
        Contains("自訂倉庫", Sheet("xl/worksheets/sheet1.xml"),
            "A custom manual spreadsheet location did not enter order details.");
        Contains("ERP備註-附發票", Sheet("xl/worksheets/sheet1.xml"),
            "ERP note did not enter order details.");
        Contains("手打明細備註", Sheet("xl/worksheets/sheet1.xml"),
            "Manual note did not enter order details.");
        Contains("單顆(無盒)(透明袋)", Sheet("xl/worksheets/sheet2.xml"),
            "Manual single rows without dates did not enter pending review.");
        Contains("創鵬", Sheet("xl/worksheets/sheet2.xml"),
            "A custom manual spreadsheet location did not enter pending review.");
        Contains("彰廠", Sheet("xl/worksheets/sheet2.xml"),
            "Each pending manual row must retain its own custom source location.");
        Contains("手打待確認備註", Sheet("xl/worksheets/sheet2.xml"),
            "Manual note did not enter pending review.");
        Contains("單顆(無盒)(霧面袋)", Sheet("xl/worksheets/sheet1.xml"),
            "The 192199 manual item did not map to its fixed product name.");
        Contains("3入條裝(無盒)(霧面袋)", Sheet("xl/worksheets/sheet1.xml"),
            "The 1902042 manual item did not map to its fixed product name.");
        Contains("6入鐵盒2024年鐵盒(黑金袋)", Sheet("xl/worksheets/sheet1.xml"),
            "The 1902197 manual item did not map to its fixed product name.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }

    return Task.CompletedTask;
}

static async Task PosMissingPickupDateUsesPersistentFirstImportSchedule()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-pos-first-seen-{Guid.NewGuid():N}");
    Directory.CreateDirectory(directory);
    var sourcePath = Path.Combine(directory, "pos.xlsx");
    var statePath = Path.Combine(directory, "pos-first-seen.json");
    try
    {
        void WriteWorkbook(bool includeNewOrder)
        {
            if (File.Exists(sourcePath)) File.Delete(sourcePath);
            var rows = new List<object?[]>
            {
                new object?[] { "預購取貨日", "訂單來源", "訂單編號", "SKU", "貨號", "品名", "出貨數", "系統內部備註" },
                new object?[] { "", "POS 正-新光A4", "POS-NO-DATE-1", "4710964232107", "1902072", "單顆", 2m, "來源舊備註" },
                new object?[] { "2026-08-22(六)", "POS 正-京站", "POS-HAS-DATE", "4710964232411", "1902248", "六入", 3m, "原備註" }
            };
            if (includeNewOrder)
            {
                rows.Add(new object?[] { "", "POS 正-夢時代", "POS-NO-DATE-2", "4710964232565", "1902247", "九入", 4m, "" });
            }
            CreateSimpleWorkbook(sourcePath, rows);
        }

        WriteWorkbook(includeNewOrder: false);
        var firstStore = new PosFirstImportStore(statePath);
        var first = await PosSpreadsheetImportWorkflow.ImportAsync(
            sourcePath,
            new DateTimeOffset(2026, 8, 12, 9, 0, 0, TimeSpan.FromHours(8)),
            firstStore);
        var firstMissing = first.Rows.Single(row => row.ExternalOrderNo == "POS-NO-DATE-1");
        Equal("2026-08-20", firstMissing.OriginalDeliveryDate,
            "A missing POS pickup date first seen on Wednesday must deliver the Thursday after next.");
        Equal("2026-08-13", firstMissing.OrderDate,
            "A missing POS pickup date first seen on Wednesday must use the next Thursday as order date.");
        Equal("未指定預購日", firstMissing.Note,
            "A missing POS pickup date must use the agreed Excel note.");
        Equal("未指定預購日", first.ExportEntries.Single(entry => entry.ExternalOrderNo == "POS-NO-DATE-1").Note,
            "The missing-date note must enter the formal Excel projection.");
        var dated = first.Rows.Single(row => row.ExternalOrderNo == "POS-HAS-DATE");
        Equal("2026-08-20", dated.OriginalDeliveryDate,
            "A POS row with a pickup date must retain the existing pickup-minus-one adjustment.");
        Equal("2026-08-13", dated.OrderDate,
            "A POS row with a pickup date must retain the existing order-date rule.");
        Equal("原備註", dated.Note, "A dated POS row must retain its system-internal note.");

        WriteWorkbook(includeNewOrder: true);
        var reloaded = await PosSpreadsheetImportWorkflow.ImportAsync(
            sourcePath,
            new DateTimeOffset(2026, 8, 14, 9, 0, 0, TimeSpan.FromHours(8)),
            new PosFirstImportStore(statePath));
        var stable = reloaded.Rows.Single(row => row.ExternalOrderNo == "POS-NO-DATE-1");
        Equal("2026-08-20", stable.OriginalDeliveryDate,
            "Reloading on another day must preserve the first inferred delivery date.");
        Equal("2026-08-13", stable.OrderDate,
            "Reloading on another day must preserve the first inferred order date.");
        var newlySeen = reloaded.Rows.Single(row => row.ExternalOrderNo == "POS-NO-DATE-2");
        Equal("2026-08-25", newlySeen.OriginalDeliveryDate,
            "A newly added missing-date row must receive its own first-import schedule.");
        Equal("2026-08-18", newlySeen.OrderDate,
            "A first import on Friday must use the next Tuesday as order date.");

        var concurrentPath = Path.Combine(directory, "pos-concurrent.xlsx");
        CreateSimpleWorkbook(concurrentPath,
        [
            ["預購取貨日", "訂單來源", "訂單編號", "SKU", "貨號", "品名", "出貨數"],
            ["", "POS 正-新光A4", "POS-CONCURRENT", "4710964232107", "1902072", "單顆", 1m]
        ]);
        var firstWriter = PosSpreadsheetImportWorkflow.ImportAsync(
            concurrentPath,
            new DateTimeOffset(2026, 8, 12, 9, 0, 0, TimeSpan.FromHours(8)),
            new PosFirstImportStore(statePath));
        var secondWriter = PosSpreadsheetImportWorkflow.ImportAsync(
            concurrentPath,
            new DateTimeOffset(2026, 8, 13, 9, 0, 0, TimeSpan.FromHours(8)),
            new PosFirstImportStore(statePath));
        var concurrent = await Task.WhenAll(firstWriter, secondWriter);
        True(concurrent.Select(batch => batch.Rows.Single().OrderDate).Distinct().Count() == 1,
            "Concurrent POS imports must converge on one persisted first-import schedule.");
        var persisted = await PosSpreadsheetImportWorkflow.ImportAsync(
            concurrentPath,
            new DateTimeOffset(2026, 8, 20, 9, 0, 0, TimeSpan.FromHours(8)),
            new PosFirstImportStore(statePath));
        Equal(concurrent[0].Rows.Single().OrderDate, persisted.Rows.Single().OrderDate,
            "A concurrent first-import race must not lose the persisted schedule.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
}

static Task PosSpreadsheetMapsExactChannelsDatesProductsAndQuantities()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-pos-{Guid.NewGuid():N}");
    Directory.CreateDirectory(directory);
    var sourcePath = Path.Combine(directory, "pos.xlsx");
    try
    {
        CreateSimpleWorkbook(sourcePath,
        [
            ["百貨專櫃門市預購報表"],
            ["匯出資訊"],
            [],
            ["預購取貨日", "訂單來源", "訂單編號", "SKU", "貨號", "品名", "出貨數",
                "買家備註", "托運單備註", "明細單備註", "系統內部備註", "掃碼備註"],
            ["2026-08-15(六)", "POS 正-新光A4", "POS-001", "4710964232107", "1902072", "芝麻蛋黃酥(1入)組", 2m,
                "不可匯出-買家", "不可匯出-托運", "不可匯出-明細", "POS系統備註-陳列盒", "不可匯出-掃碼"],
            ["2026-08-15(六)", "POS 正-台中高鐵", "POS-002", "1902072*3", "", "芝麻蛋黃酥3顆組", 3m, "", "", "", "", ""],
            ["2026-08-15(六)", "POS 正-夢時代", "POS-003", "4710964232411", "1902248", "來源端已修改的六入商品名稱", 4m, "", "", "", "", ""],
            ["2026-08-15(六)", "POS 正-京站", "POS-004", "4710964232565", "1902247", "芝初黑芝麻Q潤蛋黃酥禮盒(9入)", 5m, "", "", "", "", ""],
            ["2026-08-15(六)", "POS 臨新竹大全聯", "POS-005", "1902247-B", "", "【2盒】芝初黑芝麻Q潤蛋黃酥9入禮盒", 6m, "", "", "", "", ""],
            ["2026-08-15(六)", "POS 正-京站", "POS-X", "1201105*2", "", "高鈣黑芝麻粉隨手包7G兩入組", 7m,
                "", "", "", "非目標備註", ""]
        ]);

        var importedAt = new DateTimeOffset(2026, 8, 7, 9, 0, 0, TimeSpan.FromHours(8));
        var batch = SpreadsheetOrderImporter.ReadPos(sourcePath, importedAt);
        Equal("pos_excel", batch.SourceCode, "POS source code is not isolated from Flavor POS or other imports.");
        Equal(3, batch.Rows.Count, "POS rows without an item number must be excluded even when the name mentions egg-yolk pastry.");
        Equal(3, batch.ExcludedRowCount, "POS blank-item-number and non-target exclusion count is incorrect.");
        Equal("POS 正-新光A4", batch.Rows[0].ChannelName,
            "The complete POS order-source name must be preserved.");
        True(batch.Rows.All(row => row.OriginalDeliveryDate == "2026-08-13"),
            "POS delivery date must move from pickup minus one day to the nearest Tuesday or Thursday.");
        True(batch.Rows.All(row => row.OrderDate == "2026-08-06"),
            "POS order date must be exactly one week before the adjusted delivery date.");
        True(batch.Rows.Select(row => row.Quantity).SequenceEqual([2m, 4m, 5m]),
            "POS shipment quantities must remain numeric and unchanged.");
        Equal("POS系統備註-陳列盒", batch.Rows[0].Note,
            "POS system-internal note mapping is incorrect.");
        Equal("POS系統備註-陳列盒", batch.ExportEntries[0].Note,
            "POS system-internal note must enter the export projection.");
        Equal("來源端已修改的六入商品名稱", batch.Rows.Single(row => row.Sku == "1902248").ProductName,
            "The POS preview must preserve the source product name even when classification uses only the item number.");
        Equal("pos:4D030B253668ECAAE3A78025",
            batch.ExportEntries.Single(entry => entry.Sku == "1902248").ExternalLineKey,
            "The item-number-only rule must preserve the previously published POS identity for manual overrides.");
        Equal(3, batch.ExportEntries.Select(entry => entry.ExternalLineKey)
            .Distinct(StringComparer.OrdinalIgnoreCase).Count(),
            "POS product identity must use the stable item number.");

        var originalNamePath = Path.Combine(directory, "pos-original-name.xlsx");
        var changedNamePath = Path.Combine(directory, "pos-changed-name.xlsx");
        static object?[][] PosIdentityWorkbook(string productName) =>
        [
            ["預購取貨日", "訂單來源", "訂單編號", "SKU", "貨號", "品名", "出貨數"],
            ["2026-08-15(六)", "POS 正-夢時代", "POS-STABLE", "4710964232411", "1902248", productName, 4m]
        ];
        CreateSimpleWorkbook(originalNamePath, PosIdentityWorkbook("原始六入商品名稱"));
        CreateSimpleWorkbook(changedNamePath, PosIdentityWorkbook("修改後六入商品名稱"));
        var originalIdentity = SpreadsheetOrderImporter.ReadPos(originalNamePath, importedAt).ExportEntries.Single();
        var changedIdentity = SpreadsheetOrderImporter.ReadPos(changedNamePath, importedAt).ExportEntries.Single();
        Equal(originalIdentity.ExternalLineKey, changedIdentity.ExternalLineKey,
            "Changing only a POS source product name must not change the source-line identity.");
        Equal(originalIdentity.EntryId, changedIdentity.EntryId,
            "Changing only a POS source product name must not break manual-override linkage.");

        var outputPath = Path.Combine(directory, "pos-output.xlsx");
        GiftBoxWorkbookExporter.Export(outputPath, batch.ExportEntries);
        using var output = ZipFile.OpenRead(outputPath);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using var detailInput = output.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var detail = XDocument.Load(detailInput);
        var xml = detail.ToString();
        Contains("POS系統備註-陳列盒", xml,
            "POS system-internal note did not enter order details.");
        DoesNotContain("不可匯出-", xml,
            "POS notes other than 系統內部備註 must not enter order details.");
        DoesNotContain("POS-002", xml,
            "A POS row without an item number must not enter order details.");
        DoesNotContain("POS-005", xml,
            "A two-box POS row without an item number must not enter order details.");
        foreach (var channel in new[]
                 {
                     "POS 正-新光A4", "POS 正-夢時代", "POS 正-京站"
                 })
        {
            Contains(channel, xml, $"POS channel was shortened or missing: {channel}.");
        }
        Equal(3, detail.Descendants(ns + "row").Skip(1).Count(),
            "Only POS rows with an accepted nonblank item number may enter order details.");
        var productNames = detail.Descendants(ns + "row").Skip(1)
            .Select(row => string.Concat(row.Elements(ns + "c")
                .Single(cell => ((string?)cell.Attribute("r"))?.StartsWith("B", StringComparison.Ordinal) == true)
                .Descendants(ns + "t").Select(text => text.Value)))
            .ToArray();
        Equal(1, productNames.Count(name => name == "9入紙盒3入禮盒紙盒(黑金袋＋3包7g芝麻粉)"),
            "Only the regular nine-piece POS product with an item number may be exported.");
        True(productNames.All(name => !name.Contains("2盒", StringComparison.Ordinal)),
            "The POS product name must not add a two-box suffix that users may mistake for quantity conversion.");
        var locations = detail.Descendants(ns + "row").Skip(1)
            .Select(row => string.Concat(row.Elements(ns + "c")
                .Single(cell => ((string?)cell.Attribute("r"))?.StartsWith("G", StringComparison.Ordinal) == true)
                .Descendants(ns + "t").Select(text => text.Value)))
            .ToArray();
        True(locations.All(location => location == "躉泰"),
            "Every POS report row must default to 躉泰.");
        var quantities = detail.Descendants(ns + "c")
            .Where(cell => ((string?)cell.Attribute("r"))?.StartsWith("E", StringComparison.Ordinal) == true &&
                           (string?)cell.Attribute("r") != "E1")
            .ToArray();
        True(quantities.All(cell => (string?)cell.Attribute("t") is null),
            "POS quantities must be stored as Excel numeric cells.");
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
    return Task.CompletedTask;
}

static Task ManualSingleItemsUseSeparateWorksheet()
{
    var path = Path.Combine(Path.GetTempPath(), $"mid-autumn-single-sheet-{Guid.NewGuid():N}.xlsx");
    var at = new DateTimeOffset(2026, 8, 6, 9, 0, 0, TimeSpan.FromHours(8));
    OrderChangeEntry Entry(string id, string sku, decimal quantity) =>
        new(id, at, "manual_excel", "手打單 Excel", $"MANUAL-{id}", $"manual_excel:{sku}:1", "官網",
            null, null, sku, sku == "1902199" ? "蛋黃酥(1入)-無盒" : "蛋黃酥(1入)組-有盒",
            0m, quantity, quantity, "檔案匯入", "IMPORTED", "已匯入", false);
    try
    {
        GiftBoxWorkbookExporter.Export(path, [Entry("A", "1902199", 4m), Entry("B", "1902072", 5m)]);
        using var archive = ZipFile.OpenRead(path);
        using (var workbookReader = new StreamReader(
                   (archive.GetEntry("xl/workbook.xml")
                    ?? throw new InvalidOperationException("Workbook metadata is missing.")).Open(), Encoding.UTF8))
        {
            Contains("日期待確認", workbookReader.ReadToEnd(), "The pending-review worksheet is missing.");
        }

        using var sheetInput = (archive.GetEntry("xl/worksheets/sheet2.xml")
                                ?? throw new InvalidOperationException("Pending-review worksheet is missing.")).Open();
        var sheet = XDocument.Load(sheetInput);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Equal(3, sheet.Descendants(ns + "row").Count(), "The two single SKUs must remain separate rows.");
        var xml = sheet.ToString();
        Contains("單顆(無盒)(透明袋)", xml, "The unboxed single item is missing.");
        Contains("單顆(有盒)(霧面袋)", xml, "The boxed single item is missing.");
        var quantities = sheet.Descendants(ns + "c")
            .Where(cell => ((string?)cell.Attribute("r"))?.StartsWith("D", StringComparison.OrdinalIgnoreCase) == true &&
                           (string?)cell.Attribute("r") != "D1")
            .Select(cell => ((string?)cell.Attribute("t"), cell.Element(ns + "v")?.Value))
            .ToArray();
        True(quantities.All(value => value.Item1 is null), "Single-item quantities must be Excel numbers.");
        True(quantities.Select(value => value.Item2).ToHashSet().SetEquals(["4", "5"]),
            "Single-item quantities are incorrect.");
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
    Equal(3, GiftBoxSizePolicy.Resolve("1902226", null), "ERP three-piece item mapping is incorrect.");
    Equal(6, GiftBoxSizePolicy.Resolve("1902248", null), "ERP six-piece item mapping is incorrect.");
    Equal(9, GiftBoxSizePolicy.Resolve("1902247", null), "ERP nine-piece item mapping is incorrect.");
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

static async Task LocalHistoryResetIsRecoverable()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-reset-{Guid.NewGuid():N}");
    Directory.CreateDirectory(directory);
    var originalFiles = new Dictionary<string, string>
    {
        ["sync-status.json"] = "OLD-STATUS",
        ["order-snapshot.json"] = "OLD-SNAPSHOT",
        ["order-change-ledger.json"] = "OLD-LEDGER"
    };

    try
    {
        foreach (var file in originalFiles)
        {
            await File.WriteAllTextAsync(Path.Combine(directory, file.Key), file.Value);
        }
        await File.WriteAllTextAsync(Path.Combine(directory, "manual-orders.json"), "KEEP-MANUAL");

        var service = new LocalHistoryResetService(directory, ["site1", "site2"]);
        var reset = await service.BackupAndResetAsync(
            new DateTimeOffset(2026, 8, 5, 18, 0, 0, TimeSpan.FromHours(8)));

        foreach (var file in originalFiles)
        {
            True(!File.Exists(Path.Combine(directory, file.Key)),
                $"{file.Key} should leave the active data directory after reset.");
            Equal(file.Value, await File.ReadAllTextAsync(Path.Combine(reset.BackupDirectory, file.Key)),
                $"{file.Key} was not backed up before reset.");
        }
        Equal("KEEP-MANUAL", await File.ReadAllTextAsync(Path.Combine(directory, "manual-orders.json")),
            "Reset must not move manual orders, unrelated files, or credentials.");

        await File.WriteAllTextAsync(Path.Combine(directory, "sync-status.json"), "FAILED-REFRESH");
        await service.RestoreAsync(reset);
        foreach (var file in originalFiles)
        {
            Equal(file.Value, await File.ReadAllTextAsync(Path.Combine(directory, file.Key)),
                $"{file.Key} was not restored after a failed full refresh.");
        }
        Equal("FAILED-REFRESH", await File.ReadAllTextAsync(
                Path.Combine(reset.BackupDirectory, "failed-refresh", "sync-status.json")),
            "Partial refresh data should remain recoverable instead of being deleted during rollback.");

        try
        {
            await service.ResetAndRefreshAsync(
                new DateTimeOffset(2026, 8, 5, 19, 0, 0, TimeSpan.FromHours(8)),
                async cancellationToken =>
                {
                    await File.WriteAllTextAsync(
                        Path.Combine(directory, "order-snapshot.json"),
                        "SECOND-SOURCE-FAILED",
                        cancellationToken);
                    throw new IOException("site2 failed");
                });
            throw new InvalidOperationException("Expected recoverable full refresh failure.");
        }
        catch (LocalHistoryRefreshException exception)
        {
            True(exception.HistoryRestored,
                "A failed full refresh should report that the original history was restored.");
            Equal("OLD-SNAPSHOT", await File.ReadAllTextAsync(Path.Combine(directory, "order-snapshot.json")),
                "The reset-and-refresh coordinator did not restore the original snapshot.");
            Equal("SECOND-SOURCE-FAILED", await File.ReadAllTextAsync(Path.Combine(
                    exception.BackupDirectory, "failed-refresh", "order-snapshot.json")),
                "Partial second-source refresh data was not retained for diagnosis.");
        }
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
}

static async Task ManualOrdersPersistAuditAndExportSafely()
{
    var directory = Path.Combine(Path.GetTempPath(), $"mid-autumn-manual-{Guid.NewGuid():N}");
    var path = Path.Combine(directory, "manual-orders.json");
    var createdAt = new DateTimeOffset(2026, 8, 5, 9, 0, 0, TimeSpan.FromHours(8));
    try
    {
        var store = new ManualOrderStore(path);
        var added = await store.AddAsync(
            new ManualOrderInput(
                "蝦皮賣場",
                "4710964232435",
                2,
                new DateOnly(2026, 8, 15),
                false,
                "FACTORY-PRIVATE-NOTE",
                "HAND-001"),
            new ManualOrderActor("alice", "PC-A"),
            createdAt);

        var reopened = new ManualOrderStore(path);
        var active = await reopened.GetActiveAsync();
        Equal(1, active.Count, "A manual order was not persisted.");
        Equal("manual", active[0].SourceCode, "Manual order source should be canonical manual.");
        Equal(false, active[0].IsRegistered, "Manual registration should default to false.");
        Equal("alice", active[0].CreatedBy, "Manual order creator was not audited.");
        Equal("PC-A", active[0].CreatedOnComputer, "Manual order computer was not audited.");

        var updated = await reopened.SetRegistrationAsync(
            added.Id,
            true,
            new ManualOrderActor("bob", "PC-B"),
            createdAt.AddMinutes(5));
        Equal(true, updated.IsRegistered, "Manual registration status was not updated.");
        Equal("alice", updated.CreatedBy, "Updating registration must not overwrite the creator.");
        Equal("bob", updated.ModifiedBy, "Registration operator was not audited.");
        Equal("PC-B", updated.ModifiedOnComputer, "Registration computer was not audited.");
        var audit = await reopened.GetAuditAsync();
        Equal(2, audit.Count, "Add and registration change should each create an audit record.");
        Equal(false, audit[1].BeforeRegistered, "Registration audit lost the previous value.");
        Equal(true, audit[1].AfterRegistered, "Registration audit lost the new value.");

        var exportEntries = await reopened.GetExportEntriesAsync();
        Equal(1, exportEntries.Count, "The active manual order should enter Excel data once.");
        var exportEntry = exportEntries[0];
        Equal("HAND-001", exportEntry.ExternalOrderNo, "Manual external order number was not preserved.");
        Equal("蝦皮賣場", exportEntry.ChannelName, "Manual channel was not preserved.");
        Equal("2026-08-06", exportEntry.DeliveryDate, "Manual arrival date did not use the lead-time rule.");
        Equal(2m, exportEntry.QuantityChange, "Manual quantity is incorrect.");
        DoesNotContain("FACTORY-PRIVATE-NOTE", JsonSerializer.Serialize(exportEntries),
            "Manual notes must not enter Excel projection data.");
        var workbookPath = Path.Combine(directory, "manual-export.xlsx");
        GiftBoxWorkbookExporter.Export(workbookPath, exportEntries);
        using (var archive = ZipFile.OpenRead(workbookPath))
        {
            string ReadWorksheet(string path)
            {
                using var reader = new StreamReader(
                    (archive.GetEntry(path)
                     ?? throw new InvalidOperationException($"Workbook worksheet {path} is missing.")).Open(),
                    Encoding.UTF8);
                return reader.ReadToEnd();
            }

            var threePieceXml = ReadWorksheet("xl/worksheets/sheet1.xml");
            Contains("蝦皮賣場", threePieceXml, "Manual order channel did not enter the workbook aggregate.");
            Contains(ExcelSerial("2026-08-06"), threePieceXml, "Manual adjusted date did not enter the workbook aggregate.");
            foreach (var worksheetPath in new[]
                     {
                         "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml",
                         "xl/worksheets/sheet3.xml"
                     })
            {
                DoesNotContain("FACTORY-PRIVATE-NOTE", ReadWorksheet(worksheetPath),
                    $"Manual note leaked into {worksheetPath}.");
            }
        }

        try
        {
            await reopened.AddAsync(
                new ManualOrderInput(
                    "蝦皮賣場", "4710964232435", 0, new DateOnly(2026, 8, 15), false, null, null),
                new ManualOrderActor("alice", "PC-A"),
                createdAt);
            throw new InvalidOperationException("Expected zero manual quantity to be rejected.");
        }
        catch (ArgumentOutOfRangeException)
        {
            Equal(1, (await reopened.GetActiveAsync()).Count,
                "Invalid manual input must not change persisted orders.");
        }
    }
    finally
    {
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
    }
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

static void CreateSimpleWorkbook(
    string path,
    IReadOnlyList<IReadOnlyList<object?>> rows,
    bool prependBlankSheet = false)
{
    const string spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
    void Write(string entryPath, string content)
    {
        var entry = archive.CreateEntry(entryPath);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    var secondSheetContentType = prependBlankSheet
        ? "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        : string.Empty;
    Write("[Content_Types].xml", $"""
        <?xml version="1.0" encoding="UTF-8"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          {secondSheetContentType}
        </Types>
        """);
    Write("_rels/.rels", """
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
        """);
    var workbookSheets = prependBlankSheet
        ? "<sheet name=\"說明\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"資料\" sheetId=\"2\" r:id=\"rId2\"/>"
        : "<sheet name=\"工作表1\" sheetId=\"1\" r:id=\"rId1\"/>";
    Write("xl/workbook.xml", $"""
        <?xml version="1.0" encoding="UTF-8"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>{workbookSheets}</sheets>
        </workbook>
        """);
    var secondRelationship = prependBlankSheet
        ? "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>"
        : string.Empty;
    Write("xl/_rels/workbook.xml.rels", $"""
        <?xml version="1.0" encoding="UTF-8"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          {secondRelationship}
        </Relationships>
        """);

    XNamespace ns = spreadsheet;
    var sheetData = new XElement(ns + "sheetData");
    for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
    {
        var rowNumber = rowIndex + 1;
        var row = new XElement(ns + "row", new XAttribute("r", rowNumber));
        for (var columnIndex = 0; columnIndex < rows[rowIndex].Count; columnIndex++)
        {
            var value = rows[rowIndex][columnIndex];
            if (value is null || value is string text && text.Length == 0)
            {
                continue;
            }

            var reference = $"{(char)('A' + columnIndex)}{rowNumber}";
            row.Add(value is decimal number
                ? new XElement(ns + "c", new XAttribute("r", reference),
                    new XElement(ns + "v", number.ToString(System.Globalization.CultureInfo.InvariantCulture)))
                : new XElement(ns + "c", new XAttribute("r", reference), new XAttribute("t", "inlineStr"),
                    new XElement(ns + "is", new XElement(ns + "t", value.ToString()))));
        }
        sheetData.Add(row);
    }

    var worksheet = new XDocument(new XDeclaration("1.0", "UTF-8", null),
        new XElement(ns + "worksheet", sheetData));
    if (prependBlankSheet)
    {
        Write("xl/worksheets/sheet1.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>說明頁</t></is></c></row></sheetData></worksheet>");
    }
    var sheetEntry = archive.CreateEntry(prependBlankSheet
        ? "xl/worksheets/sheet2.xml"
        : "xl/worksheets/sheet1.xml");
    using var stream = sheetEntry.Open();
    worksheet.Save(stream);
}

static HttpResponseMessage Json(string value) => new(HttpStatusCode.OK)
{
    Content = new StringContent(value, Encoding.UTF8, "application/json")
};

static void True(bool value, string message)
{
    if (!value) throw new InvalidOperationException(message);
}

static string ExcelSerial(string date) =>
    DateOnly.ParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture)
        .ToDateTime(TimeOnly.MinValue)
        .ToOADate()
        .ToString(CultureInfo.InvariantCulture);

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
