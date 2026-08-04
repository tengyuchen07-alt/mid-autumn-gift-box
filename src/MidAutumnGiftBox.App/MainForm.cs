using System.Data;
using System.ComponentModel;
using MidAutumnGiftBox.Core;

namespace MidAutumnGiftBox.App;

internal sealed class MainForm : Form
{
    private static readonly IReadOnlyDictionary<string, string> PreviewColumnNames =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["source"] = "訂單來源",
            ["shop_name"] = "賣場名稱",
            ["source_key"] = "通路代碼",
            ["order_no"] = "訂單編號",
            ["order_date"] = "訂單成立日",
            ["arrival_date"] = "指定到貨日（API）",
            ["ship_window_start"] = "出貨區間起日",
            ["ship_window_end"] = "出貨區間迄日",
            ["derived_shipping_date"] = "指定出貨日（推導）",
            ["shipping_date_source"] = "出貨日來源",
            ["products"] = "商品明細"
        };

    private static readonly WmsSource[] Sources =
    [
        new("site1", "網站1－睿驛", new Uri("https://reyi-distribution.wms.changliu.com.tw")),
        new("site2", "網站2－Flavor", new Uri("https://flavor.wms.changliu.com.tw"))
    ];

    private readonly ComboBox _sourceBox = new() { DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly TextBox _apiIdBox = new();
    private readonly TextBox _apiKeyBox = new() { UseSystemPasswordChar = true };
    private readonly Label _savedStateLabel = new() { AutoSize = true, ForeColor = Color.DimGray };
    private readonly Label _syncStatusLabel = new()
    {
        AutoSize = true,
        ForeColor = Color.FromArgb(75, 85, 99),
        Text = "此來源尚無同步紀錄。"
    };
    private readonly Label _previewMetadataLabel = new()
    {
        AutoSize = true,
        ForeColor = Color.FromArgb(75, 85, 99),
        Text = "尚未查詢。",
        Padding = new Padding(2, 2, 0, 6)
    };
    private readonly Label _orderRangeLabel = new()
    {
        AutoSize = true,
        ForeColor = Color.FromArgb(75, 85, 99),
        Margin = new Padding(18, 8, 8, 0)
    };
    private readonly Button _saveButton = new() { Text = "安全保存憑證", AutoSize = true };
    private readonly Button _shopsButton = new() { Text = "取得店舖清單", AutoSize = true };
    private readonly Button _ordersButton = new() { Text = "載入蛋黃酥待處理訂單", AutoSize = true, Enabled = false };
    private readonly Button _exportButton = new() { Text = "匯出 Excel", AutoSize = true, Enabled = false };
    private readonly DataGridView _grid = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AllowUserToAddRows = false,
        AllowUserToDeleteRows = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
        BackgroundColor = Color.White,
        BorderStyle = BorderStyle.None
    };
    private readonly RichTextBox _jsonBox = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        Font = new Font("Consolas", 10F),
        BackColor = Color.White,
        BorderStyle = BorderStyle.None,
        WordWrap = false
    };
    private readonly ToolStripStatusLabel _statusLabel = new("準備就緒");
    private readonly ToolStripProgressBar _progress = new() { Style = ProgressBarStyle.Marquee, Visible = false };

    private readonly CredentialManager _credentialManager = new(new WindowsCredentialVault());
    private readonly SyncStatusStore _syncStatusStore;
    private readonly OrderSnapshotStore _orderSnapshotStore;
    private readonly WmsApiClient _wmsClient;
    private readonly HashSet<string> _shopValidatedSources = new(StringComparer.OrdinalIgnoreCase);
    private PreviewResult? _lastPreview;

    public MainForm()
    {
        var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("MidAutumnGiftBox/1.0");
        _wmsClient = new WmsApiClient(httpClient);
        _syncStatusStore = new SyncStatusStore(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MidAutumnGiftBox",
            "sync-status.json"));
        _orderSnapshotStore = new OrderSnapshotStore(Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MidAutumnGiftBox",
            "order-snapshot.json"));

        Text = "中秋禮盒－WMS 資料預覽";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(980, 680);
        Size = new Size(1180, 780);
        Font = new Font("Microsoft JhengHei UI", 10F);
        BackColor = Color.FromArgb(245, 247, 250);

        Controls.Add(BuildLayout());
        _sourceBox.DataSource = Sources;
        _sourceBox.DisplayMember = nameof(WmsSource.Name);
        _sourceBox.SelectedIndexChanged += async (_, _) => await LoadCredentialStatusAsync();
        _apiIdBox.TextChanged += (_, _) => InvalidateShopValidationForCurrentSource();
        _apiKeyBox.TextChanged += (_, _) => InvalidateShopValidationForCurrentSource();
        _saveButton.Click += async (_, _) => await SaveCredentialsAsync();
        _shopsButton.Click += async (_, _) => await LoadShopsAsync();
        _ordersButton.Click += async (_, _) => await LoadOrdersAsync();
        _exportButton.Click += (_, _) => ExportPreview();
        _orderRangeLabel.Text =
            $"訂單成立日：{WmsQueryPolicy.InitialOrderDate:yyyy/MM/dd} ～ {DateTime.Today:yyyy/MM/dd}；" +
            "顯示蛋黃酥／指定 SKU，到貨日可空白";

        Shown += async (_, _) => await LoadCredentialStatusAsync();
    }

    private Control BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(18)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var title = new Label
        {
            Text = "WMS 真實資料預覽",
            Font = new Font("Microsoft JhengHei UI", 20F, FontStyle.Bold),
            AutoSize = true,
            ForeColor = Color.FromArgb(31, 41, 55),
            Margin = new Padding(0, 0, 0, 14)
        };

        var settings = new TableLayoutPanel
        {
            AutoSize = true,
            Dock = DockStyle.Top,
            ColumnCount = 6,
            BackColor = Color.White,
            Padding = new Padding(14),
            Margin = new Padding(0, 0, 0, 14)
        };
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));

        _sourceBox.Dock = DockStyle.Fill;
        _apiIdBox.Dock = DockStyle.Fill;
        _apiKeyBox.Dock = DockStyle.Fill;
        _apiKeyBox.PlaceholderText = "留白可使用已保存的 Key";

        settings.Controls.Add(FieldLabel("資料來源"), 0, 0);
        settings.Controls.Add(_sourceBox, 1, 0);
        settings.Controls.Add(FieldLabel("API ID"), 2, 0);
        settings.Controls.Add(_apiIdBox, 3, 0);
        settings.Controls.Add(FieldLabel("API Key"), 4, 0);
        settings.Controls.Add(_apiKeyBox, 5, 0);
        settings.Controls.Add(_savedStateLabel, 1, 1);
        settings.SetColumnSpan(_savedStateLabel, 3);
        settings.Controls.Add(_saveButton, 4, 1);
        settings.SetColumnSpan(_saveButton, 2);
        settings.Controls.Add(_syncStatusLabel, 1, 2);
        settings.SetColumnSpan(_syncStatusLabel, 5);

        var actionPanel = new FlowLayoutPanel
        {
            AutoSize = true,
            Dock = DockStyle.Top,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 4, 0, 8)
        };
        actionPanel.Controls.Add(_shopsButton);
        actionPanel.Controls.Add(_orderRangeLabel);
        actionPanel.Controls.Add(_ordersButton);
        actionPanel.Controls.Add(_exportButton);

        var tabs = new TabControl { Dock = DockStyle.Fill };
        var tableTab = new TabPage("整理後表格") { BackColor = Color.White, Padding = new Padding(8) };
        var jsonTab = new TabPage("安全遮罩 JSON") { BackColor = Color.White, Padding = new Padding(8) };
        tableTab.Controls.Add(_grid);
        jsonTab.Controls.Add(_jsonBox);
        tabs.TabPages.Add(tableTab);
        tabs.TabPages.Add(jsonTab);

        var status = new StatusStrip { SizingGrip = false };
        status.Items.Add(_statusLabel);
        status.Items.Add(new ToolStripStatusLabel { Spring = true });
        status.Items.Add(_progress);

        var content = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1 };
        content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        content.Controls.Add(actionPanel, 0, 0);
        content.Controls.Add(_previewMetadataLabel, 0, 1);
        content.Controls.Add(tabs, 0, 2);

        root.Controls.Add(title, 0, 0);
        root.Controls.Add(settings, 0, 1);
        root.Controls.Add(content, 0, 2);
        root.Controls.Add(status, 0, 3);
        return root;
    }

    private static Label FieldLabel(string text) => new()
    {
        Text = text,
        AutoSize = true,
        Margin = new Padding(4, 8, 8, 0),
        ForeColor = Color.FromArgb(75, 85, 99)
    };

    private WmsSource SelectedSource => _sourceBox.SelectedItem as WmsSource
        ?? throw new InvalidOperationException("請選擇資料來源。");

    private async Task LoadCredentialStatusAsync()
    {
        if (_sourceBox.SelectedItem is not WmsSource source)
        {
            return;
        }

        var display = await _credentialManager.GetDisplayStatusAsync(source.Code);
        _apiIdBox.Text = display.ApiId;
        _apiKeyBox.Clear();
        _savedStateLabel.Text = display.HasSavedKey
            ? "已安全保存 API Key；畫面不會回顯 Key。"
            : "尚未保存憑證。";
        _savedStateLabel.ForeColor = display.HasSavedKey ? Color.SeaGreen : Color.DimGray;
        await LoadSyncStatusAsync(source.Code);
        UpdateOrderButtonState();
    }

    private async Task SaveCredentialsAsync()
    {
        await RunBusyAsync("正在安全保存憑證…", async () =>
        {
            var credentials = new ApiCredentials(_apiIdBox.Text.Trim(), _apiKeyBox.Text);
            await _credentialManager.SaveAsync(SelectedSource.Code, credentials);
            _shopValidatedSources.Remove(SelectedSource.Code);
            _apiKeyBox.Clear();
            await LoadCredentialStatusAsync();
            return "憑證已保存到目前 Windows 使用者的 Credential Manager。";
        });
    }

    private async Task LoadShopsAsync()
    {
        await RunBusyAsync("正在取得店舖清單…", async () =>
        {
            var credentials = await ResolveCredentialsAsync();
            var result = await _wmsClient.GetShopsAsync(SelectedSource, credentials);
            ShowPreview(result);
            _shopValidatedSources.Add(SelectedSource.Code);
            UpdateOrderButtonState();
            return result.Message + $" 共 {result.Rows.Count} 筆；現在可以查詢待處理訂單。";
        });
    }

    private async Task LoadOrdersAsync()
    {
        await RunBusyAsync("正在取得包含蛋黃酥品項的待處理訂單…", async () =>
        {
            if (!_shopValidatedSources.Contains(SelectedSource.Code))
            {
                throw new InvalidOperationException("請先成功取得這個來源的店舖清單，再查詢訂單。");
            }

            var source = SelectedSource;
            var startedAt = DateTimeOffset.Now;
            var throughDate = WmsQueryPolicy.GetTaipeiDate(startedAt);
            ApiCredentials? credentials = null;
            try
            {
                credentials = await ResolveCredentialsAsync();
                var result = await _wmsClient.GetPendingOrdersSinceInitialDateAsync(
                    source,
                    credentials,
                    throughDate);
                if (!result.IsSuccess)
                {
                    throw new WmsApiException(result.Message);
                }

                var finishedAt = DateTimeOffset.Now;
                await _orderSnapshotStore.ReplaceSourceAsync(
                    source.Code,
                    result.Rows,
                    finishedAt);
                ShowPreview(result);
                try
                {
                    await _syncStatusStore.RecordSuccessAsync(new SyncRunSummary(
                        source.Code,
                        source.Name,
                        WmsQueryPolicy.InitialOrderDate,
                        throughDate,
                        startedAt,
                        finishedAt,
                        result.Metadata.PageCount,
                        result.Metadata.RowCount));
                    await LoadSyncStatusAsync(source.Code);
                    return result.Message + "；本機訂單快照已更新。";
                }
                catch (Exception statusException) when (IsSyncStatusStorageError(statusException))
                {
                    ShowSyncStatusStorageError("訂單已載入，但本機同步摘要無法保存。", statusException);
                    return result.Message + "（本機同步摘要未保存）";
                }
            }
            catch (Exception exception) when (exception is WmsApiException or ArgumentException or
                                              InvalidOperationException or Win32Exception or IOException or
                                              UnauthorizedAccessException or System.Text.Json.JsonException)
            {
                try
                {
                    await _syncStatusStore.RecordFailureAsync(new SyncRunFailure(
                        source.Code,
                        source.Name,
                        WmsQueryPolicy.InitialOrderDate,
                        throughDate,
                        startedAt,
                        DateTimeOffset.Now,
                        RedactSyncError(
                            exception.Message,
                            _apiIdBox.Text,
                            _apiKeyBox.Text,
                            credentials?.ApiId,
                            credentials?.ApiKey,
                            BuildBasicCredential(_apiIdBox.Text, _apiKeyBox.Text),
                            BuildBasicCredential(credentials?.ApiId, credentials?.ApiKey))));
                    await LoadSyncStatusAsync(source.Code);
                }
                catch (Exception statusException) when (IsSyncStatusStorageError(statusException))
                {
                    ShowSyncStatusStorageError("同步失敗，且本機同步摘要無法保存。", statusException);
                }

                throw;
            }
        });
    }

    private async Task LoadSyncStatusAsync(string sourceCode)
    {
        SyncSourceStatus? status;
        try
        {
            status = await _syncStatusStore.GetAsync(sourceCode);
        }
        catch (Exception exception) when (IsSyncStatusStorageError(exception))
        {
            ShowSyncStatusStorageError("本機同步摘要無法讀取。", exception);
            return;
        }

        if (status is null)
        {
            _syncStatusLabel.Text = "此來源尚無同步紀錄。";
            _syncStatusLabel.ForeColor = Color.DimGray;
            return;
        }

        var resultText = status.Status == "success" ? "成功" : "失敗";
        var lastSuccessText = status.LastSuccessAt is null
            ? "尚無"
            : status.LastSuccessAt.Value.ToLocalTime().ToString("yyyy/MM/dd HH:mm:ss");
        _syncStatusLabel.Text =
            $"最近同步：{resultText}　查詢：{status.QueryFrom:yyyy/MM/dd}～{status.QueryThrough:yyyy/MM/dd}　" +
            $"分頁：{status.PageCount}　筆數：{status.RecordCount}　最後成功：{lastSuccessText}" +
            (status.Status == "failed" ? $"　原因：{status.ErrorSummary}" : string.Empty);
        _syncStatusLabel.ForeColor = status.Status == "success" ? Color.SeaGreen : Color.Firebrick;
    }

    private void ShowSyncStatusStorageError(string message, Exception exception)
    {
        _syncStatusLabel.Text = message;
        _syncStatusLabel.ForeColor = Color.Firebrick;
        _statusLabel.Text = message;
        System.Diagnostics.Debug.WriteLine($"Sync status storage error: {exception.GetType().Name}");
    }

    private static bool IsSyncStatusStorageError(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or System.Text.Json.JsonException;

    private static string RedactSyncError(string message, params string?[] secrets)
    {
        var safe = message;
        foreach (var secret in secrets.Where(value => !string.IsNullOrWhiteSpace(value)))
        {
            safe = safe.Replace(secret!, "[已遮罩]", StringComparison.OrdinalIgnoreCase);
        }

        return safe;
    }

    private static string? BuildBasicCredential(string? apiId, string? apiKey)
    {
        if (string.IsNullOrWhiteSpace(apiId) || string.IsNullOrWhiteSpace(apiKey))
        {
            return null;
        }

        return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes($"{apiId}:{apiKey}"));
    }

    private async Task<ApiCredentials> ResolveCredentialsAsync()
    {
        return await _credentialManager.ResolveForUseAsync(
            SelectedSource.Code,
            _apiIdBox.Text,
            _apiKeyBox.Text);
    }

    private void ShowPreview(PreviewResult result)
    {
        var metadata = result.Metadata;
        _previewMetadataLabel.Text =
            $"來源：{metadata.SourceName}　端點：{metadata.Endpoint}　" +
            $"查詢時間：{metadata.TestedAt.ToLocalTime():yyyy/MM/dd HH:mm:ss}　" +
            $"HTTP：{metadata.HttpStatusCode}　result.ok：{metadata.ResultOk.ToString().ToLowerInvariant()}　" +
            $"分頁：{metadata.PageCount}　筆數：{metadata.RowCount}　訊息：{metadata.ResultMessage}";

        if (!result.IsSuccess)
        {
            throw new WmsApiException(result.Message);
        }

        _grid.DataSource = BuildTable(result.Rows);
        _jsonBox.Text = result.SafeJson;
        var isOrderPreview = metadata.Endpoint.Equals(
            "/api_v1/order/order_query.php",
            StringComparison.OrdinalIgnoreCase);
        _lastPreview = isOrderPreview ? result : null;
        _exportButton.Enabled = isOrderPreview && result.Rows.Count > 0;
    }

    private void ExportPreview()
    {
        if (_lastPreview is null || _lastPreview.Rows.Count == 0)
        {
            MessageBox.Show(this, "請先載入有資料的預覽，再匯出 Excel。", "尚無可匯出資料",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var safeSourceName = string.Concat(_lastPreview.Metadata.SourceName
            .Where(character => !Path.GetInvalidFileNameChars().Contains(character)));
        using var dialog = new SaveFileDialog
        {
            Title = "匯出預覽資料",
            Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx",
            DefaultExt = "xlsx",
            AddExtension = true,
            OverwritePrompt = true,
            FileName = $"中秋禮盒預覽_{safeSourceName}_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            PreviewWorkbookExporter.Export(dialog.FileName, _lastPreview);
            _statusLabel.Text = $"Excel 已匯出：{dialog.FileName}";
            MessageBox.Show(this, "Excel 匯出完成。\n包含「訂單預覽」與「商品明細」兩張工作表。",
                "匯出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or
                                          InvalidOperationException or ArgumentException or System.Xml.XmlException)
        {
            _statusLabel.Text = "Excel 匯出失敗";
            MessageBox.Show(this, $"無法匯出 Excel：{exception.Message}", "匯出失敗",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static DataTable BuildTable(IReadOnlyList<IReadOnlyDictionary<string, string?>> rows)
    {
        var table = new DataTable();
        var columns = rows.SelectMany(row => row.Keys).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        foreach (var column in columns)
        {
            table.Columns.Add(PreviewColumnNames.TryGetValue(column, out var displayName) ? displayName : column);
        }

        foreach (var row in rows)
        {
            var dataRow = table.NewRow();
            foreach (var pair in row)
            {
                var columnName = PreviewColumnNames.TryGetValue(pair.Key, out var displayName)
                    ? displayName
                    : pair.Key;
                dataRow[columnName] = (object?)pair.Value ?? DBNull.Value;
            }

            table.Rows.Add(dataRow);
        }

        return table;
    }

    private async Task RunBusyAsync(string busyMessage, Func<Task<string>> action)
    {
        SetBusy(true, busyMessage);
        try
        {
            var completionMessage = await action();
            _statusLabel.Text = completionMessage;
        }
        catch (Exception exception) when (exception is WmsApiException or ArgumentException or InvalidOperationException or Win32Exception)
        {
            _statusLabel.Text = "操作失敗";
            MessageBox.Show(this, exception.Message, "無法完成操作", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally
        {
            SetBusy(false, _statusLabel.Text ?? "準備就緒");
        }
    }

    private void SetBusy(bool busy, string message)
    {
        _sourceBox.Enabled = !busy;
        _saveButton.Enabled = !busy;
        _shopsButton.Enabled = !busy;
        _ordersButton.Enabled = !busy &&
                                _sourceBox.SelectedItem is WmsSource source &&
                                _shopValidatedSources.Contains(source.Code);
        _exportButton.Enabled = !busy && _lastPreview is { Rows.Count: > 0 };
        _progress.Visible = busy;
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        _statusLabel.Text = message;
    }

    private void UpdateOrderButtonState()
    {
        _ordersButton.Enabled = _sourceBox.SelectedItem is WmsSource source &&
                                _shopValidatedSources.Contains(source.Code);
    }

    private void InvalidateShopValidationForCurrentSource()
    {
        if (_sourceBox.SelectedItem is WmsSource source)
        {
            _shopValidatedSources.Remove(source.Code);
            UpdateOrderButtonState();
        }
    }
}
