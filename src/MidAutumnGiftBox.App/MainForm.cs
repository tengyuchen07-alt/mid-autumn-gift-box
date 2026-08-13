using System.Data;
using System.ComponentModel;
using MidAutumnGiftBox.Core;

namespace MidAutumnGiftBox.App;

internal sealed class MainForm : Form
{
    private static readonly IReadOnlyDictionary<string, string> PreviewColumnNames =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["operational_diagnostic_fields"] = "Flavor 作業欄位名稱",
            ["operational_diagnostics"] = "Flavor 安全作業診斷",
            ["source"] = "訂單來源",
            ["shop_name"] = "賣場名稱",
            ["source_key"] = "通路代碼",
            ["note"] = "備註",
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
    private readonly Button _ordersButton = new() { Text = "重新載入完整訂單", AutoSize = true };
    private readonly Button _exportButton = new() { Text = "匯出／更新禮盒 Excel", AutoSize = true };
    private readonly Button _clearOverridesButton = new() { Text = "清除人工覆寫", AutoSize = true };
    private readonly Button _loadErpFileButton = new() { Text = "選擇 ERP Excel", AutoSize = true };
    private readonly Button _loadManualExcelButton = new() { Text = "選擇手打單 Excel", AutoSize = true };
    private readonly Button _loadPosFileButton = new() { Text = "選擇 POS機 Excel", AutoSize = true };
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
    private readonly DataGridView _erpGrid = CreateImportGrid();
    private readonly DataGridView _manualExcelGrid = CreateImportGrid();
    private readonly DataGridView _posGrid = CreateImportGrid();
    private readonly ToolStripStatusLabel _statusLabel = new("準備就緒");
    private readonly ToolStripProgressBar _progress = new() { Style = ProgressBarStyle.Marquee, Visible = false };

    private readonly CredentialManager _credentialManager = new(new WindowsCredentialVault());
    private readonly SyncStatusStore _syncStatusStore;
    private readonly OrderSnapshotStore _orderSnapshotStore;
    private readonly ManualOverrideStore _manualOverrideStore;
    private readonly PosFirstImportStore _posFirstImportStore;
    private readonly LocalHistoryResetService _historyResetService;
    private readonly WmsApiClient _wmsClient;
    private readonly StartupSpreadsheetInputs? _startupInputs;
    private readonly bool _runQuickWorkflow;
    private readonly HashSet<string> _shopValidatedSources = new(StringComparer.OrdinalIgnoreCase);
    private ImportedSpreadsheetBatch? _erpBatch;
    private ImportedSpreadsheetBatch? _manualExcelBatch;
    private ImportedSpreadsheetBatch? _posBatch;

    public MainForm(
        StartupSpreadsheetInputs? startupInputs = null,
        bool runQuickWorkflow = false)
    {
        _startupInputs = startupInputs;
        _runQuickWorkflow = runQuickWorkflow;
        var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("MidAutumnGiftBox/1.0");
        _wmsClient = new WmsApiClient(httpClient);
        var dataDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MidAutumnGiftBox");
        _syncStatusStore = new SyncStatusStore(Path.Combine(dataDirectory, "sync-status.json"));
        _orderSnapshotStore = new OrderSnapshotStore(Path.Combine(dataDirectory, "order-snapshot.json"));
        _manualOverrideStore = new ManualOverrideStore(Path.Combine(dataDirectory, "manual-overrides.json"));
        _posFirstImportStore = new PosFirstImportStore(Path.Combine(dataDirectory, "pos-first-import.json"));
        _historyResetService = new LocalHistoryResetService(
            dataDirectory,
            Sources.Select(source => source.Code).ToArray());

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
        _ordersButton.Click += async (_, _) => await LoadAllOrdersAsync();
        _exportButton.Click += async (_, _) => await ExportPreviewAsync();
        _clearOverridesButton.Click += async (_, _) => await ClearManualOverrideAsync();
        _loadErpFileButton.Click += async (_, _) => await LoadSpreadsheetAsync(SpreadsheetImportKind.Erp);
        _loadManualExcelButton.Click += async (_, _) => await LoadSpreadsheetAsync(SpreadsheetImportKind.Manual);
        _loadPosFileButton.Click += async (_, _) => await LoadSpreadsheetAsync(SpreadsheetImportKind.Pos);
        _orderRangeLabel.Text =
            $"每次依訂單成立日重抓 {WmsQueryPolicy.InitialOrderDate:yyyy/MM/dd} 至本次執行日；到貨日可空白";

        Shown += async (_, _) =>
        {
            SetBusy(true, _startupInputs is null ? "正在載入程式狀態…" : "正在準備自動載入固定 Excel…");
            try
            {
                await LoadCredentialStatusAsync();
                if (_startupInputs is not null)
                {
                    if (_runQuickWorkflow)
                    {
                        await RunQuickWorkflowAsync(_startupInputs);
                    }
                    else
                    {
                        await LoadStartupSpreadsheetsAsync(_startupInputs);
                    }
                }
            }
            catch (Exception exception) when (IsExpectedOperationException(exception))
            {
                _statusLabel.Text = _runQuickWorkflow
                    ? "快速流程失敗，未匯出新版 Excel。"
                    : "程式狀態載入失敗。";
                MessageBox.Show(
                    this,
                    _runQuickWorkflow
                        ? $"快速流程已停止：\n{exception.Message}\n\n未匯出新版「{StartupSpreadsheetInputs.OutputFileName}」。"
                        : $"無法載入程式狀態：{exception.Message}",
                    _runQuickWorkflow ? "快速流程失敗" : "載入失敗",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                SetBusy(false, _statusLabel.Text ?? "準備就緒");
            }
        };
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
            WrapContents = true,
            Padding = new Padding(0, 4, 0, 8)
        };
        actionPanel.Controls.Add(_shopsButton);
        actionPanel.Controls.Add(_orderRangeLabel);
        actionPanel.Controls.Add(_ordersButton);
        actionPanel.Controls.Add(_exportButton);
        actionPanel.Controls.Add(_clearOverridesButton);

        var tabs = new TabControl { Dock = DockStyle.Fill };
        var tableTab = new TabPage("整理後表格") { BackColor = Color.White, Padding = new Padding(8) };
        var manualTab = new TabPage("手打單") { BackColor = Color.White, Padding = new Padding(8) };
        var erpTab = new TabPage("ERP") { BackColor = Color.White, Padding = new Padding(8) };
        var posTab = new TabPage("POS機") { BackColor = Color.White, Padding = new Padding(8) };
        tableTab.Controls.Add(_grid);
        var manualLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
        manualLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        manualLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var manualActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top };
        manualActions.Controls.Add(_loadManualExcelButton);
        manualLayout.Controls.Add(manualActions, 0, 0);
        manualLayout.Controls.Add(_manualExcelGrid, 0, 1);
        manualTab.Controls.Add(manualLayout);
        var erpLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
        erpLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        erpLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var erpActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top };
        erpActions.Controls.Add(_loadErpFileButton);
        erpLayout.Controls.Add(erpActions, 0, 0);
        erpLayout.Controls.Add(_erpGrid, 0, 1);
        erpTab.Controls.Add(erpLayout);
        var posLayout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
        posLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        posLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        var posActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top };
        posActions.Controls.Add(_loadPosFileButton);
        posLayout.Controls.Add(posActions, 0, 0);
        posLayout.Controls.Add(_posGrid, 0, 1);
        posTab.Controls.Add(posLayout);
        tabs.TabPages.Add(tableTab);
        tabs.TabPages.Add(manualTab);
        tabs.TabPages.Add(erpTab);
        tabs.TabPages.Add(posTab);

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

    private static DataGridView CreateImportGrid() => new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AllowUserToAddRows = false,
        AllowUserToDeleteRows = false,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
        BackgroundColor = Color.White,
        BorderStyle = BorderStyle.None
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
        if (_sourceBox.SelectedItem is not WmsSource selected ||
            !selected.Code.Equals(source.Code, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _apiIdBox.Text = display.ApiId;
        _apiKeyBox.Clear();
        _savedStateLabel.Text = display.HasSavedKey
            ? "已安全保存 API Key；畫面不會回顯 Key。"
            : "尚未保存憑證。";
        _savedStateLabel.ForeColor = display.HasSavedKey ? Color.SeaGreen : Color.DimGray;
        await LoadSyncStatusAsync(source.Code);
        await LoadCommittedPreviewAsync(source);
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

    private async Task LoadAllOrdersAsync()
    {
        await RunBusyAsync("正在重新載入兩個網站的完整蛋黃酥訂單…", async () =>
        {
            var credentialsBySource = await ResolveAllCredentialsAsync();
            return await SynchronizeAllSourcesAsync(credentialsBySource);
        });
    }

    private async Task<IReadOnlyDictionary<string, ApiCredentials>> ResolveAllCredentialsAsync()
    {
        var credentialsBySource = new Dictionary<string, ApiCredentials>(StringComparer.OrdinalIgnoreCase);
        foreach (var source in Sources)
        {
            credentialsBySource[source.Code] = await ResolveCredentialsForSourceAsync(source);
        }

        return credentialsBySource;
    }

    private async Task<string> SynchronizeAllSourcesAsync(
        IReadOnlyDictionary<string, ApiCredentials> credentialsBySource)
    {
        using var workflowLock = await _historyResetService.AcquireWorkflowLockAsync();
        var reloadStartedAt = DateTimeOffset.Now;
        var sharedWindow = SyncWindowPolicy.Create(null, reloadStartedAt);
        var batches = new List<CompleteSourceReload>();
        foreach (var source in Sources)
        {
            batches.Add(await FetchCompleteSourceAsync(
                source,
                credentialsBySource[source.Code],
                sharedWindow));
        }

        await _orderSnapshotStore.ReplaceAllAsync(
            batches.SelectMany(batch => batch.Snapshots).ToArray());

        foreach (var batch in batches)
        {
            try
            {
                await _syncStatusStore.RecordSuccessAsync(new SyncRunSummary(
                    batch.Source.Code,
                    batch.Source.Name,
                    batch.Window.FromDate,
                    batch.Window.ThroughDate,
                    batch.StartedAt,
                    batch.FinishedAt,
                    batch.PageCount,
                    batch.RowCount));
            }
            catch (Exception statusException) when (IsSyncStatusStorageError(statusException))
            {
                ShowSyncStatusStorageError("完整訂單已載入，但本機同步摘要無法保存。", statusException);
            }
        }

        var selected = SelectedSource;
        await LoadSyncStatusAsync(selected.Code);
        await LoadCommittedPreviewAsync(selected);
        return "兩個網站完整訂單已一起更新；只保留目前仍在待處理頁面的訂單。";
    }

    private async Task<CompleteSourceReload> FetchCompleteSourceAsync(
        WmsSource source,
        ApiCredentials credentials,
        SyncWindow window)
    {
        var startedAt = window.StartedAt;
        var lockDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MidAutumnGiftBox",
            "locks");
        using var sourceLock = SourceSyncFileLock.TryAcquire(lockDirectory, source.Code)
            ?? throw new InvalidOperationException($"{source.Name} 同步進行中，請稍後再試。");
        try
        {
            var shops = await _wmsClient.GetShopsAsync(source, credentials);
            if (!shops.IsSuccess)
            {
                throw new WmsApiException(shops.Message);
            }

            _shopValidatedSources.Add(source.Code);
            var result = await _wmsClient.GetPendingOrdersAsync(
                source,
                credentials,
                window.FromDate,
                window.ThroughDate);
            if (!result.IsSuccess)
            {
                throw new WmsApiException(result.Message);
            }

            var finishedAt = DateTimeOffset.Now;
            var snapshots = OrderSnapshotStore.BuildSourceSnapshot(source.Code, result.Rows, finishedAt);
            return new CompleteSourceReload(
                source,
                window,
                startedAt,
                finishedAt,
                result.Metadata.PageCount,
                result.Metadata.RowCount,
                snapshots);
        }
        catch (Exception exception) when (exception is WmsApiException or ArgumentException or
                                          InvalidOperationException or InvalidDataException or
                                          Win32Exception or IOException or
                                          UnauthorizedAccessException or System.Text.Json.JsonException or
                                          KeyNotFoundException)
        {
            try
            {
                await _syncStatusStore.RecordFailureAsync(new SyncRunFailure(
                    source.Code,
                    source.Name,
                    window.FromDate,
                    window.ThroughDate,
                    startedAt,
                    DateTimeOffset.Now,
                    RedactSyncError(
                        exception.Message,
                        _apiIdBox.Text,
                        _apiKeyBox.Text,
                        credentials.ApiId,
                        credentials.ApiKey,
                        BuildBasicCredential(_apiIdBox.Text, _apiKeyBox.Text),
                        BuildBasicCredential(credentials.ApiId, credentials.ApiKey))));
                await LoadSyncStatusAsync(source.Code);
            }
            catch (Exception statusException) when (IsSyncStatusStorageError(statusException))
            {
                ShowSyncStatusStorageError("同步失敗，且本機同步摘要無法保存。", statusException);
            }

            throw;
        }
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

        if (_sourceBox.SelectedItem is not WmsSource selected ||
            !selected.Code.Equals(sourceCode, StringComparison.OrdinalIgnoreCase))
        {
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

    private async Task LoadCommittedPreviewAsync(WmsSource source)
    {
        try
        {
            var snapshot = await _orderSnapshotStore.GetAllAsync();
            if (_sourceBox.SelectedItem is not WmsSource selected ||
                !selected.Code.Equals(source.Code, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var sourceLines = snapshot
                .Where(line => line.SourceCode.Equals(source.Code, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (sourceLines.Length == 0)
            {
                _grid.DataSource = null;
                _previewMetadataLabel.Text = "此來源目前沒有待處理訂單。";
                return;
            }

            SyncSourceStatus? status = null;
            try
            {
                status = await _syncStatusStore.GetAsync(source.Code);
            }
            catch (Exception exception) when (IsSyncStatusStorageError(exception))
            {
                ShowSyncStatusStorageError("同步摘要無法讀取，已顯示最後提交的訂單快照。", exception);
            }

            if (_sourceBox.SelectedItem is WmsSource current &&
                current.Code.Equals(source.Code, StringComparison.OrdinalIgnoreCase))
            {
                ShowPreview(OrderSnapshotPreviewBuilder.Build(source, sourceLines, status));
            }
        }
        catch (Exception exception) when (IsSyncStatusStorageError(exception))
        {
            ShowSyncStatusStorageError("本機訂單快照無法讀取。", exception);
        }
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

    private async Task<ApiCredentials> ResolveCredentialsForSourceAsync(WmsSource source)
    {
        if (_sourceBox.SelectedItem is WmsSource selected &&
            selected.Code.Equals(source.Code, StringComparison.OrdinalIgnoreCase))
        {
            return await ResolveCredentialsAsync();
        }

        var stored = await _credentialManager.GetForUseAsync(source.Code);
        if (stored is null || !stored.IsComplete)
        {
            throw new InvalidOperationException(
                $"請先切換到「{source.Name}」並安全保存該網站的 API ID 與 API Key。");
        }

        return stored;
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
    }

    private async Task LoadSpreadsheetAsync(SpreadsheetImportKind kind)
    {
        var title = kind switch
        {
            SpreadsheetImportKind.Erp => "選擇 ERP 產出數量 Excel",
            SpreadsheetImportKind.Manual => "選擇手打單 Excel",
            SpreadsheetImportKind.Pos => "選擇百貨 POS機預購報表 Excel",
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
        using var dialog = new OpenFileDialog
        {
            Title = title,
            Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx",
            CheckFileExists = true,
            Multiselect = false
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        await RunBusyAsync(BusyMessage(kind), async () =>
        {
            var batch = await ImportSpreadsheetPathAsync(kind, dialog.FileName);
            return ImportSummary(batch);
        });
    }

    private async Task LoadStartupSpreadsheetsAsync(StartupSpreadsheetInputs inputs)
    {
        var successes = new List<string>();
        var failures = new List<string>();
        SetBusy(true, "正在自動載入固定 Excel…");
        try
        {
            foreach (var (kind, path) in new[]
                     {
                         (SpreadsheetImportKind.Erp, inputs.ErpPath),
                         (SpreadsheetImportKind.Manual, inputs.ManualPath),
                         (SpreadsheetImportKind.Pos, inputs.PosPath)
                     })
            {
                if (path is null) continue;
                try
                {
                    _statusLabel.Text = BusyMessage(kind);
                    var batch = await ImportSpreadsheetPathAsync(kind, path);
                    successes.Add(ImportSummary(batch));
                }
                catch (Exception exception) when (IsExpectedOperationException(exception))
                {
                    failures.Add($"{Path.GetFileName(path)}：{exception.Message}");
                }
            }

            var summary = new List<string>();
            if (successes.Count > 0)
            {
                summary.Add("自動載入成功：\n" + string.Join("\n", successes.Select(value => $"• {value}")));
            }
            if (inputs.MissingFileNames.Count > 0)
            {
                summary.Add("找不到檔案：\n" + string.Join("\n", inputs.MissingFileNames.Select(value => $"• {value}")));
            }
            if (failures.Count > 0)
            {
                summary.Add("載入失敗：\n" + string.Join("\n", failures.Select(value => $"• {value}")));
            }

            _statusLabel.Text = failures.Count == 0 && inputs.MissingFileNames.Count == 0
                ? "固定 Excel 已全部自動載入。"
                : "固定 Excel 自動載入完成，部分檔案需要處理。";
            MessageBox.Show(
                this,
                string.Join("\n\n", summary),
                "固定 Excel 自動載入結果",
                MessageBoxButtons.OK,
                failures.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }
        finally
        {
            SetBusy(false, _statusLabel.Text ?? "準備就緒");
        }
    }

    private async Task RunQuickWorkflowAsync(StartupSpreadsheetInputs inputs)
    {
        try
        {
            if (!inputs.HasAllRequiredInputs)
            {
                throw new InvalidOperationException(
                    "快速流程缺少必要檔案：\n" +
                    string.Join("\n", inputs.MissingFileNames.Select(fileName => $"• {fileName}")));
            }

            var importSummaries = new List<string>();
            foreach (var (kind, path) in new[]
                     {
                         (SpreadsheetImportKind.Erp, inputs.ErpPath!),
                         (SpreadsheetImportKind.Manual, inputs.ManualPath!),
                         (SpreadsheetImportKind.Pos, inputs.PosPath!)
                     })
            {
                _statusLabel.Text = BusyMessage(kind);
                var batch = await ImportSpreadsheetPathAsync(kind, path);
                importSummaries.Add(ImportSummary(batch));
            }

            _statusLabel.Text = "正在取得兩個網站的完整訂單…";
            var credentialsBySource = await ResolveAllCredentialsAsync();
            var synchronizationSummary = await SynchronizeAllSourcesAsync(credentialsBySource);

            _statusLabel.Text = "正在匯出固定中秋禮盒 Excel…";
            await ExportQuickPreviewAsync(inputs.OutputPath);
            _statusLabel.Text = $"快速流程完成：{inputs.OutputPath}";

            MessageBox.Show(
                this,
                "快速流程已完成。\n\n" +
                string.Join("\n", importSummaries.Select(summary => $"• {summary}")) +
                $"\n\n• {synchronizationSummary}\n• 已匯出：{inputs.OutputPath}\n\n" +
                "既有人工修改已保留；快速流程不會清除人工覆寫。",
                "快速流程完成",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch (Exception exception) when (IsExpectedOperationException(exception))
        {
            _statusLabel.Text = "快速流程失敗，未匯出新版 Excel。";
            MessageBox.Show(
                this,
                $"快速流程已停止：\n{exception.Message}\n\n未匯出新版「{StartupSpreadsheetInputs.OutputFileName}」。",
                "快速流程失敗",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private async Task ExportQuickPreviewAsync(string path)
    {
        using var workflowLock = await _historyResetService.AcquireWorkflowLockAsync();
        var snapshots = await _orderSnapshotStore.GetAllAsync();
        var wmsEntries = CurrentSnapshotOrderProjector.Project(
            snapshots,
            Sources.ToDictionary(
                source => source.Code,
                source => source.Name,
                StringComparer.OrdinalIgnoreCase),
            DateTimeOffset.Now);
        var changeEntries = wmsEntries
            .Concat(_erpBatch?.ExportEntries ?? [])
            .Concat(_manualExcelBatch?.ExportEntries ?? [])
            .Concat(_posBatch?.ExportEntries ?? [])
            .ToArray();
        if (changeEntries.Length == 0)
        {
            throw new InvalidOperationException("兩站及固定 Excel 都沒有可匯出的訂單資料。");
        }

        var manualState = await _manualOverrideStore.LoadAsync();
        manualState = await CaptureLastWorkbookChangesAsync(manualState, promptWhenMissing: false)
                      ?? manualState;
        var automaticEntries = GiftBoxWorkbookExporter.ResolveEntries(changeEntries, snapshots);
        var fullPath = Path.GetFullPath(path);
        if (manualState.LastWorkbookPath is null && File.Exists(fullPath))
        {
            manualState = CaptureFirstExistingWorkbook(fullPath, automaticEntries, manualState);
        }

        var effectiveEntries = ManualOverrideWorkflow.Apply(automaticEntries, manualState.Overrides);
        effectiveEntries = ManualOverrideWorkflow.ApplyAutomaticResets(
            effectiveEntries,
            manualState.PendingAutomaticResets ?? []);
        var preparedPath = PrepareWorkbook(fullPath, effectiveEntries);
        try
        {
            var exportedRows = GiftBoxWorkbookExporter.ReadEditableRows(preparedPath);
            await _manualOverrideStore.SaveAsync(manualState with
            {
                LastWorkbookPath = fullPath,
                LastExportedRows = exportedRows,
                PendingAutomaticResets = []
            });
            try
            {
                File.Move(preparedPath, fullPath, overwrite: true);
            }
            catch
            {
                await _manualOverrideStore.SaveAsync(manualState);
                throw;
            }
        }
        finally
        {
            if (File.Exists(preparedPath)) File.Delete(preparedPath);
        }
    }

    private static string PrepareWorkbook(
        string fullPath,
        IReadOnlyList<OrderChangeEntry> entries)
    {
        var directory = Path.GetDirectoryName(fullPath)
                        ?? throw new InvalidOperationException("固定匯出路徑缺少資料夾。");
        Directory.CreateDirectory(directory);
        var temporaryPath = Path.Combine(
            directory,
            $".{Path.GetFileNameWithoutExtension(fullPath)}.{Guid.NewGuid():N}.tmp.xlsx");
        try
        {
            if (File.Exists(fullPath))
            {
                File.Copy(fullPath, temporaryPath, overwrite: false);
            }

            GiftBoxWorkbookExporter.Export(temporaryPath, entries);
            return temporaryPath;
        }
        catch
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
            throw;
        }
    }

    private async Task<ImportedSpreadsheetBatch> ImportSpreadsheetPathAsync(
        SpreadsheetImportKind kind,
        string path)
    {
        var importedAt = DateTimeOffset.Now;
        var batch = kind == SpreadsheetImportKind.Pos
            ? await PosSpreadsheetImportWorkflow.ImportAsync(path, importedAt, _posFirstImportStore)
            : await Task.Run(() => kind switch
            {
                SpreadsheetImportKind.Erp => SpreadsheetOrderImporter.ReadErp(path, importedAt),
                SpreadsheetImportKind.Manual => SpreadsheetOrderImporter.ReadManual(path, importedAt),
                _ => throw new ArgumentOutOfRangeException(nameof(kind))
            });

        switch (kind)
        {
            case SpreadsheetImportKind.Erp:
                _erpBatch = batch;
                _erpGrid.DataSource = BuildImportTable(batch);
                break;
            case SpreadsheetImportKind.Manual:
                _manualExcelBatch = batch;
                _manualExcelGrid.DataSource = BuildImportTable(batch);
                break;
            case SpreadsheetImportKind.Pos:
                _posBatch = batch;
                _posGrid.DataSource = BuildImportTable(batch);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(kind));
        }

        return batch;
    }

    private static string BusyMessage(SpreadsheetImportKind kind) => kind switch
    {
        SpreadsheetImportKind.Erp => "正在讀取 ERP Excel…",
        SpreadsheetImportKind.Manual => "正在讀取手打單 Excel…",
        SpreadsheetImportKind.Pos => "正在讀取 POS機 Excel…",
        _ => throw new ArgumentOutOfRangeException(nameof(kind))
    };

    private static string ImportSummary(ImportedSpreadsheetBatch batch) =>
        $"已讀取 {batch.FileName}：預覽 {batch.Rows.Count} 列、可匯出 {batch.ExportEntries.Count} 列" +
        (batch.ExcludedRowCount > 0 ? $"、排除非目標商品 {batch.ExcludedRowCount} 列。" : "。");

    private static bool IsExpectedOperationException(Exception exception) =>
        exception is WmsApiException or ArgumentException or InvalidOperationException or Win32Exception or
            InvalidDataException or IOException or UnauthorizedAccessException or System.Text.Json.JsonException or
            System.Xml.XmlException or KeyNotFoundException;

    private static DataTable BuildImportTable(ImportedSpreadsheetBatch batch)
    {
        var table = new DataTable();
        table.Columns.Add("資料列", typeof(int));
        table.Columns.Add("通路別");
        table.Columns.Add("訂單編號");
        table.Columns.Add("料號");
        table.Columns.Add("品項");
        table.Columns.Add("地點");
        table.Columns.Add("原定指定到貨日");
        table.Columns.Add("下單日");
        table.Columns.Add("數量", typeof(decimal));
        table.Columns.Add("備註");
        foreach (var row in batch.Rows)
        {
            table.Rows.Add(
                row.SourceRowNumber,
                row.ChannelName,
                row.ExternalOrderNo,
                row.Sku,
                row.ProductName,
                row.SourceLocation ?? string.Empty,
                row.OriginalDeliveryDate ?? string.Empty,
                row.OrderDate ?? string.Empty,
                row.Quantity is null ? DBNull.Value : row.Quantity.Value,
                row.Note ?? string.Empty);
        }
        return table;
    }

    private async Task ExportPreviewAsync()
    {
        IReadOnlyList<OrderChangeEntry> changeEntries;
        IReadOnlyList<OrderLineSnapshot> snapshots;
        ManualOverrideState manualState;
        try
        {
            using var workflowLock = await _historyResetService.AcquireWorkflowLockAsync();
            snapshots = await _orderSnapshotStore.GetAllAsync();
            var wmsEntries = CurrentSnapshotOrderProjector.Project(
                snapshots,
                Sources.ToDictionary(
                    source => source.Code,
                    source => source.Name,
                    StringComparer.OrdinalIgnoreCase),
                DateTimeOffset.Now);
            changeEntries = wmsEntries
                .Concat(_erpBatch?.ExportEntries ?? [])
                .Concat(_manualExcelBatch?.ExportEntries ?? [])
                .Concat(_posBatch?.ExportEntries ?? [])
                .ToArray();
            manualState = await _manualOverrideStore.LoadAsync();
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or
                                          System.Text.Json.JsonException)
        {
            MessageBox.Show(this, $"無法讀取異動紀錄：{exception.Message}", "匯出失敗",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        ManualOverrideState? capturedState;
        try
        {
            capturedState = await CaptureLastWorkbookChangesAsync(manualState);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or
                                          InvalidOperationException or InvalidDataException or
                                          System.Xml.XmlException or System.Text.Json.JsonException)
        {
            MessageBox.Show(this,
                $"無法讀取上次正式 Excel 的人工修改：{exception.Message}\n\n請關閉 Excel 後重試，或把正確舊檔放回原路徑。",
                "人工修改讀取失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        if (capturedState is null)
        {
            return;
        }
        manualState = capturedState;

        if (changeEntries.Count == 0)
        {
            MessageBox.Show(this, "請先同步兩個網站或載入來源 Excel，再匯出。", "尚無可匯出資料",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Title = "建立或更新中秋禮盒 Excel",
            Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx",
            DefaultExt = "xlsx",
            AddExtension = true,
            OverwritePrompt = true,
            FileName = "中秋禮盒訂單統計.xlsx"
        };
        if (!string.IsNullOrWhiteSpace(manualState.LastWorkbookPath))
        {
            dialog.InitialDirectory = Path.GetDirectoryName(manualState.LastWorkbookPath);
            dialog.FileName = Path.GetFileName(manualState.LastWorkbookPath);
        }

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            var automaticEntries = GiftBoxWorkbookExporter.ResolveEntries(changeEntries, snapshots);
            if (manualState.LastWorkbookPath is null && File.Exists(dialog.FileName))
            {
                manualState = CaptureFirstExistingWorkbook(
                    dialog.FileName,
                    automaticEntries,
                    manualState);
            }

            var effectiveEntries = ManualOverrideWorkflow.Apply(automaticEntries, manualState.Overrides);
            effectiveEntries = ManualOverrideWorkflow.ApplyAutomaticResets(
                effectiveEntries,
                manualState.PendingAutomaticResets ?? []);
            GiftBoxWorkbookExporter.Export(dialog.FileName, effectiveEntries);
            var exportedRows = GiftBoxWorkbookExporter.ReadEditableRows(dialog.FileName);
            manualState = manualState with
            {
                LastWorkbookPath = Path.GetFullPath(dialog.FileName),
                LastExportedRows = exportedRows,
                PendingAutomaticResets = []
            };
            await _manualOverrideStore.SaveAsync(manualState);
            _statusLabel.Text = $"Excel 已匯出：{dialog.FileName}";
            MessageBox.Show(this,
                "Excel 匯出完成。\n包含「訂單明細」、「日期待確認」、「統計」三張工作表。\n" +
                "所有入數依來源、訂單編號、品名與指定到貨日逐筆列出；統計數量使用 Excel 原生公式。\n" +
                "下次匯出前會讀回人工修改的指定到貨日、下單日、地點與確認，包含刻意留空的值。",
                "匯出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or
                                          InvalidOperationException or ArgumentException or System.Xml.XmlException or
                                          System.Text.Json.JsonException)
        {
            _statusLabel.Text = "Excel 匯出失敗";
            MessageBox.Show(this, $"無法匯出 Excel：{exception.Message}", "匯出失敗",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private async Task<ManualOverrideState?> CaptureLastWorkbookChangesAsync(
        ManualOverrideState state,
        bool promptWhenMissing = true)
    {
        if (string.IsNullOrWhiteSpace(state.LastWorkbookPath))
        {
            return state;
        }

        var workbookPath = state.LastWorkbookPath;
        if (!File.Exists(workbookPath))
        {
            if (!promptWhenMissing)
            {
                return state;
            }

            var choice = MessageBox.Show(
                this,
                $"找不到上次正式 Excel：\n{workbookPath}\n\n" +
                "若該檔曾有人工修改，請先把舊檔放回原路徑。\n" +
                "按「是」可改選舊檔；按「否」確定匯出並沿用已保存的人工覆寫；按「取消」停止。",
                "找不到上次正式 Excel",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button3);
            if (choice == DialogResult.Cancel) return null;
            if (choice == DialogResult.No) return state;

            using var open = new OpenFileDialog
            {
                Title = "選擇先前正式匯出的 Excel",
                Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx",
                CheckFileExists = true,
                Multiselect = false
            };
            if (open.ShowDialog(this) != DialogResult.OK) return null;
            workbookPath = open.FileName;
        }

        var workbookRows = GiftBoxWorkbookExporter.ReadEditableRows(workbookPath);
        var overrides = state.LastExportedRows.Count == 0
            ? state.Overrides
            : ManualOverrideWorkflow.CaptureChanges(
                state.LastExportedRows,
                workbookRows,
                state.Overrides,
                DateTimeOffset.Now);
        var updated = state with
        {
            LastWorkbookPath = Path.GetFullPath(workbookPath),
            Overrides = overrides
        };
        await _manualOverrideStore.SaveAsync(updated);
        return updated;
    }

    private static ManualOverrideState CaptureFirstExistingWorkbook(
        string workbookPath,
        IReadOnlyList<OrderChangeEntry> automaticEntries,
        ManualOverrideState state)
    {
        var temporaryPath = Path.Combine(
            Path.GetTempPath(), $"mid-autumn-baseline-{Guid.NewGuid():N}.xlsx");
        try
        {
            GiftBoxWorkbookExporter.Export(temporaryPath, automaticEntries);
            var baselineRows = GiftBoxWorkbookExporter.ReadEditableRows(temporaryPath);
            var workbookRows = GiftBoxWorkbookExporter.ReadEditableRows(workbookPath);
            return state with
            {
                Overrides = ManualOverrideWorkflow.CaptureChanges(
                    baselineRows,
                    workbookRows,
                    state.Overrides,
                    DateTimeOffset.Now)
            };
        }
        finally
        {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
        }
    }

    private async Task ClearManualOverrideAsync()
    {
        ManualOverrideState state;
        try
        {
            state = await _manualOverrideStore.LoadAsync();
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or
                                          System.Text.Json.JsonException)
        {
            MessageBox.Show(this, $"無法讀取人工覆寫：{exception.Message}", "讀取失敗",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (state.Overrides.Count == 0)
        {
            MessageBox.Show(this, "目前沒有已保存的人工覆寫。", "清除人工覆寫",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var selected = SelectManualOverride(state);
        if (selected is null) return;
        var confirmation = MessageBox.Show(
            this,
            $"確定清除這筆訂單的四個人工欄位嗎？\n\n{selected.DisplayText}\n\n" +
            "指定到貨日、下單日、地點、確認會在下次匯出時全部恢復為程式自動值。",
            "確認清除人工覆寫",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);
        if (confirmation != DialogResult.Yes) return;

        var baselineRows = state.LastExportedRows.ToList();
        if (!string.IsNullOrWhiteSpace(state.LastWorkbookPath) && File.Exists(state.LastWorkbookPath))
        {
            var current = GiftBoxWorkbookExporter.ReadEditableRows(state.LastWorkbookPath)
                .FirstOrDefault(row => row.RowKey.Equals(selected.Override.RowKey,
                    StringComparison.OrdinalIgnoreCase));
            if (current is not null)
            {
                baselineRows.RemoveAll(row => row.RowKey.Equals(selected.Override.RowKey,
                    StringComparison.OrdinalIgnoreCase));
                baselineRows.Add(current);
            }
        }

        var pendingResets = (state.PendingAutomaticResets ?? [])
            .Where(item => !item.RowKey.Equals(selected.Override.RowKey, StringComparison.OrdinalIgnoreCase))
            .Append(new WorkbookRowAutomaticReset(
                selected.Override.RowKey,
                selected.Override.SourceLineFingerprints))
            .ToArray();
        await _manualOverrideStore.SaveAsync(state with
        {
            LastExportedRows = baselineRows,
            Overrides = ManualOverrideWorkflow.Clear(state.Overrides, selected.Override.RowKey),
            PendingAutomaticResets = pendingResets
        });
        MessageBox.Show(this, "已清除該筆訂單的四個人工覆寫；下次匯出會恢復自動值。",
            "清除完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private ManualOverrideChoice? SelectManualOverride(ManualOverrideState state)
    {
        var rowsByKey = state.LastExportedRows
            .GroupBy(row => row.RowKey, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        var choices = state.Overrides.Select(item =>
        {
            rowsByKey.TryGetValue(item.RowKey, out var row);
            var display = row is null
                ? $"識別值 {item.RowKey[..Math.Min(12, item.RowKey.Length)]}"
                : $"{row.ChannelName}｜{row.ExternalOrderNo}｜{row.ProductName}";
            return new ManualOverrideChoice(item, display);
        }).ToArray();

        using var dialog = new Form
        {
            Text = "選擇要清除人工覆寫的訂單明細",
            StartPosition = FormStartPosition.CenterParent,
            Size = new Size(720, 420),
            MinimizeBox = false,
            MaximizeBox = false
        };
        var list = new ListBox { Dock = DockStyle.Fill, DisplayMember = nameof(ManualOverrideChoice.DisplayText) };
        list.DataSource = choices;
        var buttons = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            AutoSize = true,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(8)
        };
        var cancel = new Button { Text = "取消", AutoSize = true, DialogResult = DialogResult.Cancel };
        var ok = new Button { Text = "選擇", AutoSize = true, DialogResult = DialogResult.OK };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);
        dialog.Controls.Add(list);
        dialog.Controls.Add(buttons);
        dialog.AcceptButton = ok;
        dialog.CancelButton = cancel;
        return dialog.ShowDialog(this) == DialogResult.OK
            ? list.SelectedItem as ManualOverrideChoice
            : null;
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

    private static string? GetRowValue(
        IReadOnlyDictionary<string, string?>? row,
        string name)
    {
        if (row is null)
        {
            return null;
        }

        return row.FirstOrDefault(pair => pair.Key.Equals(name, StringComparison.OrdinalIgnoreCase)).Value;
    }

    private async Task RunBusyAsync(string busyMessage, Func<Task<string>> action)
    {
        SetBusy(true, busyMessage);
        try
        {
            var completionMessage = await action();
            _statusLabel.Text = completionMessage;
        }
        catch (Exception exception) when (exception is WmsApiException or ArgumentException or
                                          InvalidOperationException or Win32Exception or IOException or
                                          UnauthorizedAccessException or System.Text.Json.JsonException)
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
        _ordersButton.Enabled = !busy && _sourceBox.SelectedItem is WmsSource;
        _exportButton.Enabled = !busy;
        _clearOverridesButton.Enabled = !busy;
        _loadErpFileButton.Enabled = !busy;
        _loadManualExcelButton.Enabled = !busy;
        _loadPosFileButton.Enabled = !busy;
        _progress.Visible = busy;
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        _statusLabel.Text = message;
    }

    private enum SpreadsheetImportKind
    {
        Erp,
        Manual,
        Pos
    }

    private sealed record CompleteSourceReload(
        WmsSource Source,
        SyncWindow Window,
        DateTimeOffset StartedAt,
        DateTimeOffset FinishedAt,
        int PageCount,
        int RowCount,
        IReadOnlyList<OrderLineSnapshot> Snapshots);

    private sealed record ManualOverrideChoice(
        OrderRowManualOverride Override,
        string DisplayText);

    private void UpdateOrderButtonState()
    {
        _ordersButton.Enabled = _sourceBox.SelectedItem is WmsSource;
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
