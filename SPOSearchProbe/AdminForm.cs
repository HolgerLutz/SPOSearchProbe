using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Web;

namespace SPOSearchProbe;

// Extended config model for admin JSON (users)
public class AdminConfig
{
    [JsonPropertyName("siteUrl")] public string SiteUrl { get; set; } = "";
    [JsonPropertyName("tenantId")] public string TenantId { get; set; } = "";
    [JsonPropertyName("queryText")] public string QueryText { get; set; } = "contentclass:STS_ListItem";
    [JsonPropertyName("selectProperties")] public string[] SelectProperties { get; set; } = ["Title", "Path", "LastModifiedTime", "WorkId"];
    [JsonPropertyName("rowLimit")] public int RowLimit { get; set; } = 10;
    [JsonPropertyName("sortList")] public string SortList { get; set; } = "";
    [JsonPropertyName("pageUrl")] public string PageUrl { get; set; } = "";
    [JsonPropertyName("intervalValue")] public int IntervalValue { get; set; } = 10;
    [JsonPropertyName("intervalUnit")] public string IntervalUnit { get; set; } = "seconds";
    [JsonPropertyName("clientId")] public string ClientId { get; set; } = "";
    [JsonPropertyName("workspaceUrl")] public string WorkspaceUrl { get; set; } = "";
    [JsonPropertyName("users")] public List<UserEntry> Users { get; set; } = [];
}

public class UserEntry
{
    [JsonPropertyName("name")] public string Name { get; set; } = "";
    [JsonPropertyName("enabled")] public bool Enabled { get; set; } = true;
    [JsonPropertyName("affected")] public bool Affected { get; set; }
    [JsonPropertyName("tokenCacheFile")] public string TokenCacheFile { get; set; } = "";
}

public class AdminForm : Form
{
    private readonly string _appDir;
    private readonly SearchClient _searchClient = new();
    private string _configPath;

    // Search Configuration controls
    private readonly TextBox _txtSite, _txtQuery, _txtProps, _txtSortList, _txtPageUrl;
    private readonly NumericUpDown _numRows;
    private readonly Button _btnLoadConfig, _btnSaveConfig, _btnCopyUrl;

    // Users controls
    private readonly DataGridView _dgvUsers;
    private readonly Button _btnAddUser, _btnRename, _btnRemove, _btnLogin;

    // Package controls
    private readonly Button _btnCreatePackage;

    // Execution controls
    private readonly ComboBox _cmbInterval, _cmbUnit;
    private readonly TextBox _txtTenant, _txtClientId;
    private readonly Label _lblClientId;
    private readonly Button _btnStart, _btnStop, _btnTestQuery, _btnValidatePage, _btnResetValidation;
    private readonly Label _lblWorkId, _lblStatus, _lblNext;

    // Log controls
    private readonly RichTextBox _txtLog;
    private readonly Button _btnCollectLogs;

    // State
    private List<UserEntry> _users = [];
    private bool _isRunning;
    private bool _runOnce;
    private DateTime? _nextExecution;
    private string _validatedWorkId = "";
    private bool _validationActive;

    // Tooltip provider for all controls
    private readonly ToolTip _tip = new() { AutoPopDelay = 15000, InitialDelay = 400, ReshowDelay = 200 };
    private readonly List<string> _validationFoundUsers = [];
    private string _workspaceUrl = "";

    // Live chart
    private LiveChartForm? _liveChart;
    private readonly List<ChartDataPoint> _chartDataPoints = [];

    // Logging
    private string _logFile = "";
    private string _tsvFile = "";
    private string _requestLogDir = "";

    // Timers
    private readonly System.Windows.Forms.Timer _searchTimer = new();
    private readonly System.Windows.Forms.Timer _countdownTimer = new();

    public AdminForm(SearchConfig config, string configPath, string appDir)
    {
        _appDir = appDir;
        _configPath = configPath;

        SuspendLayout();
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 9);
        Text = $"SPO Search Probe Tool (Admin) | {Program.AppTitleSuffix}";
        StartPosition = FormStartPosition.CenterScreen;
        AutoScroll = true;
        Size = new Size(920, 700);
        MinimumSize = new Size(640, 400);

        SetIcon();
        InitLogging();

        // ========== 1. Search Configuration Group (top, 880x170) ==========
        var grpSearch = new GroupBox
        {
            Text = "Search Configuration",
            Location = new Point(12, 8),
            Size = new Size(880, 170),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        var lblSite = new Label { Text = "Site URL:", Location = new Point(10, 22), AutoSize = true };
        _txtSite = new TextBox
        {
            Location = new Point(100, 19), Size = new Size(550, 23),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "https://yourtenant.sharepoint.com"
        };

        _btnLoadConfig = new Button
        {
            Text = "Load Config...", Location = new Point(670, 18), Size = new Size(100, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _btnLoadConfig.Click += BtnLoadConfig_Click;

        _btnSaveConfig = new Button
        {
            Text = "Save Config...", Location = new Point(775, 18), Size = new Size(95, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _btnSaveConfig.Click += BtnSaveConfig_Click;

        var lblQuery = new Label { Text = "Query:", Location = new Point(10, 52), AutoSize = true };
        _txtQuery = new TextBox
        {
            Location = new Point(100, 49), Size = new Size(770, 23),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "contentclass:STS_ListItem"
        };

        var lblProps = new Label { Text = "Properties:", Location = new Point(10, 82), AutoSize = true };
        _txtProps = new TextBox
        {
            Location = new Point(100, 79), Size = new Size(580, 23),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Text = "Title,Path,LastModifiedTime,Author,SiteId,WebId,ListId,UniqueId,WorkId,RefinableString01"
        };

        var lblRows = new Label
        {
            Text = "Row Limit:", Location = new Point(700, 82), AutoSize = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _numRows = new NumericUpDown
        {
            Location = new Point(775, 79), Size = new Size(95, 23),
            Minimum = 1, Maximum = 500, Value = 50,
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };

        var lblSortList = new Label { Text = "Sort List:", Location = new Point(10, 112), AutoSize = true };
        _txtSortList = new TextBox
        {
            Location = new Point(100, 109), Size = new Size(770, 23),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        var lblPageUrl = new Label { Text = "Page URL:", Location = new Point(10, 142), AutoSize = true };
        _txtPageUrl = new TextBox
        {
            Location = new Point(100, 139), Size = new Size(670, 23),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        _txtPageUrl.TextChanged += (_, _) =>
        {
            // Page URL changed — invalidate any previous validation
            if (_validationActive)
            {
                _validatedWorkId = "";
                _validationActive = false;
                _validationFoundUsers.Clear();
                _btnResetValidation!.Enabled = false;
                _btnStart!.Enabled = false;
                _lblWorkId!.Text = "WorkId: (none)";
                _lblWorkId!.ForeColor = Color.Gray;
            }
            _btnValidatePage!.Enabled = true;
        };

        _btnCopyUrl = new Button
        {
            Text = "Copy URL", Location = new Point(775, 138), Size = new Size(95, 25),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        _btnCopyUrl.Click += BtnCopyUrl_Click;

        grpSearch.Controls.AddRange([lblSite, _txtSite, _btnLoadConfig, _btnSaveConfig,
            lblQuery, _txtQuery, lblProps, _txtProps, lblRows, _numRows,
            lblSortList, _txtSortList, lblPageUrl, _txtPageUrl, _btnCopyUrl]);

        // ========== 2. Users Group (left side, 540x160) ==========
        var grpUsers = new GroupBox
        {
            Text = "Users",
            Location = new Point(12, 184),
            Size = new Size(540, 160),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        _dgvUsers = new DataGridView
        {
            Location = new Point(10, 22), Size = new Size(520, 95),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            AllowUserToAddRows = false, AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false, RowHeadersVisible = false,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = SystemColors.Window, ReadOnly = false,
            ScrollBars = ScrollBars.Vertical
        };

        var colEnabled = new DataGridViewCheckBoxColumn
        { Name = "Enabled", HeaderText = "", Width = 30, FillWeight = 6 };
        var colAffected = new DataGridViewCheckBoxColumn
        { Name = "Affected", HeaderText = "Affected", Width = 55, FillWeight = 10 };
        var colName = new DataGridViewTextBoxColumn
        { Name = "Name", HeaderText = "User", ReadOnly = true, FillWeight = 40 };
        var colStatus = new DataGridViewTextBoxColumn
        { Name = "Status", HeaderText = "Token Status", ReadOnly = true, FillWeight = 44 };
        _dgvUsers.Columns.AddRange([colEnabled, colAffected, colName, colStatus]);

        _dgvUsers.CellValueChanged += (_, e) =>
        {
            if (e.RowIndex >= 0 && (e.ColumnIndex == 0 || e.ColumnIndex == 1))
                SyncEnabledState();
        };
        _dgvUsers.CurrentCellDirtyStateChanged += (_, _) =>
        {
            if (_dgvUsers.IsCurrentCellDirty) _dgvUsers.CommitEdit(DataGridViewDataErrorContexts.Commit);
        };

        int btnY = 122, btnH = 28;
        _btnAddUser = new Button { Text = "Add User", Location = new Point(10, btnY), Size = new Size(72, btnH) };
        _btnAddUser.Click += BtnAddUser_Click;
        _btnLogin = new Button
        {
            Text = "Login", Location = new Point(87, btnY), Size = new Size(55, btnH),
            Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        _btnLogin.Click += BtnLoginUser_Click;
        _btnRename = new Button { Text = "Rename", Location = new Point(147, btnY), Size = new Size(62, btnH) };
        _btnRename.Click += BtnRename_Click;
        _btnRemove = new Button { Text = "Remove", Location = new Point(214, btnY), Size = new Size(62, btnH) };
        _btnRemove.Click += BtnRemove_Click;

        grpUsers.Controls.AddRange([_dgvUsers, _btnAddUser, _btnLogin, _btnRename, _btnRemove]);

        // ========== 3. Package Group (540x58) ==========
        var grpKeys = new GroupBox
        {
            Text = "Package",
            Location = new Point(12, 348),
            Size = new Size(540, 58),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        int kBtnY = 20;
        _btnCreatePackage = new Button
        {
            Text = "📦 Create EndUser Package", Location = new Point(10, kBtnY), Size = new Size(190, btnH),
            ForeColor = Color.DarkGreen, Font = new Font("Segoe UI", 9, FontStyle.Bold)
        };
        _btnCreatePackage.Click += BtnCreatePackage_Click;

        grpKeys.Controls.AddRange([_btnCreatePackage]);

        // ========== 4. Execution Group (right side, 332x225) ==========
        var grpExec = new GroupBox
        {
            Text = "Execution",
            Location = new Point(560, 184),
            Size = new Size(332, 225),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };

        var lblInterval = new Label { Text = "Interval:", Location = new Point(10, 28), AutoSize = true };
        _cmbInterval = new ComboBox
        {
            Location = new Point(75, 25), Size = new Size(55, 23),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        foreach (var v in new[] { "5", "10", "15", "30", "60", "90", "120" })
            _cmbInterval.Items.Add(v);
        _cmbInterval.SelectedItem = "5";

        _cmbUnit = new ComboBox
        {
            Location = new Point(133, 25), Size = new Size(75, 23),
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cmbUnit.Items.Add("minutes");
        _cmbUnit.Items.Add("seconds");
        _cmbUnit.SelectedItem = "minutes";

        var lblTenant = new Label { Text = "Tenant:", Location = new Point(10, 60), AutoSize = true };
        _txtTenant = new TextBox
        {
            Location = new Point(75, 57), Size = new Size(245, 23),
            Text = "yourtenant.onmicrosoft.com"
        };

        _lblClientId = new Label
        {
            Text = "Client ID:", Location = new Point(10, 128), AutoSize = true, Visible = false
        };
        _txtClientId = new TextBox
        {
            Location = new Point(75, 125), Size = new Size(245, 23),
            Text = "", Visible = false
        };

        _btnStart = new Button
        {
            Text = "Start", Location = new Point(10, 90), Size = new Size(80, 32),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(220, 255, 220),
            Enabled = false
        };
        _btnStart.Click += BtnStart_Click;

        _btnStop = new Button
        {
            Text = "Stop", Location = new Point(95, 90), Size = new Size(80, 32),
            Enabled = false, BackColor = Color.FromArgb(255, 220, 220)
        };
        _btnStop.Click += BtnStop_Click;

        _btnTestQuery = new Button
        {
            Text = "Test Query", Location = new Point(195, 90), Size = new Size(125, 32)
        };
        _btnTestQuery.Click += (_, _) => { WriteLog("--- Running test query ---", Color.Cyan); _runOnce = true; _ = ExecuteSearchForAllUsersAsync(); };

        _btnValidatePage = new Button
        {
            Text = "Validate Page", Location = new Point(10, 128), Size = new Size(150, 32),
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            BackColor = Color.FromArgb(220, 230, 255)
        };
        _btnValidatePage.Click += BtnValidatePage_Click;

        _btnResetValidation = new Button
        {
            Text = "Reset", Location = new Point(165, 128), Size = new Size(60, 32),
            Enabled = false
        };
        _btnResetValidation.Click += (_, _) =>
        {
            _validatedWorkId = "";
            _validationActive = false;
            _validationFoundUsers.Clear();
            _btnResetValidation.Enabled = false;
            _btnStart.Enabled = false;
            _btnValidatePage.Enabled = true;
            _lblWorkId!.Text = "WorkId: (none)";
            _lblWorkId!.ForeColor = Color.Gray;
            WriteLog("Validation reset.", Color.Yellow);
        };

        _lblWorkId = new Label
        {
            Text = "WorkId: (none)", Location = new Point(10, 165),
            Size = new Size(310, 18), ForeColor = Color.Gray,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        _lblStatus = new Label
        {
            Text = "Status: Idle", Location = new Point(10, 183),
            Size = new Size(310, 18),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        _lblNext = new Label
        {
            Text = "Next execution: --:--:--", Location = new Point(10, 201),
            Size = new Size(310, 18),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        grpExec.Controls.AddRange([lblInterval, _cmbInterval, _cmbUnit,
            lblTenant, _txtTenant, _lblClientId, _txtClientId,
            _btnStart, _btnStop, _btnTestQuery,
            _btnValidatePage, _btnResetValidation,
            _lblWorkId, _lblStatus, _lblNext]);

        // ========== 5. Log Group (bottom, fills remaining) ==========
        var grpLog = new GroupBox
        {
            Text = "Log",
            Location = new Point(12, 410),
            Size = new Size(880, 250),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        };

        _btnCollectLogs = new Button
        {
            Text = "Collect Logs", Location = new Point(grpLog.Width - 110, 0),
            Size = new Size(100, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Font = new Font("Segoe UI", 8)
        };
        _btnCollectLogs.Click += BtnCollectLogs_Click;

        var btnShowChart = new Button
        {
            Text = "📈 Chart", Location = new Point(grpLog.Width - 200, 0),
            Size = new Size(80, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Font = new Font("Segoe UI", 8)
        };
        btnShowChart.Click += (_, _) => ShowLiveChart();

        _txtLog = new RichTextBox
        {
            Location = new Point(10, 20),
            Size = new Size(860, 222),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            ReadOnly = true,
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.FromArgb(200, 200, 200),
            Font = new Font("Consolas", 9),
            WordWrap = false,
            ScrollBars = RichTextBoxScrollBars.Both
        };

        grpLog.Controls.AddRange([_txtLog, btnShowChart, _btnCollectLogs]);

        Controls.AddRange([grpSearch, grpUsers, grpKeys, grpExec, grpLog]);
        ResumeLayout(true);

        // --- Tooltips for all interactive controls ---
        // Search Configuration
        _tip.SetToolTip(_txtSite, "SharePoint Online site collection URL to query (e.g. https://contoso.sharepoint.com/sites/hr)");
        _tip.SetToolTip(_btnLoadConfig, "Load a search-config.json file from disk");
        _tip.SetToolTip(_btnSaveConfig, "Save the current search configuration to a JSON file");
        _tip.SetToolTip(_txtQuery, "KQL (Keyword Query Language) search query to execute against SharePoint");
        _tip.SetToolTip(_txtProps, "Comma-separated managed properties to include in search results (e.g. Title,Path,WorkId)");
        _tip.SetToolTip(_numRows, "Maximum number of search result rows returned per query (1-500)");
        _tip.SetToolTip(_txtSortList, "Optional sort order for results (e.g. LastModifiedTime:descending)");
        _tip.SetToolTip(_txtPageUrl, "Full URL of a specific page to monitor — used for page validation tracking");
        _tip.SetToolTip(_btnCopyUrl, "Copy the Page URL to the clipboard");
        // Users
        _tip.SetToolTip(_dgvUsers, "List of user accounts for search probing. Check 'Enabled' to include in probe runs, 'Affected' to flag as impacted.");
        _tip.SetToolTip(_btnAddUser, "Add a new user account (enter email address)");
        _tip.SetToolTip(_btnLogin, "Authenticate the selected user via OAuth2 device code flow");
        _tip.SetToolTip(_btnRename, "Rename the selected user entry");
        _tip.SetToolTip(_btnRemove, "Remove the selected user from the list");
        // Package
        _tip.SetToolTip(_btnCreatePackage, "Create a ZIP package containing the exe + config for distribution to end users");
        // Execution
        _tip.SetToolTip(_cmbInterval, "How often the search probe executes (number)");
        _tip.SetToolTip(_cmbUnit, "Time unit for the probe interval (seconds or minutes)");
        _tip.SetToolTip(_txtTenant, "Azure AD tenant name or ID for OAuth2 authentication (e.g. contoso.onmicrosoft.com)");
        _tip.SetToolTip(_txtClientId, "Azure AD application (client) ID for OAuth2 — leave empty to use the default PnP client");
        _tip.SetToolTip(_btnStart, "Start the recurring search probe scheduler (requires page validation first)");
        _tip.SetToolTip(_btnStop, "Stop the running search probe scheduler");
        _tip.SetToolTip(_btnTestQuery, "Run a single test query immediately without starting the scheduler");
        _tip.SetToolTip(_btnValidatePage, "Validate the Page URL by searching for it and capturing its WorkId for tracking");
        _tip.SetToolTip(_btnResetValidation, "Clear the validated WorkId and stop page validation tracking");
        // Log
        _tip.SetToolTip(btnShowChart, "Open the live execution timeline chart in a separate window");
        _tip.SetToolTip(_btnCollectLogs, "Package all session logs, TSV data, and request traces into a ZIP with an HTML report");

        // Cap form to screen after DPI auto-scaling
        Load += (_, _) =>
        {
            var wa = Screen.FromControl(this).WorkingArea;
            if (Width > wa.Width || Height > wa.Height)
                Size = new Size(Math.Min(Width, wa.Width), Math.Min(Height, wa.Height));
            CenterToScreen();
        };

        // Timers
        _countdownTimer.Interval = 1000;
        _countdownTimer.Tick += (_, _) =>
        {
            if (_nextExecution.HasValue)
            {
                var remaining = _nextExecution.Value - DateTime.Now;
                _lblNext.Text = remaining.TotalSeconds > 0
                    ? $"Next execution: {remaining:mm\\:ss}"
                    : "Next execution: Executing...";
            }
        };
        _searchTimer.Tick += async (_, _) => await ExecuteSearchForAllUsersAsync();

        // Load initial config
        LoadConfigFromFile(_configPath, silent: true);
        RefreshUserGrid();

        WriteLog("SPO Search Probe Tool (Admin) - .NET started.", Color.LimeGreen);
    }

    // =====================================================================
    // Icon (same magnifying glass as MainForm)
    // =====================================================================
    private void SetIcon()
    {
        try
        {
            var bmp = new Bitmap(32, 32);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            g.DrawEllipse(new Pen(Color.FromArgb(40, 0, 120, 212), 1), 1, 1, 22, 22);
            g.DrawEllipse(new Pen(Color.FromArgb(80, 0, 120, 212), 1), 3, 3, 18, 18);
            g.DrawEllipse(new Pen(Color.FromArgb(0, 120, 212), 2.5f), 4, 4, 16, 16);
            g.FillEllipse(new SolidBrush(Color.FromArgb(30, 0, 120, 212)), 4, 4, 16, 16);
            using var pen = new Pen(Color.FromArgb(0, 90, 170), 3)
            {
                StartCap = System.Drawing.Drawing2D.LineCap.Round,
                EndCap = System.Drawing.Drawing2D.LineCap.Round
            };
            g.DrawLine(pen, 18, 18, 28, 28);
            g.FillEllipse(new SolidBrush(Color.FromArgb(0, 180, 80)), 10, 10, 4, 4);
            Icon = Icon.FromHandle(bmp.GetHicon());
        }
        catch { }
    }

    // =====================================================================
    // Logging
    // =====================================================================
    private void InitLogging()
    {
        var logDir = Path.Combine(_appDir, "logs");
        Directory.CreateDirectory(logDir);
        var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        _logFile = Path.Combine(logDir, $"{ts}_SPOSearchProbe.log");
        _tsvFile = Path.Combine(logDir, $"{ts}_SPOSearchProbe.tsv");
        _requestLogDir = Path.Combine(logDir, "requests", ts);

        var tsvHeader = "DateTime\tUser\tHttpStatus\tQueryType\tValidationStatus\tCorrelationId\tX-SearchInternalRequestId\t" +
                        "ResultCount\tExecutionTime\tTitle\tPath\tLastModifiedTime\tSiteId\tWebId\tListId\t" +
                        "UniqueId\tWorkId\tRefinableString01\tQueryIdentityDiagnostics";
        try { File.WriteAllText(_tsvFile, tsvHeader + "\r\n"); } catch { }
    }

    private void WriteLog(string message, Color color)
    {
        if (InvokeRequired) { Invoke(() => WriteLog(message, color)); return; }
        var ts = DateTime.Now.ToString("HH:mm:ss");
        var line = $"[{ts}] {message}";
        _txtLog.SelectionStart = _txtLog.TextLength;
        _txtLog.SelectionColor = color;
        _txtLog.AppendText(line + "\r\n");
        _txtLog.ScrollToCaret();
        try { File.AppendAllText(_logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\r\n"); } catch { }
    }

    private void WriteTsvRow(string user, string httpStatus, string correlationId, string requestId,
        int resultCount, string executionTime, Dictionary<string, string>? cells = null,
        string validationStatus = "", string queryIdDiag = "", string queryType = "")
    {
        var dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        string Get(string key) => cells != null && cells.TryGetValue(key, out var v)
            ? v.Replace("\t", " ").Replace("\n", " ") : "";
        var vals = new[] { dt, user, httpStatus, queryType, validationStatus, correlationId, requestId,
            resultCount.ToString(), executionTime, Get("Title"), Get("Path"), Get("LastModifiedTime"),
            Get("SiteId"), Get("WebId"), Get("ListId"), Get("UniqueId"), Get("WorkId"),
            Get("RefinableString01"), queryIdDiag.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ") };
        try { File.AppendAllText(_tsvFile, string.Join("\t", vals) + "\r\n"); } catch { }
    }

    // =====================================================================
    // Helpers
    // =====================================================================
    private string GetEffectiveTenantId()
    {
        var tenant = _txtTenant.Text.Trim();
        if (string.IsNullOrEmpty(tenant) || tenant == "yourtenant.onmicrosoft.com")
            return "organizations";

        if (tenant.Contains(".sharepoint.com", StringComparison.OrdinalIgnoreCase))
        {
            tenant = Regex.Replace(tenant, @"-admin\.sharepoint\.com$", ".onmicrosoft.com", RegexOptions.IgnoreCase);
            tenant = Regex.Replace(tenant, @"-df\.sharepoint\.com$", ".onmicrosoft.com", RegexOptions.IgnoreCase);
            tenant = Regex.Replace(tenant, @"\.sharepoint\.com$", ".onmicrosoft.com", RegexOptions.IgnoreCase);
            _txtTenant.Text = tenant;
        }
        else if (!tenant.Contains('.') &&
                 !Regex.IsMatch(tenant, @"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$"))
        {
            tenant = $"{tenant}.onmicrosoft.com";
            _txtTenant.Text = tenant;
        }
        return tenant;
    }

    private string GetSearchScope()
    {
        var spUri = new Uri(_txtSite.Text.Trim());
        return $"{spUri.Scheme}://{spUri.Authority}/.default offline_access openid";
    }

    private int GetIntervalMs()
    {
        var val = int.Parse((string)_cmbInterval.SelectedItem!);
        return (string)_cmbUnit.SelectedItem! == "seconds" ? val * 1000 : val * 60_000;
    }

    private string GetIntervalLabel()
    {
        var val = (string)_cmbInterval.SelectedItem!;
        var unit = (string)_cmbUnit.SelectedItem!;
        return unit == "seconds" ? $"{val}s" : $"{val} min";
    }

    private static string GetUserCacheFileName(string userName)
    {
        var safe = Regex.Replace(userName, "[^a-zA-Z0-9]", "_");
        return $".token-{safe}.dat";
    }

    private string GetUserTokenStatus(string cacheFile)
    {
        var fullPath = Path.Combine(_appDir, cacheFile);
        var cache = TokenCache.Load(fullPath);
        if (cache == null) return "No Token";
        if (string.IsNullOrEmpty(cache.RefreshToken)) return "No Refresh Token";
        var expires = DateTime.Parse(cache.ExpiresOn);
        return expires > DateTime.Now
            ? $"Active (expires {expires:HH:mm})"
            : "Expired (has refresh token)";
    }

    private int GetSelectedUserIndex()
    {
        if (_dgvUsers.SelectedRows.Count == 0) return -1;
        return _dgvUsers.SelectedRows[0].Index;
    }

    // =====================================================================
    // User Grid
    // =====================================================================
    private void RefreshUserGrid()
    {
        _dgvUsers.Rows.Clear();
        foreach (var u in _users)
        {
            var status = GetUserTokenStatus(u.TokenCacheFile);
            var idx = _dgvUsers.Rows.Add(u.Enabled, u.Affected, u.Name, status);
            if (u.Affected)
            {
                _dgvUsers.Rows[idx].DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235);
                _dgvUsers.Rows[idx].DefaultCellStyle.ForeColor = Color.DarkRed;
            }
        }
    }

    private void SyncEnabledState()
    {
        for (int i = 0; i < _dgvUsers.Rows.Count && i < _users.Count; i++)
        {
            _users[i].Enabled = _dgvUsers.Rows[i].Cells["Enabled"].Value is true;
            var newAffected = _dgvUsers.Rows[i].Cells["Affected"].Value is true;
            _users[i].Affected = newAffected;

            if (newAffected)
            {
                _dgvUsers.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235);
                _dgvUsers.Rows[i].DefaultCellStyle.ForeColor = Color.DarkRed;
            }
            else
            {
                _dgvUsers.Rows[i].DefaultCellStyle.BackColor = Color.Empty;
                _dgvUsers.Rows[i].DefaultCellStyle.ForeColor = Color.Empty;
            }
        }
        SaveUsersToConfig();
    }

    // =====================================================================
    // Config Load / Save
    // =====================================================================
    private void LoadConfigFromFile(string path, bool silent = false)
    {
        if (!File.Exists(path))
        {
            if (!silent)
                WriteLog($"Config file not found: {path}", Color.Yellow);
            return;
        }
        try
        {
            var json = File.ReadAllText(path);
            var cfg = JsonSerializer.Deserialize<AdminConfig>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            }) ?? new AdminConfig();

            _txtSite.Text = cfg.SiteUrl;
            _txtQuery.Text = cfg.QueryText;
            if (cfg.SelectProperties.Length > 0)
                _txtProps.Text = string.Join(",", cfg.SelectProperties);
            _numRows.Value = Math.Clamp(cfg.RowLimit, 1, 500);
            _txtSortList.Text = cfg.SortList;
            _txtPageUrl.Text = cfg.PageUrl;
            _txtClientId.Text = cfg.ClientId;
            if (!string.IsNullOrEmpty(cfg.TenantId))
                _txtTenant.Text = cfg.TenantId;

            if (cfg.IntervalValue > 0)
            {
                var ival = cfg.IntervalValue.ToString();
                var idx = _cmbInterval.Items.IndexOf(ival);
                if (idx >= 0) _cmbInterval.SelectedIndex = idx;
            }
            if (!string.IsNullOrEmpty(cfg.IntervalUnit))
            {
                var uidx = _cmbUnit.Items.IndexOf(cfg.IntervalUnit);
                if (uidx >= 0) _cmbUnit.SelectedIndex = uidx;
            }

            _users = cfg.Users ?? [];
            _workspaceUrl = cfg.WorkspaceUrl ?? "";

            if (!silent)
                WriteLog($"Config loaded: {path}", Color.Cyan);

            // Reset validation when loading config with page URL
            if (!string.IsNullOrWhiteSpace(cfg.PageUrl))
            {
                _validatedWorkId = "";
                _validationActive = false;
                _validationFoundUsers.Clear();
                _btnResetValidation.Enabled = false;
                _lblWorkId.Text = "WorkId: (none)";
                _lblWorkId.ForeColor = Color.Gray;
                if (!silent)
                    WriteLog("Page URL configured - validate page before starting.", Color.Yellow);
            }
        }
        catch (Exception ex)
        {
            if (!silent)
                WriteLog($"Failed to load config: {ex.Message}", Color.Red);
        }
    }

    private void SaveConfigToFile(string path)
    {
        try
        {
            // Read existing file to preserve unknown fields
            var hash = new Dictionary<string, object?>();
            if (File.Exists(path))
            {
                try
                {
                    var existing = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(File.ReadAllText(path));
                    if (existing != null)
                    {
                        foreach (var kv in existing)
                            hash[kv.Key] = kv.Value;
                    }
                }
                catch { }
            }

            hash["siteUrl"] = _txtSite.Text;
            hash["tenantId"] = _txtTenant.Text;
            hash["queryText"] = _txtQuery.Text;
            hash["selectProperties"] = _txtProps.Text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            hash["rowLimit"] = (int)_numRows.Value;
            hash["sortList"] = _txtSortList.Text;
            hash["pageUrl"] = _txtPageUrl.Text;
            hash["intervalValue"] = int.Parse((string)_cmbInterval.SelectedItem!);
            hash["intervalUnit"] = (string)_cmbUnit.SelectedItem!;

            var currentClientId = _txtClientId.Text.Trim();
            if (!string.IsNullOrEmpty(currentClientId))
                hash["clientId"] = currentClientId;

            if (!string.IsNullOrEmpty(_workspaceUrl))
                hash["workspaceUrl"] = _workspaceUrl;

            hash["users"] = _users;

            var json = JsonSerializer.Serialize(hash, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
            WriteLog($"Config saved: {path}", Color.Green);
        }
        catch (Exception ex)
        {
            WriteLog($"Failed to save config: {ex.Message}", Color.Red);
        }
    }

    private void SaveUsersToConfig()
    {
        try
        {
            var hash = new Dictionary<string, object?>();
            if (File.Exists(_configPath))
            {
                try
                {
                    var existing = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(File.ReadAllText(_configPath));
                    if (existing != null)
                        foreach (var kv in existing)
                            hash[kv.Key] = kv.Value;
                }
                catch { }
            }
            hash["users"] = _users;
            var json = JsonSerializer.Serialize(hash, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_configPath, json);
        }
        catch { }
    }

    // =====================================================================
    // Event Handlers - Package
    // =====================================================================

    private void BtnCreatePackage_Click(object? sender, EventArgs e)
    {
        try
        {
            var exePath = Environment.ProcessPath ?? Application.ExecutablePath;
            var exeDir = Path.GetDirectoryName(exePath)!;
            var configPath = Path.Combine(exeDir, "search-config.json");

            if (!File.Exists(configPath))
            {
                MessageBox.Show("search-config.json not found next to the executable.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Prompt admin for an optional workspace URL to embed in the package config
            var wsUrl = PromptForInput(
                "Workspace URL (optional)",
                "Paste a workspace URL for log uploads, or leave empty to skip:");

            // User cancelled the dialog entirely
            if (wsUrl == null) return;

            using var dlg = new SaveFileDialog
            {
                Title = "Save EndUser Distribution Package",
                Filter = "Zip files (*.zip)|*.zip",
                FileName = $"SPOSearchProbe-{Program.AppVersion}.zip",
                InitialDirectory = exeDir
            };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            var zipPath = dlg.FileName;
            if (File.Exists(zipPath)) File.Delete(zipPath);

            var staging = Path.Combine(Path.GetTempPath(), $"SPOSearchProbe_pkg_{Guid.NewGuid():N}");
            Directory.CreateDirectory(staging);

            File.Copy(exePath, Path.Combine(staging, Path.GetFileName(exePath)), true);

            // Copy config, strip admin-only fields (users), and optionally patch workspace URL
            var stagedConfig = Path.Combine(staging, "search-config.json");
            File.Copy(configPath, stagedConfig, true);

            {
                var json = File.ReadAllText(stagedConfig);
                var doc = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json);
                if (doc != null)
                {
                    // Remove admin-only fields that end users should not see
                    doc.Remove("users");

                    // Patch workspace URL if provided
                    if (!string.IsNullOrWhiteSpace(wsUrl))
                        doc["workspaceUrl"] = JsonSerializer.Deserialize<JsonElement>($"\"{wsUrl.Trim().Replace("\"", "\\\"")}\"");

                    var opts = new JsonSerializerOptions { WriteIndented = true };
                    File.WriteAllText(stagedConfig, JsonSerializer.Serialize(doc, opts));
                }
            }

            System.IO.Compression.ZipFile.CreateFromDirectory(staging, zipPath);
            Directory.Delete(staging, true);

            Clipboard.SetText(zipPath);
            WriteLog($"[Package] Created: {zipPath} (path copied to clipboard)", Color.Green);
            MessageBox.Show($"EndUser package created:\n{zipPath}\n\nPath copied to clipboard.",
                "Create EndUser Package", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            WriteLog($"[Package] Failed: {ex.Message}", Color.Red);
            MessageBox.Show($"Failed to create package:\n{ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// Shows a simple input dialog with a prompt label and a text box.
    /// Returns the entered text (may be empty), or null if the user cancelled.
    /// </summary>
    private static string? PromptForInput(string title, string prompt)
    {
        using var form = new Form
        {
            Text = title, Width = 500, Height = 170,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false
        };
        var lbl = new Label { Text = prompt, Left = 12, Top = 12, Width = 460, AutoSize = true };
        var txt = new TextBox { Left = 12, Top = 40, Width = 460 };
        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 310, Top = 80, Width = 75 };
        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 395, Top = 80, Width = 75 };
        form.Controls.AddRange([lbl, txt, ok, cancel]);
        form.AcceptButton = ok;
        form.CancelButton = cancel;

        return form.ShowDialog() == DialogResult.OK ? txt.Text : null;
    }

    // =====================================================================
    // Event Handlers - Search Configuration
    // =====================================================================
    private void BtnLoadConfig_Click(object? sender, EventArgs e)
    {
        WriteLog("[Action] Load Config clicked", Color.Gray);
        using var dlg = new OpenFileDialog
        {
            Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
            InitialDirectory = _appDir
        };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            _configPath = dlg.FileName;
            LoadConfigFromFile(dlg.FileName);
            RefreshUserGrid();
        }
    }

    private void BtnSaveConfig_Click(object? sender, EventArgs e)
    {
        WriteLog("[Action] Save Config clicked", Color.Gray);
        using var dlg = new SaveFileDialog
        {
            Filter = "JSON files (*.json)|*.json",
            InitialDirectory = _appDir,
            FileName = "search-config.json"
        };
        if (dlg.ShowDialog() == DialogResult.OK)
        {
            SaveConfigToFile(dlg.FileName);
        }
    }

    private void BtnCopyUrl_Click(object? sender, EventArgs e)
    {
        var siteUrl = _txtSite.Text.Trim();
        if (string.IsNullOrEmpty(siteUrl) || siteUrl == "https://yourtenant.sharepoint.com")
        {
            MessageBox.Show("Please enter a valid Site URL first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var queryText = _txtQuery.Text.Trim();
        var pageUrl = _txtPageUrl.Text.Trim();
        if (!string.IsNullOrEmpty(pageUrl))
        {
            var pathFilter = $"path:\"{pageUrl}\"";
            queryText = string.IsNullOrEmpty(queryText) ? pathFilter : $"{queryText} {pathFilter}";
        }

        var props = _txtProps.Text.Trim();
        var rowLimit = (int)_numRows.Value;
        var searchUrl = $"{siteUrl.TrimEnd('/')}/_api/search/query";
        var escaped = queryText.Replace("'", "''");
        var q = HttpUtility.UrlEncode(escaped);
        var url = $"{searchUrl}?querytext='{q}'&selectproperties='{props}'&rowlimit={rowLimit}" +
                  "&trimduplicates=false&Properties='QueryIdentityDiagnostics:true'";

        var sortList = _txtSortList.Text.Trim();
        if (!string.IsNullOrEmpty(sortList))
            url += $"&sortlist='{sortList}'";

        Clipboard.SetText(url);
        WriteLog("Search REST API URL copied to clipboard.", Color.Cyan);
    }

    // =====================================================================
    // Event Handlers - Users
    // =====================================================================
    private void BtnAddUser_Click(object? sender, EventArgs e)
    {
        var name = ShowInputDialog("Add User", "Enter user name or email:");
        if (string.IsNullOrWhiteSpace(name)) return;

        var user = new UserEntry
        {
            Name = name.Trim(),
            Enabled = true,
            Affected = false,
            TokenCacheFile = GetUserCacheFileName(name.Trim())
        };
        _users.Add(user);
        SaveUsersToConfig();
        RefreshUserGrid();
        WriteLog($"User added: {user.Name}", Color.Green);
    }

    private void BtnRename_Click(object? sender, EventArgs e)
    {
        var idx = GetSelectedUserIndex();
        if (idx < 0 || idx >= _users.Count)
        {
            MessageBox.Show("Select a user first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var newName = ShowInputDialog("Rename User", "New name:", _users[idx].Name);
        if (string.IsNullOrWhiteSpace(newName)) return;

        var oldName = _users[idx].Name;
        _users[idx].Name = newName.Trim();
        _users[idx].TokenCacheFile = GetUserCacheFileName(newName.Trim());
        SaveUsersToConfig();
        RefreshUserGrid();
        WriteLog($"User renamed: {oldName} → {newName.Trim()}", Color.Green);
    }

    private void BtnRemove_Click(object? sender, EventArgs e)
    {
        var idx = GetSelectedUserIndex();
        if (idx < 0 || idx >= _users.Count)
        {
            MessageBox.Show("Select a user first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var name = _users[idx].Name;
        if (MessageBox.Show($"Remove user '{name}'?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        _users.RemoveAt(idx);
        SaveUsersToConfig();
        RefreshUserGrid();
        WriteLog($"User removed: {name}", Color.Yellow);
    }

    private async void BtnLoginUser_Click(object? sender, EventArgs e)
    {
        var idx = GetSelectedUserIndex();
        if (idx < 0 || idx >= _users.Count)
        {
            MessageBox.Show("Select a user first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var user = _users[idx];
        var tenantId = GetEffectiveTenantId();
        var clientId = _txtClientId.Text.Trim();
        var siteUrl = _txtSite.Text.Trim();

        WriteLog($"[{user.Name}] Opening browser for login...", Color.Cyan);
        Cursor = Cursors.WaitCursor;
        _btnLogin.Enabled = false;

        try
        {
            var tokenData = await OAuthHelper.InteractiveLoginAsync(
                clientId, tenantId, siteUrl, user.Name);

            var cacheFile = Path.Combine(_appDir, user.TokenCacheFile);
            TokenCache.Save(cacheFile, tokenData);

            WriteLog($"[{user.Name}] Login successful. Token cached.", Color.Green);
            RefreshUserGrid();
        }
        catch (Exception ex)
        {
            WriteLog($"[{user.Name}] Login failed: {ex.Message}", Color.Red);
        }
        finally
        {
            Cursor = Cursors.Default;
            _btnLogin.Enabled = true;
        }
    }

    // =====================================================================
    // Event Handlers - Execution
    // =====================================================================
    private void BtnStart_Click(object? sender, EventArgs e)
    {
        var pageUrl = _txtPageUrl.Text.Trim();
        if (!string.IsNullOrEmpty(pageUrl) && !_validationActive)
        {
            MessageBox.Show("Page URL is configured. Please validate the page first.",
                "Validation Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var enabledUsers = _users.Where(u => u.Enabled).ToList();
        if (enabledUsers.Count == 0)
        {
            MessageBox.Show("No enabled users. Add and enable at least one user.", "Info",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        _isRunning = true;
        _searchTimer.Interval = GetIntervalMs();
        _searchTimer.Start();
        _countdownTimer.Start();
        _nextExecution = DateTime.Now.AddMilliseconds(GetIntervalMs());
        _btnStart.Enabled = false;
        _btnStop.Enabled = true;
        _lblStatus.Text = "Status: Running";
        WriteLog($"Scheduler started (every {GetIntervalLabel()}).", Color.LimeGreen);
    }

    private void BtnStop_Click(object? sender, EventArgs e)
    {
        _isRunning = false;
        _searchTimer.Stop();
        _countdownTimer.Stop();
        _btnStart.Enabled = true;
        _btnStop.Enabled = false;
        _lblStatus.Text = "Status: Idle";
        _lblNext.Text = "Next execution: --:--:--";
        WriteLog("Scheduler stopped.", Color.Yellow);
    }

    private async void BtnValidatePage_Click(object? sender, EventArgs e)
    {
        var pageUrl = _txtPageUrl.Text.Trim();
        if (string.IsNullOrEmpty(pageUrl))
        {
            WriteLog("No page URL configured.", Color.Yellow);
            return;
        }

        // Find first enabled user with a token
        var enabledUsers = _users.Where(u => u.Enabled).ToList();
        if (enabledUsers.Count == 0)
        {
            WriteLog("No enabled users. Add and login a user first.", Color.Yellow);
            return;
        }

        WriteLog($"Validating page: {pageUrl}", Color.Cyan);
        var tenantId = GetEffectiveTenantId();
        var clientId = _txtClientId.Text.Trim();
        var siteUrl = _txtSite.Text.Trim();

        foreach (var user in enabledUsers)
        {
            var cacheFile = Path.Combine(_appDir, user.TokenCacheFile);
            try
            {
                var token = await OAuthHelper.GetCachedOrRefreshedTokenAsync(cacheFile, siteUrl);
                if (token == null) continue;

                var result = await _searchClient.ExecuteSearchAsync(
                    siteUrl, token, $"path:\"{pageUrl}\"",
                    ["Title", "Path", "WorkId"], 1, "");

                if (result.Rows.Count > 0 && result.Rows[0].TryGetValue("WorkId", out var workId)
                    && !string.IsNullOrEmpty(workId))
                {
                    _validatedWorkId = workId;
                    _validationActive = true;
                    _validationFoundUsers.Clear();
                    _lblWorkId.Text = $"WorkId: {workId}";
                    _lblWorkId.ForeColor = Color.DarkGreen;
                    _btnResetValidation.Enabled = true;
                    _btnStart.Enabled = true;
                    _btnValidatePage.Enabled = false;
                    WriteLog($"Page validated! WorkId: {workId} (via {user.Name})", Color.LimeGreen);
                    return;
                }
            }
            catch (Exception ex)
            {
                WriteLog($"[{user.Name}] Validation error: {ex.Message}", Color.Red);
            }
        }

        WriteLog("Page not found in search index for any user.", Color.Yellow);
    }

    // =====================================================================
    // Search Execution (all users)
    // =====================================================================
    private async Task ExecuteSearchForAllUsersAsync()
    {
        if (!_isRunning && !_runOnce) return;
        bool isTestQuery = _runOnce;
        _runOnce = false;

        var enabledUsers = _users.Where(u => u.Enabled).ToList();
        if (enabledUsers.Count == 0)
        {
            WriteLog("No enabled users. Skipping.", Color.Yellow);
            return;
        }

        var clientId = "";
        var tenantId = "";
        var siteUrl = "";
        var queryText = "";
        var pageUrl = "";
        var sortList = "";
        string[] props = [];
        int rowLimit = 10;

        Invoke(() =>
        {
            clientId = _txtClientId.Text.Trim();
            tenantId = GetEffectiveTenantId();
            siteUrl = _txtSite.Text.Trim();
            queryText = _txtQuery.Text.Trim();
            pageUrl = _txtPageUrl.Text.Trim();
            sortList = _txtSortList.Text.Trim();
            props = _txtProps.Text.Trim().Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            rowLimit = (int)_numRows.Value;
        });

        if (_validationActive && !string.IsNullOrEmpty(_validatedWorkId))
        {
            if (!props.Any(p => p.Equals("WorkId", StringComparison.OrdinalIgnoreCase)))
                props = [.. props, "WorkId"];
        }

        bool allFoundThisRound = true;

        foreach (var user in enabledUsers)
        {
            var cacheFile = Path.Combine(_appDir, user.TokenCacheFile);
            WriteLog($"[{user.Name}] Acquiring token...", Color.Cyan);

            string? token;
            try
            {
                token = await OAuthHelper.GetCachedOrRefreshedTokenAsync(cacheFile, siteUrl);
                if (token == null)
                {
                    WriteLog($"[{user.Name}] No valid token. Login required.", Color.Red);
                    allFoundThisRound = false;
                    continue;
                }
                WriteLog($"[{user.Name}] Token acquired.", Color.Green);
            }
            catch (Exception ex)
            {
                WriteLog($"[{user.Name}] Token error: {ex.Message}", Color.Red);
                allFoundThisRound = false;
                continue;
            }

            try
            {
                // Pre-determine query category for request logging
                var preQueryType = isTestQuery ? "TEST"
                    : (_validationActive && !string.IsNullOrEmpty(_validatedWorkId)) ? "VALIDATE" : "QUERY";

                var result = await _searchClient.ExecuteSearchAsync(
                    siteUrl, token, queryText, props, rowLimit, sortList,
                    _requestLogDir, user.Name, preQueryType);

                WriteLog($"[{user.Name}] URL: {result.RequestUrl}", Color.Gray);
                WriteLog($"[{user.Name}] HTTP {result.StatusCode} - {result.RowCount} of {result.TotalRows} results ({result.ElapsedMs}ms)", Color.LimeGreen);

                if (!string.IsNullOrEmpty(result.InternalRequestId))
                    WriteLog($"[{user.Name}] X-SearchInternalRequestId: {result.InternalRequestId}", Color.Gray);
                if (!string.IsNullOrEmpty(result.CorrelationId))
                    WriteLog($"[{user.Name}] CorrelationId: {result.CorrelationId}", Color.Gray);

                var validationStatus = "";
                bool workIdFound = false;

                if (_validationActive && !string.IsNullOrEmpty(_validatedWorkId))
                {
                    foreach (var row in result.Rows)
                    {
                        if (row.TryGetValue("WorkId", out var wid) && wid == _validatedWorkId)
                        { workIdFound = true; break; }
                    }
                    validationStatus = workIdFound ? "FOUND" : "NOT FOUND";
                    var c = workIdFound ? Color.LimeGreen : Color.Yellow;
                    WriteLog($"[{user.Name}] >> PAGE VALIDATION: {validationStatus} (WorkId {_validatedWorkId})", c);

                    if (workIdFound && !_validationFoundUsers.Contains(user.Name))
                        _validationFoundUsers.Add(user.Name);
                    if (!workIdFound)
                        allFoundThisRound = false;

                    var foundCount = _validationFoundUsers.Count;
                    Invoke(() => _lblStatus.Text = $"Status: Validating - {foundCount}/{enabledUsers.Count} users found page");
                }

                // Determine query type for chart and TSV
                string queryType;
                if (isTestQuery)
                    queryType = result.RowCount > 0 ? "TEST OK" : "TEST NOK";
                else if (_validationActive && !string.IsNullOrEmpty(_validatedWorkId))
                    queryType = workIdFound ? "VALIDATE OK" : "VALIDATE NOK";
                else
                    queryType = result.RowCount > 0 ? "QUERY OK" : "QUERY NOK";

                for (int i = 0; i < result.Rows.Count; i++)
                {
                    WriteLog($"[{user.Name}] --- Result {i + 1} of {result.RowCount} ---", Color.White);
                    foreach (var prop in props)
                    {
                        if (result.Rows[i].TryGetValue(prop, out var val) && !string.IsNullOrEmpty(val))
                            WriteLog($"[{user.Name}]   {prop}: {val}", Color.Cyan);
                    }
                    WriteTsvRow(user.Name, result.StatusCode.ToString(), result.CorrelationId ?? "",
                        result.InternalRequestId ?? "", result.RowCount, $"{result.ElapsedMs}ms",
                        result.Rows[i], validationStatus, result.QueryIdentityDiagnostics ?? "", queryType);
                }

                if (result.RowCount == 0)
                {
                    if (_validationActive && !string.IsNullOrEmpty(_validatedWorkId) && !isTestQuery)
                    {
                        validationStatus = "NOT FOUND";
                        queryType = "VALIDATE NOK";
                        allFoundThisRound = false;
                        WriteLog($"[{user.Name}] >> PAGE VALIDATION: NOT FOUND (0 results)", Color.Yellow);
                    }
                    WriteTsvRow(user.Name, result.StatusCode.ToString(), result.CorrelationId ?? "",
                        result.InternalRequestId ?? "", 0, $"{result.ElapsedMs}ms",
                        validationStatus: validationStatus, queryIdDiag: result.QueryIdentityDiagnostics ?? "", queryType: queryType);
                }

                // Feed live chart data
                _chartDataPoints.Add(new ChartDataPoint(DateTime.Now, result.ElapsedMs, validationStatus, user.Name, queryType));
            }
            catch (Exception ex)
            {
                WriteLog($"[{user.Name}] Search error: {ex.Message}", Color.Red);
                _chartDataPoints.Add(new ChartDataPoint(DateTime.Now, 0, "ERROR", user.Name, "ERROR"));
                allFoundThisRound = false;
            }
        }

        // Auto-stop when validation active and all users found the page
        if (_validationActive && !string.IsNullOrEmpty(_validatedWorkId) && allFoundThisRound && enabledUsers.Count > 0)
        {
            WriteLog($">> PAGE VALIDATION COMPLETE: All {enabledUsers.Count} enabled user(s) can retrieve the page!", Color.LimeGreen);
            WriteLog($">> WorkId: {_validatedWorkId}", Color.LimeGreen);
            WriteLog(">> Stopping scheduler and collecting logs.", Color.LimeGreen);

            _searchTimer.Stop();
            _countdownTimer.Stop();
            _isRunning = false;
            Invoke(() =>
            {
                _btnStart.Enabled = true;
                _btnStop.Enabled = false;
                _lblStatus.Text = "Status: Validation COMPLETE - all users found page";
                _lblNext.Text = "Next execution: --:--:--";
            });

            // Auto-collect logs
            BtnCollectLogs_Click(null, EventArgs.Empty);
        }

        if (_isRunning)
            _nextExecution = DateTime.Now.AddMilliseconds(GetIntervalMs());

        Invoke(RefreshUserGrid);
    }

    // =====================================================================
    // Event Handlers - Log
    // =====================================================================
    private void BtnCollectLogs_Click(object? sender, EventArgs e)
    {
        WriteLog("Collecting logs...", Color.Cyan);
        try
        {
            var logDir = Path.Combine(_appDir, "logs");
            var zipPath = LogCollector.CollectLogs(logDir, _appDir);
            var zipSizeMB = new FileInfo(zipPath).Length / (1024.0 * 1024.0);
            WriteLog($"Created log archive: {Path.GetFileName(zipPath)} ({zipSizeMB:F2} MB)", Color.Cyan);

            // Copy zip path to clipboard
            Clipboard.SetText(zipPath);
            WriteLog("ZIP path copied to clipboard.", Color.Green);

            // Generate standalone HTML report
            string? reportPath = null;
            try
            {
                reportPath = Path.ChangeExtension(zipPath, ".html");
                var html = LogCollector.GenerateReportHtml(logDir);
                File.WriteAllText(reportPath, html);
            }
            catch (Exception ex2)
            {
                WriteLog($"HTML report generation error: {ex2.Message}", Color.Yellow);
                reportPath = null;
            }

            var hasWorkspace = !string.IsNullOrEmpty(_workspaceUrl);
            var msg = $"Log archive created:\n{zipPath}\n\nSize: {zipSizeMB:F2} MB\n\n" +
                      "The path has been copied to your clipboard.\n" +
                      (hasWorkspace ? "Click on + Add Files in the browser window that opens after you click OK." : "");

            MessageBox.Show(msg, "Collect Logs", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // Open workspace URL in browser
            if (hasWorkspace)
            {
                WriteLog("Opening workspace in browser...", Color.Yellow);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(_workspaceUrl) { UseShellExecute = true });
            }

            // Open HTML report after short delay
            if (reportPath != null && File.Exists(reportPath))
            {
                Task.Delay(3000).ContinueWith(_ =>
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(reportPath) { UseShellExecute = true });
                        Invoke(() => WriteLog("HTML report opened in browser.", Color.Green));
                    }
                    catch { }
                });
            }
        }
        catch (Exception ex)
        {
            WriteLog($"Log collection failed: {ex.Message}", Color.Red);
        }
    }

    private void ShowLiveChart()
    {
        if (_liveChart == null || _liveChart.IsDisposed)
        {
            _liveChart = new LiveChartForm(_chartDataPoints);
            _liveChart.FormClosed += (_, _) => _liveChart = null;
            _liveChart.Show();
        }
        else
        {
            _liveChart.BringToFront();
            _liveChart.Focus();
        }
    }

    // =====================================================================
    // Input Dialog Helper
    // =====================================================================
    private static string? ShowInputDialog(string title, string prompt, string defaultValue = "")
    {
        using var form = new Form
        {
            Text = title,
            Size = new Size(400, 150),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false, MinimizeBox = false
        };

        var lbl = new Label { Text = prompt, Location = new Point(12, 15), AutoSize = true };
        var txt = new TextBox { Location = new Point(12, 40), Size = new Size(360, 23), Text = defaultValue };
        var btnOk = new Button
        {
            Text = "OK", DialogResult = DialogResult.OK,
            Location = new Point(212, 75), Size = new Size(75, 28)
        };
        var btnCancel = new Button
        {
            Text = "Cancel", DialogResult = DialogResult.Cancel,
            Location = new Point(297, 75), Size = new Size(75, 28)
        };

        form.Controls.AddRange([lbl, txt, btnOk, btnCancel]);
        form.AcceptButton = btnOk;
        form.CancelButton = btnCancel;

        return form.ShowDialog() == DialogResult.OK ? txt.Text : null;
    }
}
