using System.Text;
using System.Text.Json;

namespace SPOSearchProbe;

/// <summary>
/// End-User mode main form. Provides a single-user GUI for:
/// - Interactive OAuth2 login to SharePoint Online.
/// - One-shot and scheduled (timer-based) search query execution.
/// - Optional page validation: checks whether a specific page's WorkId appears
///   in the query results, and auto-stops + collects logs when found.
/// - Real-time colored log display and TSV data logging.
/// - Live chart popup and log collection with HTML report generation.
///
/// The entire UI is built programmatically (no .Designer.cs) so the form is
/// self-contained in a single file and fully portable.
/// </summary>
public class MainForm : Form
{
    private readonly SearchConfig _config;
    private readonly string _appDir;        // Directory where the executable lives (for logs, token cache files)
    private readonly SearchClient _searchClient = new();

    // --- UI Controls ---
    // All controls are created programmatically in the constructor (no designer file).
    private readonly GroupBox _grpAuth;
    private readonly TextBox _txtEmail;
    private readonly Button _btnLogin, _btnRunOnce, _btnStart, _btnStop;
    private readonly Label _lblTokenStatus, _lblSchedStatus;
    private readonly GroupBox _grpLog;
    private readonly RichTextBox _txtLog;
    private readonly Button _btnCollectLogs, _btnShowChart;

    // Page validation controls — only created when config.PageUrl is set.
    // Nullable because they don't exist in every configuration.
    private TextBox? _txtPageUrl;
    private Button? _btnValidatePage, _btnResetValidation;
    private Label? _lblWorkIdStatus;

    // --- Application State ---
    private string? _tokenCacheFile;        // Path to the current user's DPAPI-encrypted token file
    private bool _isRunning;                // True when the periodic search scheduler is active
    private bool _runOnce;                  // Flag for single-shot execution (bypasses scheduler)
    private DateTime? _nextExecution;       // Tracks when the next scheduled search will fire
    private string _validatedWorkId = "";   // WorkId of the page being validated (set by Validate Page)
    private bool _validationActive;         // True after a successful page validation lookup

    // --- Live Chart ---
    // Chart data is stored in a shared list so the chart window can be closed and
    // reopened without losing data. The LiveChartForm reads from the same list.
    private LiveChartForm? _liveChart;
    private readonly List<ChartDataPoint> _chartDataPoints = [];

    // --- Logging ---
    // Three log outputs per session:
    // 1. .log file — human-readable timestamped text log
    // 2. .tsv file — tab-separated structured data for analysis/charting
    // 3. requests/ directory — ZIP archives of raw HTTP request/response pairs
    private string _logFile = "";
    private string _tsvFile = "";
    private string _requestLogDir = "";

    // --- Dual Timer System ---
    // _searchTimer: fires at the configured interval to execute the next search query.
    // _countdownTimer: fires every 1 second to update the "Next in mm:ss" countdown label.
    // Two separate timers are used because the search interval can be minutes/hours,
    // but the countdown display needs per-second updates.
    private readonly System.Windows.Forms.Timer _searchTimer = new();
    private readonly System.Windows.Forms.Timer _countdownTimer = new();

    /// <summary>
    /// Constructs the End-User mode form with all controls laid out programmatically.
    /// </summary>
    /// <param name="config">The loaded search configuration.</param>
    /// <param name="appDir">Directory where the executable resides (used as base for logs and token files).</param>
    public MainForm(SearchConfig config, string appDir)
    {
        _config = config;
        _appDir = appDir;

        // SuspendLayout/ResumeLayout batches all control additions to avoid
        // redundant layout calculations during construction.
        SuspendLayout();
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi; // Support high-DPI displays
        Font = new Font("Segoe UI", 9);
        Text = $"SPO Search Probe - End User | {Program.AppTitleSuffix}";
        StartPosition = FormStartPosition.CenterScreen;
        AutoScroll = true;
        Size = new Size(960, 420);
        MinimumSize = new Size(640, 300);

        SetIcon();
        InitLogging();

        // The auth group height and log top position depend on whether page validation
        // is enabled (adds a second row of controls for the page URL).
        bool hasPageUrl = !string.IsNullOrEmpty(_config.PageUrl);
        int authHeight = hasPageUrl ? 78 : 55;
        int logTop = hasPageUrl ? 86 : 63;

        // --- Authentication & Execution Group ---
        // Contains: email field, login button, token status, run once, start/stop, schedule status
        _grpAuth = new GroupBox
        {
            Text = "Authentication and Execution",
            Location = new Point(10, 4),
            Size = new Size(924, authHeight),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };

        var lblEmail = new Label { Text = "Email:", Location = new Point(8, 20), AutoSize = true };
        _txtEmail = new TextBox
        {
            Location = new Point(48, 17), Size = new Size(190, 23),
            Text = Environment.UserName // Pre-fill with the Windows login name as a convenience
        };

        _btnLogin = new Button
        {
            Text = "Login", Location = new Point(245, 16), Size = new Size(55, 25),
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            BackColor = Color.FromArgb(220, 255, 220) // Light green to draw attention
        };
        _btnLogin.Click += BtnLogin_Click;

        _lblTokenStatus = new Label
        {
            Text = "No token", Location = new Point(305, 20), Size = new Size(155, 16),
            ForeColor = Color.Gray, Font = new Font("Segoe UI", 8)
        };

        // "Run Once" executes a single search without starting the periodic scheduler
        _btnRunOnce = new Button
        {
            Text = "Run Once", Location = new Point(470, 16), Size = new Size(72, 25), Enabled = false
        };
        _btnRunOnce.Click += (_, _) => { WriteLog("--- Running search (once) ---", Color.Cyan); _runOnce = true; _ = ExecuteSearchAsync(); };

        // "Start" begins periodic search execution at the configured interval
        _btnStart = new Button
        {
            Text = "Start", Location = new Point(548, 16), Size = new Size(50, 25), Enabled = false,
            BackColor = Color.FromArgb(220, 255, 220)
        };
        _btnStart.Click += BtnStart_Click;

        // "Stop" halts the periodic scheduler
        _btnStop = new Button
        {
            Text = "Stop", Location = new Point(602, 16), Size = new Size(45, 25), Enabled = false,
            BackColor = Color.FromArgb(255, 220, 220) // Light red to indicate "stop"
        };
        _btnStop.Click += BtnStop_Click;

        // Shows the current interval and countdown to next execution
        _lblSchedStatus = new Label
        {
            Text = $"every {_config.IntervalValue} {_config.IntervalUnit} | Idle",
            Location = new Point(655, 20), Size = new Size(260, 16),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Font = new Font("Segoe UI", 8)
        };

        _grpAuth.Controls.AddRange([lblEmail, _txtEmail, _btnLogin, _lblTokenStatus, _btnRunOnce, _btnStart, _btnStop, _lblSchedStatus]);

        // --- Optional Page Validation Row ---
        // Only shown when the admin has configured a PageUrl in search-config.json.
        // Allows the user to validate that a specific page is in the search index,
        // then monitors subsequent scheduled queries for that page's WorkId.
        if (hasPageUrl)
        {
            var lblPage = new Label
            {
                Text = "Page URL:", Location = new Point(8, 50), AutoSize = true,
                Font = new Font("Segoe UI", 8.5f)
            };
            _txtPageUrl = new TextBox
            {
                Location = new Point(70, 47), Size = new Size(590, 22),
                Text = _config.PageUrl, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 8.5f)
            };
            _btnValidatePage = new Button
            {
                Text = "Validate Page", Location = new Point(668, 46), Size = new Size(95, 24),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                BackColor = Color.FromArgb(220, 230, 255) // Light blue
            };
            _btnValidatePage.Click += BtnValidatePage_Click;

            // "Reset" clears the validated WorkId and returns to normal query mode
            _btnResetValidation = new Button
            {
                Text = "Reset", Location = new Point(768, 46), Size = new Size(50, 24),
                Anchor = AnchorStyles.Top | AnchorStyles.Right, Enabled = false
            };
            _btnResetValidation.Click += (_, _) =>
            {
                _validatedWorkId = "";
                _validationActive = false;
                _btnResetValidation!.Enabled = false;
                _lblWorkIdStatus!.Text = "";
                UpdateTokenStatus();
                WriteLog("Validation reset.", Color.Yellow);
            };

            _lblWorkIdStatus = new Label
            {
                Text = "", Location = new Point(823, 50), Size = new Size(95, 16),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                ForeColor = Color.Gray, Font = new Font("Segoe UI", 7.5f)
            };

            _grpAuth.Controls.AddRange([lblPage, _txtPageUrl, _btnValidatePage, _btnResetValidation, _lblWorkIdStatus]);
        }

        // --- Log Output Group ---
        // Contains a dark-themed RichTextBox for colored log output, plus buttons
        // for log collection and opening the live chart window.
        _grpLog = new GroupBox
        {
            Text = "Log", Location = new Point(10, logTop),
            Size = new Size(924, 370 - logTop),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        };

        _btnCollectLogs = new Button
        {
            Text = "Collect Logs",
            Location = new Point(_grpLog.Width - 100, 0), Size = new Size(80, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Font = new Font("Segoe UI", 7.5f),
            FlatStyle = FlatStyle.System
        };
        _btnCollectLogs.Click += BtnCollectLogs_Click;

        _btnShowChart = new Button
        {
            Text = "📈 Chart",
            Location = new Point(_grpLog.Width - 188, 0), Size = new Size(80, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Font = new Font("Segoe UI", 7.5f),
            FlatStyle = FlatStyle.System
        };
        _btnShowChart.Click += (_, _) => ShowLiveChart();

        // Dark-themed log output mimicking a terminal/console appearance
        _txtLog = new RichTextBox
        {
            Location = new Point(8, 18),
            Size = new Size(_grpLog.Width - 16, _grpLog.Height - 24),
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
            ReadOnly = true,
            BackColor = Color.FromArgb(30, 30, 30),
            ForeColor = Color.FromArgb(200, 200, 200),
            Font = new Font("Consolas", 9),
            WordWrap = false,
            ScrollBars = RichTextBoxScrollBars.Both
        };

        _grpLog.Controls.AddRange([_txtLog, _btnShowChart, _btnCollectLogs]);
        Controls.AddRange([_grpAuth, _grpLog]);
        ResumeLayout(true);

        // After DPI auto-scaling, the form may exceed the screen dimensions.
        // Cap it to the working area and re-center.
        Load += (_, _) =>
        {
            var wa = Screen.FromControl(this).WorkingArea;
            if (Width > wa.Width || Height > wa.Height)
                Size = new Size(Math.Min(Width, wa.Width), Math.Min(Height, wa.Height));
            CenterToScreen();
        };

        // --- Timer Setup ---
        // Countdown timer: ticks every 1 second to update the "Next in mm:ss" status label.
        _countdownTimer.Interval = 1000;
        _countdownTimer.Tick += (_, _) =>
        {
            if (_nextExecution.HasValue)
            {
                var remaining = _nextExecution.Value - DateTime.Now;
                _lblSchedStatus.Text = remaining.TotalSeconds > 0
                    ? $"every {_config.IntervalValue} {_config.IntervalUnit} | Next in {remaining:mm\\:ss}"
                    : $"every {_config.IntervalValue} {_config.IntervalUnit} | Executing...";
            }
        };
        // Search timer: fires at the configured interval to execute the actual search query.
        _searchTimer.Tick += async (_, _) => await ExecuteSearchAsync();

        // Log initial configuration info so the user can verify the setup
        WriteLog($"SPO Search Probe (.NET) started. Config: {_config.SiteUrl}", Color.LimeGreen);
        WriteLog($"Query: {_config.QueryText} | RowLimit: {_config.RowLimit} | Sort: {_config.SortList}", Color.Gray);
    }

    /// <summary>
    /// Generates a simple programmatic icon (magnifying glass with a green dot)
    /// so the app has a distinctive taskbar presence without requiring an .ico resource file.
    /// Wrapped in try/catch because GDI+ icon creation can fail on some RDP/terminal sessions.
    /// </summary>
    private void SetIcon()
    {
        try
        {
            var bmp = new Bitmap(32, 32);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            // Draw concentric circles for the magnifying glass lens
            g.DrawEllipse(new Pen(Color.FromArgb(40, 0, 120, 212), 1), 1, 1, 22, 22);
            g.DrawEllipse(new Pen(Color.FromArgb(80, 0, 120, 212), 1), 3, 3, 18, 18);
            g.DrawEllipse(new Pen(Color.FromArgb(0, 120, 212), 2.5f), 4, 4, 16, 16);
            g.FillEllipse(new SolidBrush(Color.FromArgb(30, 0, 120, 212)), 4, 4, 16, 16);
            // Draw the handle
            using var pen = new Pen(Color.FromArgb(0, 90, 170), 3)
            {
                StartCap = System.Drawing.Drawing2D.LineCap.Round,
                EndCap = System.Drawing.Drawing2D.LineCap.Round
            };
            g.DrawLine(pen, 18, 18, 28, 28);
            // Green dot in the center represents "active/running"
            g.FillEllipse(new SolidBrush(Color.FromArgb(0, 180, 80)), 10, 10, 4, 4);
            Icon = Icon.FromHandle(bmp.GetHicon());
        }
        catch { }
    }

    /// <summary>
    /// Creates the log directory structure and initializes the .log, .tsv, and request log paths.
    /// All files for a session share the same timestamp prefix so they can be correlated.
    /// The TSV file gets a header row immediately so it's valid even if no queries run.
    /// </summary>
    private void InitLogging()
    {
        var logDir = Path.Combine(_appDir, "logs");
        Directory.CreateDirectory(logDir);
        // Timestamp prefix groups all session files together (log, tsv, requests/)
        var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        _logFile = Path.Combine(logDir, $"{ts}_SPOSearchProbe.log");
        _tsvFile = Path.Combine(logDir, $"{ts}_SPOSearchProbe.tsv");
        _requestLogDir = Path.Combine(logDir, "requests", ts);

        // Write the TSV header row. These columns are designed to capture all
        // diagnostic information needed for search troubleshooting:
        // DateTime, User, HTTP status, query type, validation result, correlation IDs,
        // result count, execution time, and key managed properties from the first result.
        var tsvHeader = "DateTime\tUser\tHttpStatus\tQueryType\tValidationStatus\tCorrelationId\tX-SearchInternalRequestId\t" +
                        "ResultCount\tExecutionTime\tTitle\tPath\tLastModifiedTime\tSiteId\tWebId\tListId\t" +
                        "UniqueId\tWorkId\tRefinableString01\tQueryIdentityDiagnostics";
        try { File.WriteAllText(_tsvFile, tsvHeader + "\r\n"); } catch { }
    }

    /// <summary>
    /// Appends a timestamped, colored line to both the on-screen RichTextBox log
    /// and the persistent .log file on disk. Thread-safe via Invoke().
    /// </summary>
    private void WriteLog(string message, Color color)
    {
        // Marshal to the UI thread if called from a background/async context
        if (InvokeRequired) { Invoke(() => WriteLog(message, color)); return; }
        var ts = DateTime.Now.ToString("HH:mm:ss");
        var line = $"[{ts}] {message}";
        // RichTextBox coloring: set SelectionStart to the end, change the color,
        // then append — this makes the new text appear in the specified color.
        _txtLog.SelectionStart = _txtLog.TextLength;
        _txtLog.SelectionColor = color;
        _txtLog.AppendText(line + "\r\n");
        _txtLog.ScrollToCaret(); // Auto-scroll to the latest entry
        // Also persist to the .log file (fire-and-forget, wrapped in try/catch)
        try { File.AppendAllText(_logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\r\n"); } catch { }
    }

    /// <summary>
    /// Appends a row to the TSV data file. Each row captures a single search result
    /// (or a zero-result entry) with all diagnostic columns for later analysis.
    /// The TSV format is chosen for easy import into Excel, Power BI, or custom parsers.
    /// </summary>
    private void WriteTsvRow(string user, string httpStatus, string correlationId, string requestId,
        int resultCount, string executionTime, Dictionary<string, string>? cells = null,
        string validationStatus = "", string queryIdDiag = "", string queryType = "")
    {
        var dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        // Helper lambda to safely extract a value from the cells dictionary,
        // stripping tabs and newlines to prevent TSV format corruption.
        string Get(string key) => cells != null && cells.TryGetValue(key, out var v) ? v.Replace("\t", " ").Replace("\n", " ") : "";
        var vals = new[] { dt, user, httpStatus, queryType, validationStatus, correlationId, requestId,
            resultCount.ToString(), executionTime, Get("Title"), Get("Path"), Get("LastModifiedTime"),
            Get("SiteId"), Get("WebId"), Get("ListId"), Get("UniqueId"), Get("WorkId"),
            Get("RefinableString01"), queryIdDiag.Replace("\t", " ").Replace("\n", " ").Replace("\r", " ") };
        try { File.AppendAllText(_tsvFile, string.Join("\t", vals) + "\r\n"); } catch { }
    }

    /// <summary>
    /// Refreshes the token status label and enables/disables action buttons based on
    /// the current state of the token cache. Called after login, token refresh, and
    /// validation state changes. Thread-safe via Invoke().
    /// </summary>
    private void UpdateTokenStatus()
    {
        if (InvokeRequired) { Invoke(UpdateTokenStatus); return; }
        if (_tokenCacheFile == null)
        {
            _lblTokenStatus.Text = "Token: Not logged in";
            _lblTokenStatus.ForeColor = Color.Gray;
            _btnRunOnce.Enabled = false; _btnStart.Enabled = false;
            return;
        }
        var cache = TokenCache.Load(_tokenCacheFile);
        if (cache == null)
        {
            _lblTokenStatus.Text = "Token: Not logged in";
            _lblTokenStatus.ForeColor = Color.Gray;
            _btnRunOnce.Enabled = false; _btnStart.Enabled = false;
            return;
        }
        var exp = DateTime.Parse(cache.ExpiresOn);
        if (exp > DateTime.Now)
        {
            _lblTokenStatus.Text = $"Token: Active (expires {exp:HH:mm})";
            _lblTokenStatus.ForeColor = Color.DarkGreen;
        }
        else
        {
            // Token is expired but we still have a refresh token, so silent refresh
            // will happen automatically on the next search execution.
            _lblTokenStatus.Text = "Token: Expired (has refresh token)";
            _lblTokenStatus.ForeColor = Color.DarkOrange;
        }
        _btnRunOnce.Enabled = true;
        // "Start" is only enabled when either there's no page validation configured,
        // or the user has already successfully validated the page (so we know the WorkId).
        _btnStart.Enabled = string.IsNullOrEmpty(_config.PageUrl) || _validationActive;
    }

    // =================================================================
    // Event Handlers
    // =================================================================

    /// <summary>
    /// Handles the Login button click. Opens the browser for OAuth2 interactive login,
    /// then saves the resulting tokens to a DPAPI-encrypted cache file named after
    /// the user's email address.
    /// </summary>
    private async void BtnLogin_Click(object? sender, EventArgs e)
    {
        var email = _txtEmail.Text.Trim();
        if (string.IsNullOrEmpty(email))
        {
            MessageBox.Show("Please enter your email address.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        WriteLog("Opening browser for login...", Color.Cyan);
        Cursor = Cursors.WaitCursor;
        _btnLogin.Enabled = false;
        try
        {
            var tokenData = await OAuthHelper.InteractiveLoginAsync(
                _config.ClientId, _config.TenantId, _config.SiteUrl, email);

            // Create a per-user token cache file using a sanitized email as the filename.
            // This supports multiple users logging in on the same machine.
            var safeEmail = System.Text.RegularExpressions.Regex.Replace(email, "[^a-zA-Z0-9]", "_");
            _tokenCacheFile = Path.Combine(_appDir, $".token-user-{safeEmail}.dat");
            TokenCache.Save(_tokenCacheFile, tokenData);

            WriteLog("Login successful. Token cached locally.", Color.Green);
            UpdateTokenStatus();
        }
        catch (Exception ex)
        {
            WriteLog($"Login failed: {ex.Message}", Color.Red);
        }
        finally
        {
            Cursor = Cursors.Default;
            _btnLogin.Enabled = true;
        }
    }

    /// <summary>
    /// Starts the periodic search scheduler. Both timers are started:
    /// the search timer fires at the configured interval, and the countdown timer
    /// updates the status label every second.
    /// </summary>
    private void BtnStart_Click(object? sender, EventArgs e)
    {
        _isRunning = true;
        _searchTimer.Interval = _config.GetIntervalMs();
        _searchTimer.Start();
        _countdownTimer.Start();
        _nextExecution = DateTime.Now.AddMilliseconds(_config.GetIntervalMs());
        _btnStart.Enabled = false;
        _btnStop.Enabled = true;
        WriteLog($"Scheduler started (every {_config.IntervalValue} {_config.IntervalUnit}).", Color.LimeGreen);
    }

    /// <summary>
    /// Stops the periodic search scheduler and resets the UI to idle state.
    /// </summary>
    private void BtnStop_Click(object? sender, EventArgs e)
    {
        _isRunning = false;
        _searchTimer.Stop();
        _countdownTimer.Stop();
        _btnStart.Enabled = true;
        _btnStop.Enabled = false;
        _lblSchedStatus.Text = $"every {_config.IntervalValue} {_config.IntervalUnit} | Idle";
        WriteLog("Scheduler stopped.", Color.Yellow);
    }

    /// <summary>
    /// Validates whether a specific page URL is in the search index.
    /// Uses a targeted "path:" query to find the page and extract its WorkId.
    /// If found, subsequent scheduled queries will check whether that WorkId
    /// appears in the configured query's results (freshness validation).
    /// </summary>
    private async void BtnValidatePage_Click(object? sender, EventArgs e)
    {
        if (_tokenCacheFile == null) { WriteLog("Login first.", Color.Yellow); return; }
        var pageUrl = _txtPageUrl?.Text.Trim();
        if (string.IsNullOrEmpty(pageUrl)) { WriteLog("No page URL.", Color.Yellow); return; }

        WriteLog("=== PAGE VALIDATION ===", Color.White);
        WriteLog($"Page URL: {pageUrl}", Color.Cyan);
        WriteLog("Step 1: Checking if page is crawled (path: query)...", Color.Yellow);
        _btnValidatePage!.Enabled = false;
        try
        {
            var token = await OAuthHelper.GetCachedOrRefreshedTokenAsync(_tokenCacheFile, _config.SiteUrl);
            if (token == null) { WriteLog("Token expired. Login again.", Color.Red); return; }

            // Execute a targeted search for the specific page by its full URL
            var result = await _searchClient.ExecuteSearchAsync(
                _config.SiteUrl, token, $"path:\"{pageUrl}\"",
                ["Title", "Path", "WorkId", "LastModifiedTime", "UniqueId"], 1, "",
                _requestLogDir, "VALIDATION", "VALIDATE");

            if (result.Rows.Count > 0 && result.Rows[0].TryGetValue("WorkId", out var workId) && !string.IsNullOrEmpty(workId))
            {
                // Page is in the index — save its WorkId for monitoring
                var title = result.Rows[0].GetValueOrDefault("Title", "");
                _validatedWorkId = workId;
                _validationActive = true;
                _lblWorkIdStatus!.Text = $"WID: {workId}";
                _lblWorkIdStatus.ForeColor = Color.DarkBlue;
                _btnResetValidation!.Enabled = true;
                WriteLog($">> Page found! Title: {title}", Color.LimeGreen);
                WriteLog($">> WorkId: {workId}", Color.LimeGreen);
                WriteLog("Step 2: Use 'Start' or 'Run Once' to monitor the real query.", Color.Yellow);
                WriteLog($"  The tool will check if WorkId {workId} appears in the results.", Color.Cyan);
                WriteLog("  Monitoring will auto-stop when the page is found.", Color.Cyan);
                UpdateTokenStatus(); // Re-evaluate button states (Start may now be enabled)
            }
            else
            {
                // Page is NOT in the index — may not have been crawled yet
                WriteLog(">> Page NOT found via path: query. It may not be crawled yet.", Color.Red);
                WriteLog(">> Try again later, or verify the page URL is correct.", Color.Yellow);
                _lblWorkIdStatus!.Text = "Not crawled";
                _lblWorkIdStatus.ForeColor = Color.Red;
            }
        }
        catch (Exception ex)
        {
            WriteLog($"Validation error: {ex.Message}", Color.Red);
        }
        finally
        {
            // Only re-enable the Validate button if validation didn't succeed
            // (if it did, the user should use Start/Run Once instead)
            if (!_validationActive) _btnValidatePage!.Enabled = true;
        }
    }

    /// <summary>
    /// Collects all log files into a ZIP archive, generates a standalone HTML report,
    /// copies the ZIP path to clipboard, and optionally opens the DTM workspace URL
    /// and the HTML report in the browser.
    /// </summary>
    private void CollectLogsAndOpenReport()
    {
        WriteLog("Collecting logs...", Color.Cyan);
        try
        {
            var logDir = Path.Combine(_appDir, "logs");
            var zipPath = LogCollector.CollectLogs(logDir, _appDir);
            var zipSizeMB = new FileInfo(zipPath).Length / (1024.0 * 1024.0);
            WriteLog($"Created log archive: {Path.GetFileName(zipPath)} ({zipSizeMB:F2} MB)", Color.Cyan);

            // Copy the ZIP path to clipboard so the user can paste it into a file upload dialog
            Clipboard.SetText(zipPath);
            WriteLog("ZIP path copied to clipboard.", Color.Green);

            // Generate a standalone HTML report with embedded charts next to the ZIP
            string? reportPath = null;
            try
            {
                reportPath = Path.ChangeExtension(zipPath, ".html");
                var html = LogCollector.GenerateReportHtml(logDir);
                File.WriteAllText(reportPath, html);
            }
            catch (Exception ex)
            {
                WriteLog($"HTML report generation error: {ex.Message}", Color.Yellow);
                reportPath = null;
            }

            // If a workspace URL is configured, guide the user to upload the ZIP there
            var workspaceUrl = _config.WorkspaceUrl;
            var hasWorkspace = !string.IsNullOrEmpty(workspaceUrl);

            var msg = $"Log archive created:\n{zipPath}\n\nSize: {zipSizeMB:F2} MB\n\n" +
                      "The path has been copied to your clipboard.\n" +
                      (hasWorkspace ? "Click on + Add Files in the browser window that opens after you click OK." : "Send this file to your admin.");

            MessageBox.Show(msg, "Collect Logs", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // Open the DTM workspace URL in browser (after the user dismisses the dialog)
            if (hasWorkspace)
            {
                WriteLog("Opening workspace in browser...", Color.Yellow);
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(workspaceUrl) { UseShellExecute = true });
            }

            // Open the HTML report in the browser after a short delay to avoid
            // opening too many browser tabs simultaneously
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

    /// <summary>Click handler for the "Collect Logs" button — delegates to <see cref="CollectLogsAndOpenReport"/>.</summary>
    private void BtnCollectLogs_Click(object? sender, EventArgs e)
    {
        CollectLogsAndOpenReport();
    }

    /// <summary>
    /// Opens or brings to front the live chart popup window. The chart shares the
    /// _chartDataPoints list with this form, so data persists across open/close cycles.
    /// </summary>
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

    /// <summary>
    /// Core search execution method. Called by both the periodic scheduler timer
    /// and the "Run Once" button. Handles:
    /// 1. Token acquisition (cached or silently refreshed).
    /// 2. Search query execution via <see cref="SearchClient"/>.
    /// 3. Result logging to the console, .log file, and .tsv file.
    /// 4. Page validation: checks if the validated WorkId appears in results.
    /// 5. Auto-stop: halts the scheduler and collects logs when validation succeeds.
    /// 6. Feeding data points to the live chart.
    /// </summary>
    private async Task ExecuteSearchAsync()
    {
        // Guard: only execute if the scheduler is running or a one-shot was requested
        if (!_isRunning && !_runOnce) return;
        bool isTestQuery = _runOnce;
        _runOnce = false; // Reset the one-shot flag

        // Read the email from the UI thread (we may be on a timer callback thread)
        var email = "";
        Invoke(() => email = _txtEmail.Text.Trim());

        if (_tokenCacheFile == null) { WriteLog("No token. Please login first.", Color.Yellow); return; }

        // --- Token Acquisition ---
        // First call validates the token is available; the actual token for the request
        // is fetched again below. This two-step pattern provides a clean error message
        // before attempting the search.
        try
        {
            var token = await OAuthHelper.GetCachedOrRefreshedTokenAsync(_tokenCacheFile, _config.SiteUrl);
            if (token == null) { WriteLog($"[{email}] Token expired. Please login again.", Color.Red); return; }
            WriteLog($"[{email}] Token acquired.", Color.Green);
        }
        catch (Exception ex)
        {
            WriteLog($"[{email}] Token error: {ex.Message}", Color.Red);
            return;
        }

        var accessToken = await OAuthHelper.GetCachedOrRefreshedTokenAsync(_tokenCacheFile, _config.SiteUrl);
        if (accessToken == null) return;

        // If page validation is active, ensure WorkId is in the select properties
        // so we can check whether the validated page appears in the results.
        var props = _config.SelectProperties.ToList();
        if (_validationActive && !string.IsNullOrEmpty(_validatedWorkId))
        {
            if (!props.Any(p => p.Equals("WorkId", StringComparison.OrdinalIgnoreCase)))
                props.Add("WorkId");
        }

        try
        {
            // Determine query category for log file naming and TSV data
            var preQueryType = isTestQuery ? "TEST"
                : (_validationActive && !string.IsNullOrEmpty(_validatedWorkId)) ? "VALIDATE" : "QUERY";

            // --- Execute the search query ---
            var result = await _searchClient.ExecuteSearchAsync(
                _config.SiteUrl, accessToken, _config.QueryText,
                [.. props], _config.RowLimit, _config.SortList,
                _requestLogDir, email, preQueryType);

            // --- Log the results ---
            WriteLog($"[{email}] URL: {result.RequestUrl}", Color.Gray);
            WriteLog($"[{email}] HTTP {result.StatusCode} - {result.RowCount} of {result.TotalRows} results ({result.ElapsedMs}ms)", Color.LimeGreen);

            if (!string.IsNullOrEmpty(result.InternalRequestId))
                WriteLog($"[{email}] X-SearchInternalRequestId: {result.InternalRequestId}", Color.Gray);
            if (!string.IsNullOrEmpty(result.CorrelationId))
                WriteLog($"[{email}] CorrelationId: {result.CorrelationId}", Color.Gray);

            // --- Page Validation: check if the WorkId appears in the results ---
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
                WriteLog($"[{email}] >> PAGE VALIDATION: {validationStatus} (WorkId {_validatedWorkId})", c);
            }

            // --- Determine the final query type label for chart and TSV ---
            // Combines the query category (TEST/VALIDATE/QUERY) with the success/failure result.
            string queryType;
            if (isTestQuery)
                queryType = result.RowCount > 0 ? "TEST OK" : "TEST NOK";
            else if (_validationActive && !string.IsNullOrEmpty(_validatedWorkId))
                queryType = workIdFound ? "VALIDATE OK" : "VALIDATE NOK";
            else
                queryType = result.RowCount > 0 ? "QUERY OK" : "QUERY NOK";

            // --- Log individual result rows ---
            for (int i = 0; i < result.Rows.Count; i++)
            {
                WriteLog($"[{email}] --- Result {i + 1} of {result.RowCount} ---", Color.White);
                foreach (var prop in props)
                {
                    if (result.Rows[i].TryGetValue(prop, out var val) && !string.IsNullOrEmpty(val))
                        WriteLog($"[{email}]   {prop}: {val}", Color.Cyan);
                }
                WriteTsvRow(email, result.StatusCode.ToString(), result.CorrelationId ?? "",
                    result.InternalRequestId ?? "", result.RowCount, $"{result.ElapsedMs}ms",
                    result.Rows[i], validationStatus, result.QueryIdentityDiagnostics ?? "", queryType);
            }

            // Handle zero-result case (still needs a TSV row for the chart/report)
            if (result.RowCount == 0)
            {
                if (_validationActive && !string.IsNullOrEmpty(_validatedWorkId) && !isTestQuery)
                {
                    validationStatus = "NOT FOUND";
                    queryType = "VALIDATE NOK";
                    WriteLog($"[{email}] >> PAGE VALIDATION: NOT FOUND (0 results)", Color.Yellow);
                }
                WriteTsvRow(email, result.StatusCode.ToString(), result.CorrelationId ?? "",
                    result.InternalRequestId ?? "", 0, $"{result.ElapsedMs}ms",
                    validationStatus: validationStatus, queryIdDiag: result.QueryIdentityDiagnostics ?? "", queryType: queryType);
            }

            // --- Feed data to the live chart ---
            _chartDataPoints.Add(new ChartDataPoint(DateTime.Now, result.ElapsedMs, validationStatus, email, queryType));

            // --- Auto-stop on validation success ---
            // When page validation is active and the WorkId was found in the query
            // results, the goal is achieved: stop the scheduler and collect logs
            // automatically so the user gets immediate results.
            if (_validationActive && !string.IsNullOrEmpty(_validatedWorkId) && workIdFound)
            {
                WriteLog(">> PAGE VALIDATION COMPLETE: Page found via query!", Color.LimeGreen);
                WriteLog($">> WorkId: {_validatedWorkId}", Color.LimeGreen);
                if (_isRunning)
                {
                    WriteLog(">> Stopping scheduler and collecting logs.", Color.LimeGreen);
                    _searchTimer.Stop(); _countdownTimer.Stop();
                    _isRunning = false;
                    Invoke(() => { _btnStart.Enabled = true; _btnStop.Enabled = false; _lblSchedStatus.Text = "Validation COMPLETE"; });
                }

                // Auto-collect logs and open the report — saves the user a manual step
                Invoke(() => CollectLogsAndOpenReport());
            }

            // Update the countdown target for the next execution cycle
            if (_isRunning)
                _nextExecution = DateTime.Now.AddMilliseconds(_config.GetIntervalMs());

            Invoke(UpdateTokenStatus);
        }
        catch (Exception ex)
        {
            WriteLog($"[{email}] Search error: {ex.Message}", Color.Red);
            // Record error in chart data so it's visible in the timeline
            _chartDataPoints.Add(new ChartDataPoint(DateTime.Now, 0, "ERROR", email, "ERROR"));
        }
    }
}
