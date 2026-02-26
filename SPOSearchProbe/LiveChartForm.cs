namespace SPOSearchProbe;

/// <summary>
/// Data point for the live execution chart.
/// Each point represents a single search query execution with its timing and outcome.
/// QueryType values: "TEST OK", "TEST NOK", "QUERY OK", "QUERY NOK", "VALIDATE OK", "VALIDATE NOK", "ERROR"
/// </summary>
public record ChartDataPoint(DateTime Time, long ElapsedMs, string ValidationStatus, string User, string QueryType = "");

/// <summary>
/// Live chart popup window that renders search execution data in real-time using
/// GDI+ custom painting on a double-buffered panel.
///
/// Architecture:
/// - Uses a shared <see cref="List{ChartDataPoint}"/> that is populated by <see cref="MainForm"/>
///   and read by this form. The list persists across open/close cycles of the chart window.
/// - A 2-second refresh timer triggers repaint of the chart panel.
/// - When multiple users are detected in the data, the chart automatically switches
///   to a stacked layout where each user gets their own vertically-stacked sub-chart
///   with independent Y-axis scaling and its own line color.
/// - Single-user mode renders one full-size chart.
/// - Hover detection uses Euclidean distance to find the nearest data point within
///   a 20px radius across all stacked sub-charts, showing a tooltip with details.
/// - All rendering is done via GDI+ (no charting library) for zero external dependencies.
/// </summary>
public class LiveChartForm : Form
{
    /// <summary>Shared data points list — populated by MainForm, read by this chart.</summary>
    private readonly List<ChartDataPoint> _points;
    /// <summary>Lock for thread-safe access to _points (written by UI thread timers, read during paint).</summary>
    private readonly object _lock = new();
    /// <summary>The main chart rendering surface (uses DoubleBufferedPanel to prevent flicker).</summary>
    private readonly Panel _chartPanel;
    /// <summary>Statistics bar at the top showing aggregated metrics.</summary>
    private readonly Label _lblStats;
    /// <summary>Legend panel at the bottom showing color codes for statuses and per-user line colors.</summary>
    private readonly Panel _pnlLegend;
    /// <summary>Timer that triggers chart repaints every 2 seconds for "live" updates.</summary>
    private readonly System.Windows.Forms.Timer _refreshTimer;

    // --- Hover state ---
    // Tracks which data point the mouse is currently near.
    // Series index identifies the user's sub-chart, point index identifies the specific point.
    private int _hoverSeriesIdx = -1;
    private int _hoverPointIdx = -1;
    /// <summary>Tooltip control for showing data point details on hover.</summary>
    private readonly ToolTip _tip = new() { InitialDelay = 0, ReshowDelay = 0, ShowAlways = true };

    // --- Color palette ---
    // Per-user line colors are cycled if there are more than 8 users.
    // Colors are chosen for good contrast against the dark background.
    private static readonly Color[] UserLineColors =
    [
        Color.FromArgb(0, 120, 212),   // blue
        Color.FromArgb(180, 80, 220),  // purple
        Color.FromArgb(0, 180, 180),   // teal
        Color.FromArgb(220, 140, 0),   // amber
        Color.FromArgb(80, 180, 80),   // green
        Color.FromArgb(220, 80, 140),  // pink
        Color.FromArgb(100, 140, 220), // light blue
        Color.FromArgb(160, 200, 40),  // lime
    ];

    // --- Theme colors for the dark chart background ---
    private static readonly Color ColBackground = Color.FromArgb(30, 30, 30);
    private static readonly Color ColGrid = Color.FromArgb(55, 55, 55);
    private static readonly Color ColGridText = Color.FromArgb(140, 140, 140);
    private static readonly Color ColFound = Color.FromArgb(16, 124, 16);        // Green: OK / FOUND
    private static readonly Color ColNotFound = Color.FromArgb(220, 200, 0);     // Yellow: NOK / NOT FOUND
    private static readonly Color ColError = Color.FromArgb(216, 59, 1);         // Red: ERROR
    private static readonly Color ColTest = Color.White;                          // White: TEST OK
    private static readonly Color ColFoundMarker = Color.FromArgb(16, 180, 16);  // Bright green: "PAGE FOUND" marker

    /// <summary>
    /// Creates the live chart window. Pass a shared list so data persists across open/close cycles.
    /// The chart reads from this list on each repaint — no data copying required.
    /// </summary>
    /// <param name="sharedPoints">The shared data point list owned by MainForm.</param>
    public LiveChartForm(List<ChartDataPoint> sharedPoints)
    {
        _points = sharedPoints;
        Text = $"SPO Search Probe - Live Chart | {Program.AppVersion}";
        Size = new Size(900, 500);
        MinimumSize = new Size(500, 300);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = ColBackground;
        Font = new Font("Segoe UI", 9);
        AutoScaleMode = AutoScaleMode.Dpi;
        DoubleBuffered = true;

        // --- Chart panel ---
        // Added first to Controls → gets Dock.Fill after the Top/Bottom panels are added.
        // WinForms docking order: last added with Dock.Fill fills the remaining space.
        // Uses DoubleBufferedPanel (see inner class) to eliminate flicker during repaint.
        _chartPanel = new DoubleBufferedPanel
        {
            Dock = DockStyle.Fill,
            BackColor = ColBackground,
            Cursor = Cursors.Cross
        };
        _chartPanel.Paint += ChartPanel_Paint;
        _chartPanel.MouseMove += ChartPanel_MouseMove;
        _chartPanel.MouseLeave += (_, _) => ResetHover();
        _chartPanel.Resize += (_, _) => _chartPanel.Invalidate(); // Repaint on resize
        Controls.Add(_chartPanel);

        // --- Stats bar (top) ---
        _lblStats = new Label
        {
            Dock = DockStyle.Top, Height = 28,
            BackColor = Color.FromArgb(40, 40, 40),
            ForeColor = Color.FromArgb(200, 200, 200),
            Font = new Font("Segoe UI", 8.5f),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(8, 0, 0, 0),
            Text = "Waiting for data..."
        };
        Controls.Add(_lblStats);

        // --- Legend panel (bottom) ---
        // Custom-painted to show status dot colors and per-user line color swatches.
        // Dynamically updates when new users appear in the data.
        _pnlLegend = new Panel
        {
            Dock = DockStyle.Bottom, Height = 24,
            BackColor = Color.FromArgb(40, 40, 40)
        };
        _pnlLegend.Paint += LegendPaint;
        Controls.Add(_pnlLegend);

        // Refresh timer: repaints the chart every 2 seconds to show new data points.
        // This is cheaper than invalidating on every data point addition.
        _refreshTimer = new System.Windows.Forms.Timer { Interval = 2000 };
        _refreshTimer.Tick += (_, _) => { UpdateStats(); _pnlLegend.Invalidate(); _chartPanel.Invalidate(); };
        _refreshTimer.Start();

        SetChartIcon();
    }

    /// <summary>Adds a data point to the shared list (thread-safe).</summary>
    public void AddDataPoint(ChartDataPoint point) { lock (_lock) { _points.Add(point); } }
    /// <summary>Clears all data points (thread-safe).</summary>
    public void ClearData() { lock (_lock) { _points.Clear(); } }

    // =================================================================
    // Stats bar — aggregated metrics across all data points
    // =================================================================

    /// <summary>
    /// Updates the statistics label with aggregated counts, averages, and max values.
    /// Called on each refresh timer tick.
    /// </summary>
    private void UpdateStats()
    {
        if (InvokeRequired) { Invoke(UpdateStats); return; }
        lock (_lock)
        {
            if (_points.Count == 0) { _lblStats.Text = "Waiting for data..."; return; }
            var total = _points.Count;
            var tests = _points.Count(p => p.QueryType is "TEST OK" or "TEST NOK");
            var valOk = _points.Count(p => p.QueryType == "VALIDATE OK");
            var valNok = _points.Count(p => p.QueryType == "VALIDATE NOK");
            var qOk = _points.Count(p => p.QueryType == "QUERY OK");
            var qNok = _points.Count(p => p.QueryType == "QUERY NOK");
            var errors = _points.Count(p => p.QueryType == "ERROR" || p.ValidationStatus == "ERROR");
            var avgMs = _points.Average(p => p.ElapsedMs);
            var maxMs = _points.Max(p => p.ElapsedMs);
            var users = _points.Select(p => p.User).Where(u => !string.IsNullOrEmpty(u)).Distinct().Count();
            _lblStats.Text = $"Queries: {total}  |  Users: {users}  |  Avg: {avgMs:F0}ms  |  Max: {maxMs}ms  |  " +
                             $"TEST: {tests}  |  VALIDATE: {valOk}✓/{valNok}✗  |  QUERY: {qOk}✓/{qNok}✗" +
                             (errors > 0 ? $"  |  ERRORS: {errors}" : "");
        }
    }

    // =================================================================
    // Legend — status dots + per-user line swatches (when multi-user)
    // =================================================================

    /// <summary>
    /// Custom paints the legend panel with status color dots and, when multiple users
    /// are present, per-user line color swatches separated by a vertical divider.
    /// </summary>
    private void LegendPaint(object? sender, PaintEventArgs e)
    {
        var g = e.Graphics;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
        int x = 10;
        // Status dot legend entries
        DrawLegendDot(g, ref x, ColTest, "TEST OK");
        DrawLegendDot(g, ref x, ColError, "TEST NOK / ERROR");
        DrawLegendDot(g, ref x, ColFound, "OK");
        DrawLegendDot(g, ref x, ColNotFound, "NOK");
        DrawLegendDot(g, ref x, ColFoundMarker, "First PAGE FOUND");

        // Per-user line color swatches (only shown when >1 user)
        string[] users;
        lock (_lock) { users = _points.Select(p => p.User).Where(u => !string.IsNullOrEmpty(u)).Distinct().ToArray(); }
        if (users.Length > 1)
        {
            // Draw a vertical separator between status dots and user swatches
            x += 6;
            using var sep = new SolidBrush(Color.FromArgb(80, 80, 80));
            g.FillRectangle(sep, x, 4, 1, 16);
            x += 6;
            for (int i = 0; i < users.Length; i++)
                DrawLegendLine(g, ref x, UserLineColors[i % UserLineColors.Length], users[i]);
        }
    }

    // =================================================================
    // Chart painting — auto-switches between single and stacked per-user
    // =================================================================

    /// <summary>
    /// Main chart paint handler. When multiple users exist, divides the chart area
    /// into vertically stacked sub-charts (one per user) with individual Y-axis scaling.
    /// Single-user mode uses the full chart area for one chart.
    ///
    /// Each sub-chart renders:
    /// - Horizontal grid lines with Y-axis labels (ms values)
    /// - X-axis time labels (rotated 25° for readability)
    /// - A connected line through all data points
    /// - Color-coded dots at each data point (color = query outcome)
    /// - A dashed vertical "PAGE FOUND" marker at the first successful validation
    /// - A white highlight ring around the currently hovered data point
    /// </summary>
    private void ChartPanel_Paint(object? sender, PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

        // Snapshot the data under lock to avoid concurrent modification during paint
        ChartDataPoint[] allData;
        lock (_lock) { allData = [.. _points]; }

        // Show a placeholder message when there's no data yet
        if (allData.Length == 0)
        {
            using var fnt = new Font("Segoe UI", 11);
            var msg = "No search executions yet.\nStart a search to see the live chart.";
            var sz = g.MeasureString(msg, fnt);
            g.DrawString(msg, fnt, new SolidBrush(Color.FromArgb(120, 120, 120)),
                (_chartPanel.Width - sz.Width) / 2, (_chartPanel.Height - sz.Height) / 2);
            return;
        }

        // Build per-user data series for stacked sub-chart rendering
        var users = allData.Select(p => p.User).Where(u => !string.IsNullOrEmpty(u)).Distinct().ToArray();
        if (users.Length == 0) users = [""]; // Fallback: treat all data as one anonymous user

        var series = new List<(string User, ChartDataPoint[] Data, Color LineColor)>();
        for (int u = 0; u < users.Length; u++)
        {
            var uName = users[u];
            var uData = allData.Where(p => p.User == uName).ToArray();
            series.Add((uName, uData, UserLineColors[u % UserLineColors.Length]));
        }

        int count = series.Count;
        bool multi = count > 1; // Multi-user mode: stacked sub-charts
        float totalH = _chartPanel.Height;
        float gap = multi ? 8 : 0;         // Vertical gap between stacked sub-charts
        float perH = (totalH - gap * Math.Max(count - 1, 0)) / count;

        using var gridPen = new Pen(ColGrid, 1);
        using var gridFont = new Font("Segoe UI", 7.5f);
        using var gridBrush = new SolidBrush(ColGridText);
        using var userFont = new Font("Segoe UI", 8.5f, FontStyle.Bold);

        // --- Render each user's sub-chart ---
        for (int s = 0; s < count; s++)
        {
            var (userName, data, lineColor) = series[s];
            float sTop = s * (perH + gap); // Vertical offset of this sub-chart

            // Padding within each sub-chart:
            // Top: username label + marker text + dot radius clearance
            // Bottom: rotated X-axis time labels
            // Left: Y-axis labels (e.g. "1200ms")
            // Right: clearance for PAGE FOUND text overflow
            int pT = multi ? 38 : 30;
            int pB = multi ? 56 : 64;
            int pL = 62, pR = 30;

            float cw = _chartPanel.Width - pL - pR;  // Chart area width
            float ch = perH - pT - pB;               // Chart area height
            if (cw < 20 || ch < 10) continue;        // Skip if too small to render
            float aTop = sTop + pT;                   // Absolute top of the chart area

            // Draw a dotted separator line between sub-charts (not before the first one)
            if (multi && s > 0)
            {
                using var sepPen = new Pen(Color.FromArgb(70, 70, 70), 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dot };
                g.DrawLine(sepPen, 4, sTop - gap / 2, _chartPanel.Width - 4, sTop - gap / 2);
            }

            // Draw the user name label in the top-left corner (multi-user mode only)
            if (multi)
            {
                using var uBrush = new SolidBrush(lineColor);
                g.DrawString(userName, userFont, uBrush, pL, sTop + 6);
            }

            // --- Y-axis scale ---
            // Add 10% headroom so the highest point isn't clipped at the top edge.
            // Round up to nearest 100ms for clean grid line values.
            long maxMs = data.Length > 0 ? Math.Max(data.Max(d => d.ElapsedMs), 100) : 100;
            maxMs = (long)(Math.Ceiling(maxMs * 1.1 / 100.0) * 100);

            // Draw horizontal grid lines and Y-axis labels
            int gridN = multi ? 3 : 5; // Fewer grid lines in stacked mode to save space
            for (int i = 0; i <= gridN; i++)
            {
                float y = aTop + ch - ch * i / (float)gridN;
                g.DrawLine(gridPen, pL, y, pL + cw, y);
                var lbl = $"{maxMs * i / gridN}ms";
                var sz = g.MeasureString(lbl, gridFont);
                g.DrawString(lbl, gridFont, gridBrush, pL - sz.Width - 4, y - sz.Height / 2);
            }

            // Draw X-axis time labels (rotated 25° to fit more without overlapping)
            int lblStep = Math.Max(1, data.Length / (multi ? 5 : 8));
            for (int i = 0; i < data.Length; i += lblStep)
            {
                float x = pL + (data.Length > 1 ? i / (float)(data.Length - 1) : 0.5f) * cw;
                var lbl = data[i].Time.ToString("HH:mm:ss");
                g.TranslateTransform(x, aTop + ch + 4);
                g.RotateTransform(25);
                g.DrawString(lbl, gridFont, gridBrush, 0, 0);
                g.ResetTransform();
            }

            // --- Draw the connecting line between data points ---
            if (data.Length > 1)
            {
                var pts = new PointF[data.Length];
                for (int i = 0; i < data.Length; i++)
                {
                    float x = pL + i / (float)(data.Length - 1) * cw;
                    float y = aTop + ch - ch * (data[i].ElapsedMs / (float)maxMs);
                    pts[i] = new PointF(x, y);
                }
                using var lp = new Pen(lineColor, 1.8f);
                g.DrawLines(lp, pts);
            }

            // --- Draw individual data point dots ---
            // Color is determined by the query type/validation status:
            // White = TEST OK, Red = TEST NOK / ERROR, Green = VALIDATE OK / QUERY OK,
            // Yellow = VALIDATE NOK / QUERY NOK. Dot size is larger for errors/tests.
            int firstFound = -1;
            for (int i = 0; i < data.Length; i++)
            {
                float x = pL + (data.Length > 1 ? i / (float)(data.Length - 1) : 0.5f) * cw;
                float y = aTop + ch - ch * (data[i].ElapsedMs / (float)maxMs);

                // Determine dot color based on query type, with fallback to validation status
                var qt = data[i].QueryType;
                var dc = qt switch
                {
                    "TEST OK" => ColTest,
                    "TEST NOK" or "ERROR" => ColError,
                    "VALIDATE OK" or "QUERY OK" => ColFound,
                    "VALIDATE NOK" or "QUERY NOK" => ColNotFound,
                    _ => data[i].ValidationStatus switch
                    {
                        "FOUND" => ColFound,
                        "NOT FOUND" => ColNotFound,
                        "ERROR" => ColError,
                        _ => lineColor // Default: use the user's line color
                    }
                };
                // Dot size: larger for significant events (errors, tests, validations)
                float r = string.IsNullOrEmpty(qt) && string.IsNullOrEmpty(data[i].ValidationStatus) ? 3f : 4.5f;
                if (qt is "ERROR" or "TEST NOK" || data[i].ValidationStatus == "ERROR") r = 5f;
                g.FillEllipse(new SolidBrush(dc), x - r, y - r, r * 2, r * 2);

                // Track the first "found" point for the PAGE FOUND marker
                if ((qt == "VALIDATE OK" || qt == "QUERY OK" || data[i].ValidationStatus == "FOUND") && firstFound == -1)
                    firstFound = i;

                // Draw a white highlight ring around the currently hovered point
                if (s == _hoverSeriesIdx && i == _hoverPointIdx)
                    g.DrawEllipse(new Pen(Color.White, 2), x - r - 2, y - r - 2, (r + 2) * 2, (r + 2) * 2);
            }

            // --- Draw "First time PAGE FOUND" vertical marker ---
            // A dashed vertical line + circle + label marks the exact point where
            // the validated page first appeared in the search results.
            if (firstFound >= 0)
            {
                float fx = pL + (data.Length > 1 ? firstFound / (float)(data.Length - 1) : 0.5f) * cw;
                float fy = aTop + ch - ch * (data[firstFound].ElapsedMs / (float)maxMs);
                // Dashed vertical line spanning the full chart height
                using var dp2 = new Pen(ColFoundMarker, 2) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash };
                g.DrawLine(dp2, fx, aTop, fx, aTop + ch);
                // Circle around the found point
                g.DrawEllipse(new Pen(ColFoundMarker, 2), fx - 7, fy - 7, 14, 14);
                // "▼ PAGE FOUND" label above the chart, clamped to chart bounds
                using var ff = new Font("Segoe UI", 8f, FontStyle.Bold);
                var fl = "▼ PAGE FOUND";
                var flSz = g.MeasureString(fl, ff);
                float lx = Math.Clamp(fx - flSz.Width / 2, pL, pL + cw - flSz.Width);
                g.DrawString(fl, ff, new SolidBrush(ColFoundMarker), lx, aTop - 16);
            }
        }
    }

    // =================================================================
    // Mouse hover — finds nearest point across all stacked sub-charts
    // =================================================================

    /// <summary>
    /// Hit-tests all data points across all stacked sub-charts to find the one
    /// nearest to the mouse cursor (within a 20px radius). Updates the hover state
    /// and shows a tooltip with the point's details (time, ms, user, query type).
    ///
    /// The hit-testing must duplicate the coordinate calculation from ChartPanel_Paint
    /// because the painted positions aren't stored — they're recalculated each frame.
    /// </summary>
    private void ChartPanel_MouseMove(object? sender, MouseEventArgs e)
    {
        // Snapshot data under lock (same as in Paint)
        ChartDataPoint[] allData;
        lock (_lock) { allData = [.. _points]; }
        if (allData.Length == 0) { ResetHover(); return; }

        var users = allData.Select(p => p.User).Where(u => !string.IsNullOrEmpty(u)).Distinct().ToArray();
        if (users.Length == 0) users = [""];

        int count = users.Length;
        bool multi = count > 1;
        float totalH = _chartPanel.Height;
        float gap = multi ? 8 : 0;
        float perH = (totalH - gap * Math.Max(count - 1, 0)) / count;

        // Find the nearest point across all series using Euclidean distance
        int bestS = -1, bestI = -1;
        float bestDist = 20f; // Maximum hover distance in pixels
        ChartDataPoint? bestDp = null;

        for (int s = 0; s < count; s++)
        {
            var data = allData.Where(p => p.User == users[s]).ToArray();
            float sTop = s * (perH + gap);
            // Must match the padding values from ChartPanel_Paint exactly
            int pT = multi ? 38 : 30;
            int pB = multi ? 56 : 64;
            int pL = 62, pR = 30;
            float cw = _chartPanel.Width - pL - pR;
            float ch = perH - pT - pB;
            if (cw < 20 || ch < 10) continue;
            float aTop = sTop + pT;

            long maxMs = data.Length > 0 ? Math.Max(data.Max(d => d.ElapsedMs), 100) : 100;
            maxMs = (long)(Math.Ceiling(maxMs * 1.1 / 100.0) * 100);

            for (int i = 0; i < data.Length; i++)
            {
                float x = pL + (data.Length > 1 ? i / (float)(data.Length - 1) : 0.5f) * cw;
                float y = aTop + ch - ch * (data[i].ElapsedMs / (float)maxMs);
                float d = MathF.Sqrt((e.X - x) * (e.X - x) + (e.Y - y) * (e.Y - y));
                if (d < bestDist) { bestDist = d; bestS = s; bestI = i; bestDp = data[i]; }
            }
        }

        // Update hover state and trigger repaint only if it changed (avoids flicker)
        if (bestS != _hoverSeriesIdx || bestI != _hoverPointIdx)
        {
            _hoverSeriesIdx = bestS;
            _hoverPointIdx = bestI;
            _chartPanel.Invalidate();
        }

        // Show or hide the tooltip
        if (bestDp != null)
        {
            var qtLabel = !string.IsNullOrEmpty(bestDp.QueryType) ? bestDp.QueryType : bestDp.ValidationStatus;
            var text = $"#{bestI + 1}  {bestDp.Time:HH:mm:ss}\n{bestDp.ElapsedMs}ms" +
                       (!string.IsNullOrEmpty(qtLabel) ? $"  [{qtLabel}]" : "") +
                       (!string.IsNullOrEmpty(bestDp.User) ? $"\n{bestDp.User}" : "");
            _tip.Show(text, _chartPanel, e.X + 14, e.Y - 14);
        }
        else
        {
            _tip.Hide(_chartPanel);
        }
    }

    /// <summary>
    /// Resets hover state when the mouse leaves the chart area.
    /// </summary>
    private void ResetHover()
    {
        _hoverSeriesIdx = -1;
        _hoverPointIdx = -1;
        _tip.Hide(_chartPanel);
    }

    // =================================================================
    // Legend drawing helpers
    // =================================================================

    /// <summary>
    /// Draws a colored circle (dot) + label in the legend panel.
    /// The x position is advanced by the dot width + text width for sequential layout.
    /// </summary>
    private static void DrawLegendDot(Graphics g, ref int x, Color color, string label)
    {
        using var brush = new SolidBrush(color);
        using var tb = new SolidBrush(Color.FromArgb(180, 180, 180));
        using var font = new Font("Segoe UI", 7.5f);
        g.FillEllipse(brush, x, 7, 10, 10);
        x += 14;
        g.DrawString(label, font, tb, x, 5);
        x += (int)g.MeasureString(label, font).Width + 10;
    }

    /// <summary>
    /// Draws a colored line swatch + user name in the legend panel (multi-user mode).
    /// Similar to DrawLegendDot but uses a short horizontal line instead of a circle.
    /// </summary>
    private static void DrawLegendLine(Graphics g, ref int x, Color color, string label)
    {
        using var pen = new Pen(color, 2.5f);
        using var tb = new SolidBrush(Color.FromArgb(180, 180, 180));
        using var font = new Font("Segoe UI", 7.5f);
        g.DrawLine(pen, x, 12, x + 16, 12);
        x += 20;
        g.DrawString(label, font, tb, x, 5);
        x += (int)g.MeasureString(label, font).Width + 10;
    }

    // =================================================================
    // Icon / cleanup
    // =================================================================

    /// <summary>
    /// Generates a simple programmatic chart icon (line graph with a green dot)
    /// for the taskbar. Similar approach to MainForm.SetIcon() — no .ico resource needed.
    /// </summary>
    private void SetChartIcon()
    {
        try
        {
            var bmp = new Bitmap(32, 32);
            using var g = Graphics.FromImage(bmp);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            // Draw a small line chart icon
            using var pen = new Pen(Color.FromArgb(0, 120, 212), 2.5f);
            g.DrawLines(pen, [new(4, 24), new(10, 16), new(16, 20), new(22, 8), new(28, 12)]);
            // Green dot at the peak to represent a successful data point
            g.FillEllipse(new SolidBrush(Color.FromArgb(16, 124, 16)), 20, 6, 6, 6);
            Icon = Icon.FromHandle(bmp.GetHicon());
        }
        catch { }
    }

    /// <summary>
    /// Cleans up the refresh timer and tooltip when the form is closed.
    /// The shared _points list is NOT cleared — it persists for the next chart open.
    /// </summary>
    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _refreshTimer.Stop();
        _refreshTimer.Dispose();
        _tip.Dispose();
        base.OnFormClosed(e);
    }

    /// <summary>
    /// A Panel subclass with double-buffering enabled to prevent flicker during
    /// GDI+ custom painting. Standard Panel doesn't enable double-buffering by default.
    ///
    /// ControlStyles.AllPaintingInWmPaint — paint in WM_PAINT only (no WM_ERASEBKGND).
    /// ControlStyles.OptimizedDoubleBuffer — use an offscreen bitmap for painting.
    /// ControlStyles.UserPaint — the control paints itself (required for custom Paint handler).
    /// </summary>
    private class DoubleBufferedPanel : Panel
    {
        public DoubleBufferedPanel()
        {
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        }
    }
}
