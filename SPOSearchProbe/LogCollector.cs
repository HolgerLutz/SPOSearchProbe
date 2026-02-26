using System.IO.Compression;
using System.Text;
using System.Text.Json;

namespace SPOSearchProbe;

/// <summary>
/// Collects all session log files into a single ZIP archive and generates a standalone
/// HTML report with embedded charts. The report uses only vanilla HTML/CSS/JavaScript
/// with a canvas element for the execution timeline chart — no external
/// dependencies or CDN links — so it works fully offline.
///
/// ZIP archive structure:
///   ├── {timestamp}_SPOSearchProbe.log    — Human-readable session log(s)
///   ├── {timestamp}_SPOSearchProbe.tsv    — Tab-separated query data for analysis
///   ├── requests/{session}/               — Per-query request/response ZIP archives
///   │   ├── {ts}_{user}_{type}_request.txt
///   │   └── {ts}_{user}_{type}_response.json
///   └── summary.html                      — Self-contained HTML report with charts
/// </summary>
public static class LogCollector
{
    /// <summary>
    /// Collects all log files from <paramref name="logDir"/> and its subdirectories
    /// into a single ZIP archive placed in <paramref name="appDir"/>.
    /// </summary>
    /// <param name="logDir">The logs/ directory containing .log, .tsv files and requests/ subdirectory.</param>
    /// <param name="appDir">The application root directory where the output ZIP is written.</param>
    /// <returns>Full path to the created ZIP archive.</returns>
    public static string CollectLogs(string logDir, string appDir)
    {
        if (!Directory.Exists(logDir))
            throw new DirectoryNotFoundException($"Log directory not found: {logDir}");

        var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var zipName = $"SPOSearchProbe_Logs_{ts}.zip";
        var zipPath = Path.Combine(appDir, zipName);

        // Build the ZIP in memory first, then write to disk in one shot.
        // This avoids partial/corrupt files if the process is interrupted.
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
        {
            // Add all .log and .tsv files from the top-level logs/ directory
            foreach (var file in Directory.GetFiles(logDir, "*.*", SearchOption.TopDirectoryOnly)
                         .Where(f => f.EndsWith(".log", StringComparison.OrdinalIgnoreCase) ||
                                     f.EndsWith(".tsv", StringComparison.OrdinalIgnoreCase)))
            {
                archive.CreateEntryFromFile(file, Path.GetFileName(file), CompressionLevel.Optimal);
            }

            // Add per-query request/response ZIP archives from the requests/ subdirectory.
            // These are nested ZIPs (ZIPs inside the main ZIP) — this preserves the
            // per-request grouping while keeping everything in one deliverable package.
            var reqDir = Path.Combine(logDir, "requests");
            if (Directory.Exists(reqDir))
            {
                foreach (var subDir in Directory.GetDirectories(reqDir))
                {
                    foreach (var file in Directory.GetFiles(subDir, "*.zip"))
                    {
                        // Preserve the directory structure: requests/{session}/{file}.zip
                        var relPath = $"requests/{Path.GetFileName(subDir)}/{Path.GetFileName(file)}";
                        archive.CreateEntryFromFile(file, relPath, CompressionLevel.Optimal);
                    }
                }
            }

            // Add a self-contained HTML summary report inside the ZIP
            var summary = GenerateSummaryHtml(logDir);
            var entry = archive.CreateEntry("summary.html", CompressionLevel.Optimal);
            using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
            writer.Write(summary);
        }

        File.WriteAllBytes(zipPath, ms.ToArray());
        return zipPath;
    }

    /// <summary>
    /// Internal wrapper that delegates to <see cref="GenerateReportHtml"/>.
    /// Kept as a separate method for clarity — "summary" is the in-ZIP version,
    /// while the public method is also used for the standalone browser report.
    /// </summary>
    private static string GenerateSummaryHtml(string logDir)
    {
        return GenerateReportHtml(logDir);
    }

    /// <summary>
    /// Generates a standalone HTML report from the TSV/log data in <paramref name="logDir"/>.
    /// Used both for embedding in the ZIP archive (summary.html) and for opening
    /// directly in the browser after log collection.
    ///
    /// The report includes:
    /// - Overview statistics (total queries, users, success/error counts, timestamps)
    /// - Search configuration summary (read from search-config.json)
    /// - Interactive execution timeline chart (canvas-based, no external dependencies)
    /// - Tabular search results (last 50 rows)
    /// - Log package file listing
    /// - Full session log content
    ///
    /// All sections are collapsible. The chart is rendered using a vanilla JavaScript
    /// canvas painter that draws grid lines, data points, connecting lines, and a
    /// "PAGE FOUND" marker with hover tooltips — all without any charting library.
    /// </summary>
    /// <param name="logDir">The logs/ directory containing .tsv and .log files.</param>
    /// <returns>A complete HTML document as a string.</returns>
    public static string GenerateReportHtml(string logDir)
    {
        // --- Parse TSV data ---
        // Only use the most recent TSV file that contains data rows.
        // Each probe session creates its own timestamped TSV file; combining
        // multiple sessions would produce a misleading chart with data gaps.
        var tsvRows = new List<Dictionary<string, string>>();
        string[] tsvHeaders = [];
        var tsvFiles = Directory.Exists(logDir)
            ? Directory.GetFiles(logDir, "*.tsv", SearchOption.TopDirectoryOnly)
            : [];
        // Sort descending so the newest file (by timestamp-based name) is tried first
        foreach (var tsvFile in tsvFiles.OrderByDescending(f => f))
        {
            try
            {
                // Use ReadFileShared to avoid locking conflicts with the running probe
                var lines = ReadFileShared(tsvFile)
                    .Split('\n')
                    .Select(l => l.TrimEnd('\r'))
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .ToArray();
                if (lines.Length <= 1) continue; // Skip files with only a header
                var hdr = lines[0].Split('\t');
                if (hdr.Length < 2) continue; // not a valid TSV header
                tsvHeaders = hdr;
                // Parse each data row into a case-insensitive dictionary.
                // Skip rows that don't have enough columns (corrupted by multi-line
                // field values like QueryIdentityDiagnostics JSON).
                for (int i = 1; i < lines.Length; i++)
                {
                    var vals = lines[i].Split('\t');
                    if (vals.Length < tsvHeaders.Length) continue; // malformed row
                    var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    for (int j = 0; j < tsvHeaders.Length && j < vals.Length; j++)
                        row[tsvHeaders[j]] = vals[j];
                    tsvRows.Add(row);
                }
                break; // Use only the most recent valid TSV file
            }
            catch { /* skip unreadable files */ }
        }

        // --- Compute summary statistics ---
        // Deduplicate TSV rows to one per query execution using CorrelationId.
        // Multiple result rows from the same query share the same CorrelationId.
        var executions = tsvRows
            .GroupBy(r => r.GetValueOrDefault("CorrelationId", Guid.NewGuid().ToString()))
            .Select(g =>
            {
                var first = g.First();
                // Prefer FOUND status if any result row in this execution had it
                var bestStatus = g.Any(r => r.GetValueOrDefault("ValidationStatus", "") == "FOUND")
                    ? "FOUND"
                    : first.GetValueOrDefault("ValidationStatus", "");
                return new { Row = first, ValidationStatus = bestStatus };
            })
            .OrderBy(e => e.Row.GetValueOrDefault("DateTime", ""))
            .ThenBy(e => e.Row.GetValueOrDefault("User", ""))
            .ToList();

        int totalQueries = executions.Count;
        var uniqueUsers = executions
            .Select(e => e.Row.GetValueOrDefault("User", ""))
            .Where(u => !string.IsNullOrEmpty(u))
            .Distinct()
            .ToList();
        int httpOk = executions.Count(e => e.Row.GetValueOrDefault("HttpStatus", "") == "200");
        int httpErr = totalQueries - httpOk;
        // First and last timestamps for the "Probe Start" and "Page Found" stats
        string probeStartTime = executions.Count > 0
            ? Esc(executions[0].Row.GetValueOrDefault("DateTime", "-"))
            : "-";
        var firstFound = executions.FirstOrDefault(e => e.ValidationStatus == "FOUND");
        string pageFoundTime = firstFound != null
            ? Esc(firstFound.Row.GetValueOrDefault("DateTime", "-"))
            : "-";
        // CSS classes for conditional coloring of stat cards
        string statCssSuccErr = httpErr > 0 ? " warn" : " ok";
        string statCssFound = pageFoundTime != "-" ? " ok" : " warn";

        // --- Read search configuration (best-effort) ---
        // The config file is expected to be one level up from the logs/ directory
        string siteUrl = "", tenantId = "", queryText = "", selectProps = "", rowLimit = "", sortList = "", pageUrl = "";
        string rawConfigJson = "";
        try
        {
            var configPath = Path.Combine(Directory.GetParent(logDir)?.FullName ?? logDir, "search-config.json");
            if (File.Exists(configPath))
            {
                rawConfigJson = ReadFileShared(configPath);
                var cfg = JsonSerializer.Deserialize<SearchConfig>(rawConfigJson);
                if (cfg != null)
                {
                    siteUrl = cfg.SiteUrl;
                    tenantId = cfg.TenantId;
                    queryText = cfg.QueryText;
                    selectProps = string.Join(", ", cfg.SelectProperties ?? []);
                    rowLimit = cfg.RowLimit.ToString();
                    sortList = string.IsNullOrEmpty(cfg.SortList) ? "(default)" : cfg.SortList;
                    pageUrl = cfg.PageUrl;
                }
            }
        }
        catch { /* config read is best-effort */ }

        // --- Build chart data as a JSON array for the canvas painter ---
        // Each point has: t (timestamp), ms (execution time), v (validation status), u (user)
        // Deduplication already done above via CorrelationId grouping into `executions`.
        var chartPoints = new StringBuilder("[");
        for (int ci = 0; ci < executions.Count; ci++)
        {
            var r = executions[ci].Row;
            int execMs = 0;
            var execRaw = r.GetValueOrDefault("ExecutionTime", "");
            if (!string.IsNullOrEmpty(execRaw))
                int.TryParse(new string(execRaw.Where(char.IsDigit).ToArray()), out execMs);
            var vs = executions[ci].ValidationStatus;
            var usr = r.GetValueOrDefault("User", "");
            if (ci > 0) chartPoints.Append(',');
            chartPoints.Append($"{{t:\"{EscJs(r.GetValueOrDefault("DateTime", ""))}\",ms:{execMs},v:\"{EscJs(vs)}\",u:\"{EscJs(usr)}\"}}");
        }
        chartPoints.Append(']');

        // --- Build the result table rows (limited to the last 50 for readability) ---
        var displayRows = tsvRows.Count > 50 ? tsvRows.Skip(tsvRows.Count - 50).ToList() : tsvRows;
        var resultRowsSb = new StringBuilder();
        foreach (var r in displayRows)
        {
            var vs = r.GetValueOrDefault("ValidationStatus", "");
            // Apply row highlighting: green for FOUND, orange for NOT FOUND
            string rowClass = vs == "FOUND" ? " class='found'" : vs == "NOT FOUND" ? " class='notfound'" : "";
            resultRowsSb.AppendLine($"<tr{rowClass}><td>{Esc(r.GetValueOrDefault("DateTime", ""))}</td><td>{Esc(r.GetValueOrDefault("User", ""))}</td><td>{Esc(r.GetValueOrDefault("HttpStatus", ""))}</td><td>{Esc(vs)}</td><td>{Esc(r.GetValueOrDefault("ResultCount", ""))}</td><td>{Esc(r.GetValueOrDefault("ExecutionTime", ""))}</td><td>{Esc(r.GetValueOrDefault("Title", ""))}</td><td>{Esc(r.GetValueOrDefault("Path", ""))}</td><td>{Esc(r.GetValueOrDefault("WorkId", ""))}</td></tr>");
        }

        // --- Build the package file listing table ---
        var packageRowsSb = new StringBuilder();
        if (Directory.Exists(logDir))
        {
            foreach (var f in Directory.GetFiles(logDir, "*.*", SearchOption.TopDirectoryOnly)
                         .Where(f => f.EndsWith(".log", StringComparison.OrdinalIgnoreCase) ||
                                     f.EndsWith(".tsv", StringComparison.OrdinalIgnoreCase))
                         .OrderBy(f => f))
            {
                var fi = new FileInfo(f);
                string ftype = fi.Extension.Equals(".log", StringComparison.OrdinalIgnoreCase) ? "Session Log" : "TSV Data";
                packageRowsSb.AppendLine($"<tr><td>{Esc(fi.Name)}</td><td>{Math.Round(fi.Length / 1024.0, 1)} KB</td><td>{ftype}</td></tr>");
            }
            // Include request archive files in the listing
            var reqDir = Path.Combine(logDir, "requests");
            if (Directory.Exists(reqDir))
            {
                foreach (var rf in Directory.GetFiles(reqDir, "*.zip", SearchOption.AllDirectories).OrderBy(f => f))
                {
                    var rfi = new FileInfo(rf);
                    var relName = Path.GetRelativePath(logDir, rf);
                    packageRowsSb.AppendLine($"<tr><td>{Esc(relName)}</td><td>{Math.Round(rfi.Length / 1024.0, 1)} KB</td><td>Request Archive</td></tr>");
                }
            }
        }

        // --- Read full log file content for the "Full Session Log" section ---
        var logContentSb = new StringBuilder();
        if (Directory.Exists(logDir))
        {
            foreach (var logFile in Directory.GetFiles(logDir, "*.log", SearchOption.TopDirectoryOnly).OrderByDescending(f => f))
            {
                try
                {
                    logContentSb.AppendLine(Esc(ReadFileShared(logFile)));
                }
                catch { logContentSb.AppendLine("(Could not read log file)"); }
            }
        }

        // --- Conditional "Affected Page URL" section ---
        string pageUrlHtml = "";
        if (!string.IsNullOrEmpty(pageUrl))
            pageUrlHtml = $"<div class='page-url'><strong>Affected Page URL:</strong> {Esc(pageUrl)}</div>";

        // --- Collapsed state: auto-collapse empty sections ---
        string timelineCollapsed = executions.Count == 0 ? " collapsed" : "";
        string timelineHidden = executions.Count == 0 ? " style=\"display:none\"" : "";
        // Legend entry for the "First discovery" dashed line (only shown if found)
        bool hasFound = executions.Any(e => e.ValidationStatus == "FOUND");
        string chartFoundMarkerLegend = hasFound
            ? "<span style='border-left:2px dashed #107c10;padding-left:6px'> First discovery (dashed line)</span>"
            : "";

        // --- Raw config JSON for display ---
        string rawConfigSection = "";
        if (!string.IsNullOrEmpty(rawConfigJson))
            rawConfigSection = $@"
<div class=""section"">
<div class=""section-header collapsed"">search-config.json <span class=""arrow"">&#9660;</span></div>
<div class=""section-body"" style=""display:none"">
<pre>{Esc(rawConfigJson)}</pre>
</div>
</div>";

        // --- Build the complete HTML document ---
        // The report is fully self-contained: all CSS is inline, and the chart is
        // rendered via a vanilla JavaScript canvas painter embedded in a <script> tag.
        // No external dependencies, CDN links, or frameworks are used.
        var sb = new StringBuilder();
        sb.Append($@"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""UTF-8"">
<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
<title>SPO Search Probe - Session Report</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',system-ui,sans-serif;background:#f5f6f8;color:#1a1a1a;line-height:1.5;padding:20px}}
.container{{max-width:1200px;margin:0 auto}}
.header{{display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap;margin-bottom:4px}}
h1{{font-size:1.4em;color:#0078d4}}
.tool-info{{font-size:.82em;color:#555}}
.subtitle{{color:#666;font-size:.85em;margin-bottom:16px}}
.section{{background:#fff;border:1px solid #e0e0e0;border-radius:6px;margin-bottom:12px;overflow:hidden}}
.section-header{{background:#f0f2f5;padding:10px 16px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;font-weight:600;font-size:.95em;user-select:none}}
.section-header:hover{{background:#e4e8ed}}
.section-header .arrow{{transition:transform .2s;font-size:.7em;color:#888}}
.section-header.collapsed .arrow{{transform:rotate(-90deg)}}
.section-body{{padding:14px 16px}}
.section-header.collapsed+.section-body{{display:none}}
.stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin-bottom:8px}}
.stat{{background:#f8f9fa;border:1px solid #e8e8e8;border-radius:5px;padding:10px 14px;text-align:center}}
.stat .val{{font-size:1.6em;font-weight:700;color:#0078d4}}
.stat .lbl{{font-size:.78em;color:#666;margin-top:2px}}
.stat.warn .val{{color:#d83b01}}
.stat.ok .val{{color:#107c10}}
table{{width:100%;border-collapse:collapse;font-size:.82em}}
th{{background:#f0f2f5;text-align:left;padding:6px 10px;border-bottom:2px solid #d0d0d0;white-space:nowrap}}
td{{padding:5px 10px;border-bottom:1px solid #eee;word-break:break-all}}
tr:hover{{background:#f5f8fc}}
tr.found td{{background:#f0fff0;color:#107c10}}
tr.notfound td{{background:#fff8f0;color:#d83b01}}
.config-grid{{display:grid;grid-template-columns:120px 1fr;gap:4px 12px;font-size:.88em}}
.config-grid .label{{font-weight:600;color:#555}}
.config-grid .value{{word-break:break-all}}
.page-url{{background:#fff8e1;border:1px solid #ffd54f;border-radius:4px;padding:8px 12px;margin-top:8px;font-size:.88em;word-break:break-all}}
pre{{background:#1e1e1e;color:#d4d4d4;padding:12px;border-radius:5px;font-size:.78em;max-height:400px;overflow:auto;white-space:pre;font-family:Consolas,'Courier New',monospace}}
.chart-container{{position:relative;width:100%;overflow-x:auto}}
.chart-container canvas{{display:block;width:100%;height:280px}}
.chart-legend{{display:flex;gap:16px;flex-wrap:wrap;margin-top:8px;font-size:.8em;color:#555}}
.chart-legend span{{display:flex;align-items:center;gap:4px}}
.chart-legend .dot{{width:10px;height:10px;border-radius:50%;display:inline-block}}
.footer{{text-align:center;color:#999;font-size:.75em;margin-top:16px;padding-top:10px;border-top:1px solid #e0e0e0}}
</style>
<script>
document.addEventListener('DOMContentLoaded',()=>{{
  document.querySelectorAll('.section-header').forEach(h=>{{
    h.addEventListener('click',()=>{{
      h.classList.toggle('collapsed');
      var body=h.nextElementSibling;
      if(body)body.style.display=h.classList.contains('collapsed')?'none':'';
    }});
  }});
}});
</script>
</head>
<body>
<div class=""container"">
<div class=""header"">
<h1>SPO Search Probe - Session Report</h1>
<span class=""tool-info"">{Program.AppVersion} - {Program.AppAuthor}</span>
</div>
<p class=""subtitle"">Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>

<div class=""section"">
<div class=""section-header"">Overview &amp; Statistics <span class=""arrow"">&#9660;</span></div>
<div class=""section-body"">
<div class=""stats"">
<div class=""stat""><div class=""val"">{totalQueries}</div><div class=""lbl"">Total Queries</div></div>
<div class=""stat""><div class=""val"">{uniqueUsers.Count}</div><div class=""lbl"">Users</div></div>
<div class=""stat{statCssSuccErr}""><div class=""val"">{httpOk} / {httpErr}</div><div class=""lbl"">Success / Errors</div></div>
<div class=""stat""><div class=""val"" style=""font-size:1em"">{probeStartTime}</div><div class=""lbl"">Probe Start</div></div>
<div class=""stat{statCssFound}""><div class=""val"" style=""font-size:1em"">{pageFoundTime}</div><div class=""lbl"">Page Found</div></div>
</div>
</div>
</div>

<div class=""section"">
<div class=""section-header"">Search Configuration <span class=""arrow"">&#9660;</span></div>
<div class=""section-body"">
<div class=""config-grid"">
<div class=""label"">Site URL:</div><div class=""value"">{Esc(siteUrl)}</div>
<div class=""label"">Tenant:</div><div class=""value"">{Esc(tenantId)}</div>
<div class=""label"">Query:</div><div class=""value"">{Esc(queryText)}</div>
<div class=""label"">Properties:</div><div class=""value"">{Esc(selectProps)}</div>
<div class=""label"">Row Limit:</div><div class=""value"">{Esc(rowLimit)}</div>
<div class=""label"">Sort:</div><div class=""value"">{Esc(sortList)}</div>
</div>
{pageUrlHtml}
</div>
</div>

<div class=""section"">
<div class=""section-header{timelineCollapsed}"">Execution Timeline <span class=""arrow"">&#9660;</span></div>
<div class=""section-body""{timelineHidden}>
<div class=""chart-container""><canvas id=""execChart""></canvas></div>
<div class=""chart-legend"">
<span><span class=""dot"" style=""background:#d83b01""></span> Page NOT FOUND in results</span>
<span><span class=""dot"" style=""background:#107c10""></span> Page FOUND in results</span>
{chartFoundMarkerLegend}
</div>
</div>
</div>

<div class=""section"">
<div class=""section-header collapsed"">Search Results (last {displayRows.Count} of {tsvRows.Count}) <span class=""arrow"">&#9660;</span></div>
<div class=""section-body"" style=""display:none"">
<table><tr><th>Time</th><th>User</th><th>HTTP</th><th>Validation</th><th>Results</th><th>Duration</th><th>Title</th><th>Path</th><th>WorkId</th></tr>
{resultRowsSb}</table>
</div>
</div>

<div class=""section"">
<div class=""section-header collapsed"">Log Package Contents <span class=""arrow"">&#9660;</span></div>
<div class=""section-body"" style=""display:none"">
<table><tr><th>File</th><th>Size</th><th>Type</th></tr>
{packageRowsSb}</table>
</div>
</div>

<div class=""section"">
<div class=""section-header collapsed"">Full Session Log <span class=""arrow"">&#9660;</span></div>
<div class=""section-body"" style=""display:none"">
<pre>{logContentSb}</pre>
</div>
</div>

{rawConfigSection}

<div class=""footer"">{Program.AppVersion} - {Program.AppAuthor}<br>Licensed under the MIT License. Use at your own risk. See <a href=""https://github.com/HolgerLutz/SPOSearchProbe/blob/main/LICENSE"" style=""color:#0078d4"">LICENSE</a> for details.</div>
</div>
<script>
// Per-user stacked sub-chart renderer (matches the live chart layout).
// Groups data by user, renders each user in their own vertically-stacked chart area.
(function(){{
  var data={chartPoints};
  var canvas=document.getElementById('execChart');
  if(!canvas||data.length===0)return;

  // Group data points by user
  var userMap={{}};
  var userOrder=[];
  for(var i=0;i<data.length;i++){{
    var u=data[i].u||'(unknown)';
    if(!userMap[u]){{userMap[u]=[];userOrder.push(u);}}
    userMap[u].push(data[i]);
  }}
  var numUsers=userOrder.length;
  var lineColors=['#0078d4','#9b59b6','#e67e22','#2ecc71','#e74c3c','#1abc9c','#f39c12','#3498db'];

  // Calculate canvas dimensions: 260px per user sub-chart
  var subH=260;
  var dpr=window.devicePixelRatio||1;
  var rect=canvas.getBoundingClientRect();
  var totalH=subH*numUsers;
  canvas.style.height=totalH+'px';
  var W=rect.width*dpr, H=totalH*dpr;
  canvas.width=W; canvas.height=H;
  ctx=canvas.getContext('2d');
  ctx.scale(dpr,dpr);
  var w=rect.width;

  // Store per-user geometry for tooltip hit-testing
  var allSeries=[];

  for(var ui=0;ui<numUsers;ui++){{
    var userName=userOrder[ui];
    var pts=userMap[userName];
    var lineColor=lineColors[ui%lineColors.length];
    var oY=ui*subH;
    var pad={{t:28,r:20,b:50,l:55}};
    var cw=w-pad.l-pad.r, ch=subH-pad.t-pad.b;

    // Y-axis scale
    var maxMs=0;
    for(var i=0;i<pts.length;i++) if(pts[i].ms>maxMs) maxMs=pts[i].ms;
    maxMs=Math.ceil(maxMs/100)*100||100;

    // Background fill for sub-chart area
    ctx.fillStyle=ui%2===0?'#fafafa':'#f4f4f4';
    ctx.fillRect(0,oY,w,subH);

    // User name label
    ctx.fillStyle=lineColor; ctx.font='bold 11px Segoe UI'; ctx.textAlign='left'; ctx.textBaseline='top';
    ctx.fillText(userName,pad.l+4,oY+6);

    // Grid lines
    ctx.strokeStyle='#e0e0e0'; ctx.lineWidth=1;
    var gridLines=4;
    for(var gi=0;gi<=gridLines;gi++){{
      var gy=oY+pad.t+ch-ch*(gi/gridLines);
      ctx.beginPath(); ctx.moveTo(pad.l,gy); ctx.lineTo(pad.l+cw,gy); ctx.stroke();
      ctx.fillStyle='#888'; ctx.font='10px Segoe UI'; ctx.textAlign='right'; ctx.textBaseline='middle';
      ctx.fillText(Math.round(maxMs*gi/gridLines)+'ms',pad.l-6,gy);
    }}

    // X-axis time labels
    ctx.fillStyle='#888'; ctx.font='9px Segoe UI'; ctx.textAlign='center'; ctx.textBaseline='top';
    var labelStep=Math.max(1,Math.floor(pts.length/8));
    for(var li=0;li<pts.length;li+=labelStep){{
      var lx=pad.l+(pts.length>1?li/(pts.length-1):0.5)*cw;
      var lbl=pts[li].t;
      if(lbl.length>=19) lbl=lbl.substring(11,19);
      else if(lbl.length>10) lbl=lbl.substring(11);
      ctx.save(); ctx.translate(lx,oY+pad.t+ch+6); ctx.rotate(Math.PI/6); ctx.fillText(lbl,0,0); ctx.restore();
    }}

    // Connecting line
    ctx.strokeStyle=lineColor; ctx.lineWidth=1.5; ctx.beginPath();
    for(var i=0;i<pts.length;i++){{
      var x=pad.l+(pts.length>1?i/(pts.length-1):0.5)*cw;
      var y=oY+pad.t+ch-ch*(pts[i].ms/maxMs);
      if(i===0)ctx.moveTo(x,y); else ctx.lineTo(x,y);
    }}
    ctx.stroke();

    // Data point dots
    var foundIdx=-1;
    for(var i=0;i<pts.length;i++){{
      var x=pad.l+(pts.length>1?i/(pts.length-1):0.5)*cw;
      var y=oY+pad.t+ch-ch*(pts[i].ms/maxMs);
      var col=pts[i].v==='FOUND'?'#107c10':pts[i].v==='NOT FOUND'?'#d83b01':lineColor;
      var r=pts[i].v?4:3;
      ctx.beginPath(); ctx.arc(x,y,r,0,Math.PI*2); ctx.fillStyle=col; ctx.fill();
      if(pts[i].v==='FOUND'&&foundIdx===-1) foundIdx=i;
    }}

    // PAGE FOUND marker (dashed vertical line + circle + label)
    if(foundIdx>=0){{
      var fx=pad.l+(pts.length>1?foundIdx/(pts.length-1):0.5)*cw;
      var fy=oY+pad.t+ch-ch*(pts[foundIdx].ms/maxMs);
      ctx.strokeStyle='#107c10'; ctx.lineWidth=2; ctx.setLineDash([4,3]);
      ctx.beginPath(); ctx.moveTo(fx,oY+pad.t); ctx.lineTo(fx,oY+pad.t+ch); ctx.stroke();
      ctx.setLineDash([]);
      ctx.beginPath(); ctx.arc(fx,fy,7,0,Math.PI*2); ctx.strokeStyle='#107c10'; ctx.lineWidth=2; ctx.stroke();
      ctx.fillStyle='#107c10'; ctx.font='bold 11px Segoe UI'; ctx.textAlign='center'; ctx.textBaseline='bottom';
      ctx.fillText('PAGE FOUND',fx,oY+pad.t-4);
    }}

    // Y-axis label
    ctx.fillStyle='#666'; ctx.font='10px Segoe UI';
    ctx.save(); ctx.translate(10,oY+pad.t+ch/2); ctx.rotate(-Math.PI/2); ctx.textAlign='center'; ctx.fillText('Execution Time',0,0); ctx.restore();

    // Store geometry for tooltip
    allSeries.push({{user:userName,pts:pts,oY:oY,pad:pad,cw:cw,ch:ch,maxMs:maxMs}});
  }}

  // Tooltip (shared across all sub-charts)
  canvas.style.cursor='crosshair';
  var tip=document.createElement('div');
  tip.style.cssText='position:absolute;background:#333;color:#fff;padding:4px 8px;border-radius:4px;font-size:12px;pointer-events:none;display:none;white-space:nowrap;z-index:10';
  canvas.parentNode.appendChild(tip);
  canvas.addEventListener('mousemove',function(e){{
    var br=canvas.getBoundingClientRect();
    var mx=e.clientX-br.left, my=e.clientY-br.top;
    var best=null, bestD=20;
    for(var si=0;si<allSeries.length;si++){{
      var s=allSeries[si];
      for(var i=0;i<s.pts.length;i++){{
        var px=s.pad.l+(s.pts.length>1?i/(s.pts.length-1):0.5)*s.cw;
        var py=s.oY+s.pad.t+s.ch-s.ch*(s.pts[i].ms/s.maxMs);
        var d=Math.sqrt((mx-px)*(mx-px)+(my-py)*(my-py));
        if(d<bestD){{bestD=d;best={{s:si,i:i,x:px,y:py}};}}
      }}
    }}
    if(best){{
      tip.style.display='block';
      tip.style.left=(e.clientX-br.left+12)+'px';
      tip.style.top=(e.clientY-br.top-10)+'px';
      var d=allSeries[best.s].pts[best.i];
      tip.innerHTML='#'+(best.i+1)+' '+d.t+'<br>'+d.ms+'ms - '+allSeries[best.s].user+(d.v?' ['+d.v+']':'');
    }}else{{tip.style.display='none';}}
  }});
  canvas.addEventListener('mouseleave',function(){{tip.style.display='none';}});
}})();
</script>
</body>
</html>");

        return sb.ToString();
    }

    /// <summary>
    /// Reads a file with shared access so it can be read even while the probe is
    /// actively writing to it. Standard File.ReadAllText() would fail with an
    /// IOException because the probe keeps the log files open for appending.
    /// FileShare.ReadWrite allows concurrent readers and writers.
    /// </summary>
    /// <param name="path">Full path to the file to read.</param>
    /// <returns>The complete file content as a UTF-8 string.</returns>
    private static string ReadFileShared(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var sr = new StreamReader(fs, Encoding.UTF8);
        return sr.ReadToEnd();
    }

    /// <summary>
    /// HTML-escapes a string to prevent XSS and rendering issues when embedding
    /// user-provided data (e.g. page URLs, query text) in the HTML report.
    /// </summary>
    private static string Esc(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
    }

    /// <summary>
    /// JavaScript-escapes a string for safe embedding in JS string literals.
    /// Handles backslashes, double quotes, and newlines that would break JS syntax.
    /// </summary>
    private static string EscJs(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "");
    }
}
