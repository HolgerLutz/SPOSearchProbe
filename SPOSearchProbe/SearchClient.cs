using System.Diagnostics;
using System.IO.Compression;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Web;

namespace SPOSearchProbe;

/// <summary>
/// Encapsulates the results of a single SharePoint Online search REST API call.
/// Contains both the parsed result data and HTTP-level diagnostic information
/// (correlation IDs, timing) needed for troubleshooting search issues.
/// </summary>
public class SearchResult
{
    /// <summary>HTTP status code from the SharePoint REST API response (e.g. 200, 401, 500).</summary>
    public int StatusCode { get; set; }

    /// <summary>
    /// SharePoint correlation ID from the SPRequestGuid or request-id response header.
    /// This GUID is essential for support escalations — it allows Microsoft engineers
    /// to trace the request through SharePoint's internal telemetry systems.
    /// </summary>
    public string? CorrelationId { get; set; }

    /// <summary>
    /// Internal search request ID from the X-SearchInternalRequestId header.
    /// Specific to the SharePoint Search service (as opposed to the general SPRequestGuid),
    /// useful for diagnosing search-specific issues like index staleness or query routing.
    /// </summary>
    public string? InternalRequestId { get; set; }

    /// <summary>Total number of matching items in the search index (may exceed RowLimit).</summary>
    public int TotalRows { get; set; }

    /// <summary>Number of rows actually returned in this response (≤ RowLimit).</summary>
    public int RowCount { get; set; }

    /// <summary>Wall-clock time in milliseconds for the entire HTTP round-trip (measured client-side).</summary>
    public long ElapsedMs { get; set; }

    /// <summary>The full request URL including all query parameters (for diagnostic logging).</summary>
    public string RequestUrl { get; set; } = "";

    /// <summary>
    /// Raw QueryIdentityDiagnostics value from the search response properties.
    /// Contains the effective user identity used for search authorization and security
    /// trimming — critical for diagnosing "user can't see results" issues.
    /// </summary>
    public string? QueryIdentityDiagnostics { get; set; }

    /// <summary>
    /// Parsed result rows. Each row is a dictionary mapping managed property names
    /// (e.g. "Title", "Path", "WorkId") to their string values. The dictionary uses
    /// case-insensitive keys because SharePoint property names can vary in casing.
    /// </summary>
    public List<Dictionary<string, string>> Rows { get; set; } = [];
}

/// <summary>
/// Executes SharePoint Online search queries via the REST API (/_api/search/query)
/// and parses the OData verbose JSON response format. Also handles diagnostic logging
/// of raw request/response pairs to ZIP archives for offline analysis.
/// </summary>
public class SearchClient
{
    /// <summary>
    /// Shared HttpClient instance — reused across calls to benefit from connection pooling.
    /// HttpClient is designed to be long-lived and thread-safe.
    /// </summary>
    private readonly HttpClient _http = new();

    /// <summary>
    /// Executes a single search query against the SharePoint Online search REST API.
    /// </summary>
    /// <param name="siteUrl">SharePoint site URL (e.g. "https://contoso.sharepoint.com/sites/hr").</param>
    /// <param name="accessToken">A valid OAuth2 Bearer token with search permissions.</param>
    /// <param name="queryText">KQL query text (e.g. "contentclass:STS_ListItem").</param>
    /// <param name="selectProperties">Managed properties to return in each result row.</param>
    /// <param name="rowLimit">Maximum number of rows to return.</param>
    /// <param name="sortList">Optional sort specification (e.g. "LastModifiedTime:descending").</param>
    /// <param name="requestLogDir">Optional directory path where request/response ZIP logs are saved.</param>
    /// <param name="userLabel">Optional user label for log file naming (e.g. email address).</param>
    /// <param name="queryType">Optional query type label ("QUERY", "VALIDATE", "TEST") for log file naming.</param>
    /// <param name="ct">Cancellation token.</param>
    /// <returns>A <see cref="SearchResult"/> with parsed data and diagnostic headers.</returns>
    public async Task<SearchResult> ExecuteSearchAsync(
        string siteUrl, string accessToken, string queryText,
        string[] selectProperties, int rowLimit, string sortList,
        string? requestLogDir = null, string? userLabel = null,
        string? queryType = null,
        CancellationToken ct = default)
    {
        // --- Build the REST API URL ---
        // The SharePoint search REST API expects single-quoted string parameters in the URL.
        // Single quotes within the query text must be escaped as '' (double single-quote).
        var searchUrl = $"{siteUrl.TrimEnd('/')}/_api/search/query";
        var props = string.Join(",", selectProperties);
        var escaped = queryText.Replace("'", "''");
        var q = HttpUtility.UrlEncode(escaped);
        // trimduplicates=false ensures all matching items are returned (duplicates are
        // collapsed by default, which can hide results during validation scenarios).
        // QueryIdentityDiagnostics=true asks the search service to include the effective
        // user identity in the response metadata — critical for security trimming debugging.
        var url = $"{searchUrl}?querytext='{q}'&selectproperties='{props}'&rowlimit={rowLimit}" +
                  "&trimduplicates=false&Properties='QueryIdentityDiagnostics:true'";
        if (!string.IsNullOrEmpty(sortList))
            url += $"&sortlist='{sortList}'";

        // --- Construct the HTTP request ---
        var request = new HttpRequestMessage(HttpMethod.Get, url);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        // Request the OData verbose JSON format (application/json;odata=verbose).
        // This legacy format wraps everything in a "d" envelope with "__metadata" objects.
        // The verbose format is used because SharePoint's search endpoint returns a deeply
        // nested structure (d → query → PrimaryQueryResult → RelevantResults → Table → Rows)
        // that is well-documented in the SharePoint REST API reference.
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        request.Headers.Accept.First().Parameters.Add(new NameValueHeaderValue("odata", "verbose"));

        // --- Execute and time the request ---
        var sw = Stopwatch.StartNew();
        var response = await _http.SendAsync(request, ct);
        sw.Stop();

        var content = await response.Content.ReadAsStringAsync(ct);
        var result = new SearchResult
        {
            StatusCode = (int)response.StatusCode,
            ElapsedMs = sw.ElapsedMilliseconds,
            RequestUrl = url
        };

        // --- Extract diagnostic response headers ---
        // X-SearchInternalRequestId: unique ID within the search service for this query
        if (response.Headers.TryGetValues("X-SearchInternalRequestId", out var reqIds))
            result.InternalRequestId = reqIds.FirstOrDefault();
        // SPRequestGuid: SharePoint's correlation ID for end-to-end tracing.
        // Falls back to the generic "request-id" header if SPRequestGuid is absent
        // (can happen with certain proxy/gateway configurations).
        if (response.Headers.TryGetValues("SPRequestGuid", out var corrIds))
            result.CorrelationId = corrIds.FirstOrDefault();
        else if (response.Headers.TryGetValues("request-id", out var rids))
            result.CorrelationId = rids.FirstOrDefault();

        // --- Log request/response to a ZIP archive for offline analysis ---
        // Each query is saved as a ZIP containing the raw request and the formatted
        // JSON response. The bearer token is redacted to prevent credential leakage.
        // File naming: {timestamp}_{user}_{queryType}.zip
        if (!string.IsNullOrEmpty(requestLogDir))
        {
            try
            {
                Directory.CreateDirectory(requestLogDir);
                var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                // Sanitize user label and query type for use as file name components
                var safeUser = string.IsNullOrEmpty(userLabel) ? "" : $"_{System.Text.RegularExpressions.Regex.Replace(userLabel, "[^a-zA-Z0-9]", "_")}";
                var safeQT = string.IsNullOrEmpty(queryType) ? "" : $"_{queryType.Replace(" ", "_")}";
                var baseName = $"{ts}{safeUser}{safeQT}";
                var zipPath = Path.Combine(requestLogDir, $"{baseName}.zip");

                // Build a human-readable request log (with the token redacted for security)
                var reqContent = $"GET {url}\r\nAccept: application/json;odata=verbose\r\nAuthorization: Bearer <redacted>\r\n";
                // Capture all response headers for diagnostic purposes
                var respHeaders = new StringBuilder();
                foreach (var h in response.Headers)
                    respHeaders.AppendLine($"{h.Key}: {string.Join(", ", h.Value)}");
                // Pretty-print the JSON body for readability
                var formattedBody = FormatJson(content);
                var respContent = $"HTTP/1.1 {(int)response.StatusCode} {response.ReasonPhrase}\r\n{respHeaders}\r\n{formattedBody}";

                // Write both files into a single ZIP archive (keeps the log directory tidy)
                using var ms = new MemoryStream();
                using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    var reqEntry = archive.CreateEntry($"{baseName}_request.txt", CompressionLevel.Optimal);
                    using (var writer = new StreamWriter(reqEntry.Open()))
                        writer.Write(reqContent);
                    var respEntry = archive.CreateEntry($"{baseName}_response.json", CompressionLevel.Optimal);
                    using (var writer = new StreamWriter(respEntry.Open()))
                        writer.Write(respContent);
                }
                await File.WriteAllBytesAsync(zipPath, ms.ToArray(), ct);
            }
            catch { } // Logging must never break the actual search flow
        }

        // --- Parse the OData verbose JSON response ---
        // The SharePoint search REST API returns a deeply nested JSON structure:
        //   { "d": { "query": { "PrimaryQueryResult": { "RelevantResults": {
        //       "TotalRows": N, "RowCount": N,
        //       "Table": { "Rows": { "results": [
        //           { "Cells": { "results": [ { "Key": "Title", "Value": "..." }, ... ] } }
        //       ]}}
        //   }}}}}
        // The "results" arrays are an OData verbose convention (collections are wrapped).
        if (response.IsSuccessStatusCode)
        {
            try
            {
                using var doc = JsonDocument.Parse(content);
                // Navigate the nested OData verbose structure to the RelevantResults
                var relevant = doc.RootElement
                    .GetProperty("d").GetProperty("query")
                    .GetProperty("PrimaryQueryResult").GetProperty("RelevantResults");

                result.TotalRows = relevant.GetProperty("TotalRows").GetInt32();
                result.RowCount = relevant.GetProperty("RowCount").GetInt32();

                // --- Extract QueryIdentityDiagnostics from the query-level properties ---
                // This is in a separate "Properties" collection at the query level (not per-row).
                // It contains the effective user identity used for security trimming.
                try
                {
                    var props2 = doc.RootElement.GetProperty("d").GetProperty("query")
                        .GetProperty("Properties").GetProperty("results");
                    foreach (var p in props2.EnumerateArray())
                    {
                        if (p.GetProperty("Key").GetString() == "QueryIdentityDiagnostics")
                        {
                            result.QueryIdentityDiagnostics = p.GetProperty("Value").GetString();
                            break;
                        }
                    }
                }
                catch { } // Property may not exist in all response types

                // --- Parse individual result rows ---
                // Each row contains a "Cells" array of Key/Value pairs (not a flat object).
                // This is the OData verbose representation of a DataTable-like structure.
                var rows = relevant.GetProperty("Table").GetProperty("Rows").GetProperty("results");
                foreach (var row in rows.EnumerateArray())
                {
                    var cells = row.GetProperty("Cells").GetProperty("results");
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var cell in cells.EnumerateArray())
                    {
                        var key = cell.GetProperty("Key").GetString() ?? "";
                        var val = cell.GetProperty("Value").GetString() ?? "";
                        // Some values come wrapped in curly braces (e.g. GUIDs) — strip them
                        // for cleaner display and easier comparison.
                        if (val.StartsWith('{') && val.EndsWith('}'))
                            val = val.TrimStart('{').TrimEnd('}');
                        dict[key] = val;
                    }
                    result.Rows.Add(dict);
                }
            }
            catch { } // Parse failures shouldn't crash — we still return the HTTP status/timing
        }

        return result;
    }

    /// <summary>
    /// Pretty-prints a JSON string with indentation for readability in log files.
    /// Returns the original string unchanged if parsing fails (e.g. non-JSON error responses).
    /// </summary>
    private static string FormatJson(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            return JsonSerializer.Serialize(doc, new JsonSerializerOptions { WriteIndented = true });
        }
        catch { return json; }
    }
}
