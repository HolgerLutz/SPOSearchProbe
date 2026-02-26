using System.Text.Json.Serialization;

namespace SPOSearchProbe;

/// <summary>
/// Strongly-typed model for the search-config.json configuration file.
/// Each property maps to a camelCase JSON key via <see cref="JsonPropertyNameAttribute"/>.
/// The admin distributes this file alongside the executable so that end-users
/// only need to click "Login" and "Start" without manual configuration.
/// </summary>
public class SearchConfig
{
    /// <summary>
    /// The SharePoint Online site collection URL to query against.
    /// Example: "https://contoso.sharepoint.com/sites/hr".
    /// Used to construct both the REST API endpoint and the OAuth2 resource/scope.
    /// </summary>
    [JsonPropertyName("siteUrl")]
    public string SiteUrl { get; set; } = "";

    /// <summary>
    /// Azure AD tenant ID (GUID) for the OAuth2 token endpoint.
    /// Example: "72f988bf-86f1-41af-91ab-2d7cd011db47".
    /// Required for building the login.microsoftonline.com authorize/token URLs.
    /// </summary>
    [JsonPropertyName("tenantId")]
    public string TenantId { get; set; } = "";

    /// <summary>
    /// The KQL (Keyword Query Language) search query to execute.
    /// Default: "contentclass:STS_ListItem" returns all list items.
    /// Admins can customize this to target specific content types or scopes.
    /// </summary>
    [JsonPropertyName("queryText")]
    public string QueryText { get; set; } = "contentclass:STS_ListItem";

    /// <summary>
    /// Managed properties to include in each search result row.
    /// These are passed as the 'selectproperties' parameter in the REST API call.
    /// Default set includes common diagnostic properties like WorkId for page validation.
    /// </summary>
    [JsonPropertyName("selectProperties")]
    public string[] SelectProperties { get; set; } = ["Title", "Path", "LastModifiedTime", "WorkId"];

    /// <summary>
    /// Maximum number of rows to return per search query (maps to REST API 'rowlimit').
    /// Keep this small for probing scenarios to minimize server load.
    /// </summary>
    [JsonPropertyName("rowLimit")]
    public int RowLimit { get; set; } = 10;

    /// <summary>
    /// Optional sort specification for the search results.
    /// Format: "PropertyName:direction" (e.g. "LastModifiedTime:descending").
    /// Empty string means default relevance ranking.
    /// </summary>
    [JsonPropertyName("sortList")]
    public string SortList { get; set; } = "";

    /// <summary>
    /// Optional URL of a specific SharePoint page to validate.
    /// When set, the UI shows a "Validate Page" button that checks whether
    /// the page is in the search index and then monitors subsequent queries
    /// for its WorkId. Used for freshness/crawl-lag investigations.
    /// </summary>
    [JsonPropertyName("pageUrl")]
    public string PageUrl { get; set; } = "";

    /// <summary>
    /// Numeric part of the polling interval (combined with <see cref="IntervalUnit"/>).
    /// Example: IntervalValue=10, IntervalUnit="seconds" → poll every 10 seconds.
    /// </summary>
    [JsonPropertyName("intervalValue")]
    public int IntervalValue { get; set; } = 10;

    /// <summary>
    /// Unit for the polling interval: "seconds", "minutes", or "hours".
    /// Combined with <see cref="IntervalValue"/> and converted to milliseconds
    /// by <see cref="GetIntervalMs"/>.
    /// </summary>
    [JsonPropertyName("intervalUnit")]
    public string IntervalUnit { get; set; } = "seconds";

    /// <summary>
    /// Azure AD application (client) ID used for the OAuth2 authorization code flow.
    /// Typically a first-party app like PnP Management Shell (9bc3ab49-...) that has
    /// delegated permissions for SharePoint search.
    /// </summary>
    [JsonPropertyName("clientId")]
    public string ClientId { get; set; } = "";

    /// <summary>
    /// Optional URL of a DTM (Data Transfer Manager) workspace where users can
    /// upload their collected log archives. When set, the "Collect Logs" action
    /// opens this URL in the browser after creating the ZIP.
    /// </summary>
    [JsonPropertyName("workspaceUrl")]
    public string WorkspaceUrl { get; set; } = "";

    /// <summary>
    /// Converts the human-readable interval (value + unit) into milliseconds
    /// suitable for <see cref="System.Windows.Forms.Timer.Interval"/>.
    /// Falls back to treating the unit as "seconds" for any unrecognized unit string,
    /// which is the safest default for a polling tool.
    /// </summary>
    public int GetIntervalMs()
    {
        return IntervalUnit.ToLowerInvariant() switch
        {
            "minutes" => IntervalValue * 60_000,
            "hours" => IntervalValue * 3_600_000,
            _ => IntervalValue * 1000 // default: treat as seconds
        };
    }
}
