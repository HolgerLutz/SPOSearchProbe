using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Web;

namespace SPOSearchProbe;

/// <summary>
/// Implements the OAuth 2.0 Authorization Code flow with PKCE (Proof Key for Code Exchange)
/// for Azure AD / Entra ID, targeting SharePoint Online delegated permissions.
///
/// Flow overview:
/// 1. Generate a PKCE code verifier + challenge (random 32 bytes → base64url).
/// 2. Start a temporary local HTTP listener on an available port (18700–18799).
/// 3. Open the user's default browser to the Azure AD /authorize endpoint.
/// 4. Azure AD authenticates the user and redirects back to http://localhost:{port}/
///    with an authorization code in the query string.
/// 5. The local listener captures the code, shows a "success" page, then shuts down.
/// 6. Exchange the authorization code + PKCE verifier for access + refresh tokens
///    via the /token endpoint.
///
/// Also provides silent token refresh using the cached refresh token.
/// </summary>
public static class OAuthHelper
{
    /// <summary>
    /// Generates a PKCE (RFC 7636) code verifier and its SHA-256 challenge.
    ///
    /// PKCE prevents authorization code interception attacks. The verifier is a
    /// cryptographically random string sent only to the /token endpoint (never to
    /// the browser). The challenge (SHA-256 hash of the verifier) is sent to /authorize.
    /// Azure AD verifies that the entity exchanging the code is the same one that
    /// started the flow by checking SHA256(verifier) == challenge.
    ///
    /// Both values are base64url-encoded (RFC 4648 §5): '+' → '-', '/' → '_', no padding.
    /// </summary>
    private static (string Verifier, string Challenge) NewPkceChallenge()
    {
        // 32 random bytes → 43-character base64url string (well within the 43–128 range required by RFC 7636)
        var buf = RandomNumberGenerator.GetBytes(32);
        var verifier = Convert.ToBase64String(buf)
            .Replace('+', '-').Replace('/', '_').TrimEnd('=');
        // The challenge is the base64url-encoded SHA-256 hash of the ASCII verifier
        var hash = SHA256.HashData(Encoding.ASCII.GetBytes(verifier));
        var challenge = Convert.ToBase64String(hash)
            .Replace('+', '-').Replace('/', '_').TrimEnd('=');
        return (verifier, challenge);
    }

    /// <summary>
    /// Performs an interactive browser-based login using the OAuth2 Authorization Code + PKCE flow.
    /// Opens the user's default browser to Azure AD, captures the redirect on a local HTTP listener,
    /// and exchanges the authorization code for access + refresh tokens.
    /// </summary>
    /// <param name="clientId">Azure AD app registration client ID.</param>
    /// <param name="tenantId">Azure AD tenant ID (GUID).</param>
    /// <param name="siteUrl">SharePoint site URL — used to derive the OAuth2 resource scope.</param>
    /// <param name="loginHint">Optional UPN/email to pre-fill the Azure AD login form.</param>
    /// <param name="timeoutSeconds">How long to wait for the browser redirect before timing out.</param>
    /// <param name="ct">Cancellation token for cooperative cancellation.</param>
    /// <returns>A <see cref="TokenData"/> containing the access token, refresh token, and metadata.</returns>
    public static async Task<TokenData> InteractiveLoginAsync(
        string clientId, string tenantId, string siteUrl, string? loginHint = null,
        int timeoutSeconds = 120, CancellationToken ct = default)
    {
        // Derive the OAuth2 resource/scope from the SharePoint site URL.
        // For SPO, the scope is "{authority}/.default" which requests all delegated
        // permissions configured on the app registration. "offline_access" requests
        // a refresh token, "openid" requests an ID token (standard OIDC).
        var spUri = new Uri(siteUrl);
        var resource = $"{spUri.Scheme}://{spUri.Authority}";
        var scope = $"{resource.TrimEnd('/')}/.default offline_access openid";
        var pkce = NewPkceChallenge();

        // --- Step 1: Start a temporary local HTTP listener ---
        // We scan ports 18700–18799 to find one that's not in use. This range is
        // chosen to avoid conflicts with common services while being above the
        // ephemeral port range on most systems. The listener captures Azure AD's
        // redirect after the user authenticates in the browser.
        HttpListener? listener = null;
        int port = 0;
        for (int p = 18700; p <= 18799; p++)
        {
            try
            {
                var l = new HttpListener();
                l.Prefixes.Add($"http://localhost:{p}/");
                l.Start();
                listener = l;
                port = p;
                break;
            }
            catch { } // Port in use or permission denied — try the next one
        }
        if (listener == null || port == 0)
            throw new InvalidOperationException("Could not find an available port for OAuth2 listener.");

        // --- Step 2: Build the authorization URL and open the browser ---
        var redirectUri = $"http://localhost:{port}/";
        var authUrl = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/authorize" +
            $"?client_id={clientId}&response_type=code&redirect_uri={HttpUtility.UrlEncode(redirectUri)}" +
            $"&scope={HttpUtility.UrlEncode(scope)}&code_challenge={pkce.Challenge}" +
            $"&code_challenge_method=S256&prompt=select_account";
        // login_hint pre-fills the email field so the user doesn't have to type it again
        if (!string.IsNullOrEmpty(loginHint))
            authUrl += $"&login_hint={HttpUtility.UrlEncode(loginHint)}";

        // UseShellExecute = true opens the URL in the default browser
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(authUrl)
        { UseShellExecute = true });

        // --- Step 3: Wait for Azure AD to redirect back with the authorization code ---
        string? code = null;
        string? error = null;
        try
        {
            // Create a linked cancellation token that also fires after the timeout.
            // This prevents the app from hanging forever if the user closes the browser
            // or the redirect never arrives.
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));

            // GetContextAsync blocks until a single HTTP request arrives on our listener
            var context = await listener.GetContextAsync().WaitAsync(cts.Token);

            // Azure AD appends "?code=..." or "?error=...&error_description=..." to the redirect URI
            var query = HttpUtility.ParseQueryString(context.Request.Url!.Query);
            code = query["code"];
            error = query["error_description"];

            // Show a user-friendly HTML page in the browser so the user knows they can close the tab
            var html = code != null
                ? "<html><body><h2>Login successful!</h2><p>You can close this tab.</p></body></html>"
                : $"<html><body><h2>Login failed</h2><p>{error}</p></body></html>";
            var buf = Encoding.UTF8.GetBytes(html);
            context.Response.ContentLength64 = buf.Length;
            context.Response.ContentType = "text/html";
            await context.Response.OutputStream.WriteAsync(buf, ct);
            context.Response.OutputStream.Close();
        }
        finally
        {
            // Always clean up the listener — we only need it for the single redirect
            listener.Stop();
            listener.Close();
        }

        if (string.IsNullOrEmpty(code))
            throw new InvalidOperationException($"Authorization failed: {error}");

        // --- Step 4: Exchange the authorization code for access + refresh tokens ---
        // This POST to the /token endpoint proves we have the PKCE verifier that
        // matches the challenge sent in step 2. Without the verifier, a stolen
        // authorization code is useless.
        using var http = new HttpClient();
        var body = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["client_id"] = clientId,
            ["grant_type"] = "authorization_code",
            ["code"] = code,
            ["redirect_uri"] = redirectUri,
            ["code_verifier"] = pkce.Verifier, // PKCE proof — must match the challenge
            ["scope"] = scope
        });

        var resp = await http.PostAsync(
            $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token", body, ct);
        resp.EnsureSuccessStatusCode();
        var json = await resp.Content.ReadAsStringAsync(ct);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        // Calculate the absolute expiry time from the relative "expires_in" (seconds)
        var expiresIn = root.GetProperty("expires_in").GetInt32();
        return new TokenData
        {
            AccessToken = root.GetProperty("access_token").GetString()!,
            RefreshToken = root.GetProperty("refresh_token").GetString()!,
            ExpiresOn = DateTime.Now.AddSeconds(expiresIn).ToString("o"), // ISO 8601 round-trip format
            Scope = scope,
            TenantId = tenantId,
            ClientId = clientId
        };
    }

    /// <summary>
    /// Returns a valid access token from the cache, refreshing it silently if needed.
    ///
    /// Token freshness logic:
    /// - If the cached token expires more than 5 minutes from now → return it as-is.
    /// - If the token expires within 5 minutes (or has already expired) → use the
    ///   refresh token to get a new access token from Azure AD.
    /// - If no refresh token is available → return null (caller must re-authenticate).
    ///
    /// The 5-minute buffer ensures we refresh proactively before expiry, avoiding
    /// race conditions where a token expires mid-request.
    /// </summary>
    /// <param name="cacheFile">Path to the DPAPI-encrypted token cache file.</param>
    /// <param name="siteUrl">SharePoint site URL — needed to reconstruct the OAuth2 scope for the refresh request.</param>
    /// <param name="ct">Cancellation token.</param>
    /// <returns>A valid access token string, or null if no token is available and interactive login is required.</returns>
    public static async Task<string?> GetCachedOrRefreshedTokenAsync(
        string cacheFile, string siteUrl, CancellationToken ct = default)
    {
        var cache = TokenCache.Load(cacheFile);
        if (cache == null) return null; // No cache file or decryption failed

        var expiresOn = DateTime.Parse(cache.ExpiresOn);
        // 5-minute buffer: return the cached token if it's still valid with margin
        if (expiresOn > DateTime.Now.AddMinutes(5))
            return cache.AccessToken;

        // Token is expired (or about to expire) — try silent refresh
        if (string.IsNullOrEmpty(cache.RefreshToken)) return null;

        // Reconstruct the scope from the site URL (same logic as in InteractiveLoginAsync)
        var spUri = new Uri(siteUrl);
        var resource = $"{spUri.Scheme}://{spUri.Authority}";
        var scope = $"{resource.TrimEnd('/')}/.default offline_access openid";

        // Exchange the refresh token for a new access + refresh token pair.
        // Azure AD rotates refresh tokens by default — each refresh response
        // includes a new refresh token that replaces the old one.
        using var http = new HttpClient();
        var body = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["client_id"] = cache.ClientId,
            ["grant_type"] = "refresh_token",
            ["refresh_token"] = cache.RefreshToken,
            ["scope"] = scope
        });

        var resp = await http.PostAsync(
            $"https://login.microsoftonline.com/{cache.TenantId}/oauth2/v2.0/token", body, ct);
        resp.EnsureSuccessStatusCode();
        var json = await resp.Content.ReadAsStringAsync(ct);
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        var expiresIn = root.GetProperty("expires_in").GetInt32();
        var newData = new TokenData
        {
            AccessToken = root.GetProperty("access_token").GetString()!,
            RefreshToken = root.GetProperty("refresh_token").GetString()!,
            ExpiresOn = DateTime.Now.AddSeconds(expiresIn).ToString("o"),
            Scope = scope,
            TenantId = cache.TenantId,
            ClientId = cache.ClientId
        };
        // Persist the new token data so the next call can use the fresh tokens
        TokenCache.Save(cacheFile, newData);
        return newData.AccessToken;
    }
}
