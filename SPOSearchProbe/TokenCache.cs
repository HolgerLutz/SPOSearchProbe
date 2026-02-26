using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace SPOSearchProbe;

/// <summary>
/// Holds all OAuth2 token data that needs to be persisted between sessions.
/// Serialized to JSON before being encrypted and written to disk by <see cref="TokenCache"/>.
/// Contains both the short-lived access token and the long-lived refresh token,
/// plus metadata needed to perform silent token refresh without re-prompting the user.
/// </summary>
public class TokenData
{
    /// <summary>The Bearer access token used in Authorization headers for SharePoint REST API calls.</summary>
    public string AccessToken { get; set; } = "";

    /// <summary>
    /// The OAuth2 refresh token. Used to obtain a new access token silently after expiry.
    /// Refresh tokens are long-lived (typically 90 days for Azure AD) and are rotated
    /// on each use (the token endpoint returns a new refresh token with each refresh).
    /// </summary>
    public string RefreshToken { get; set; } = "";

    /// <summary>
    /// ISO 8601 timestamp ("o" format) indicating when the access token expires.
    /// Compared against DateTime.Now with a 5-minute buffer in <see cref="OAuthHelper"/>
    /// to trigger proactive refresh before the token actually expires.
    /// </summary>
    public string ExpiresOn { get; set; } = "";

    /// <summary>The OAuth2 scope string used when the token was issued (e.g. "https://contoso.sharepoint.com/.default offline_access openid").</summary>
    public string Scope { get; set; } = "";

    /// <summary>Azure AD tenant ID — needed for constructing the token endpoint URL during refresh.</summary>
    public string TenantId { get; set; } = "";

    /// <summary>Azure AD client (application) ID — needed for the refresh token grant request.</summary>
    public string ClientId { get; set; } = "";
}

/// <summary>
/// Persists <see cref="TokenData"/> to disk using Windows DPAPI (Data Protection API)
/// encryption scoped to the current Windows user account.
///
/// Design rationale:
/// - DPAPI with <see cref="DataProtectionScope.CurrentUser"/> ensures that only the same
///   Windows user on the same machine can decrypt the token file. No master key or
///   password is needed — Windows derives the encryption key from the user's credentials.
/// - The file is opaque binary (not readable JSON) so tokens aren't accidentally exposed
///   if the file is copied or shared.
/// - Each user gets their own cache file (named by email), enabling multi-user scenarios
///   on the same machine without cross-user token leakage.
/// </summary>
public static class TokenCache
{
    /// <summary>
    /// Serializes <paramref name="data"/> to JSON, encrypts it with DPAPI, and writes
    /// the encrypted bytes to the specified file path. Overwrites any existing file.
    /// </summary>
    /// <param name="path">Absolute path to the token cache file (e.g. .token-user-john_doe.dat).</param>
    /// <param name="data">The token data to persist.</param>
    public static void Save(string path, TokenData data)
    {
        var json = JsonSerializer.Serialize(data);
        var bytes = Encoding.UTF8.GetBytes(json);
        // Encrypt with DPAPI scoped to the current Windows user.
        // The 'null' entropy parameter means no additional secret is mixed in —
        // the user's Windows login credentials alone protect the data.
        var encrypted = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
        File.WriteAllBytes(path, encrypted);
    }

    /// <summary>
    /// Loads and decrypts a previously saved token cache file.
    ///
    /// Returns null silently in several failure scenarios by design:
    /// - File does not exist (user hasn't logged in yet).
    /// - DPAPI decryption fails (file was created by a different user, or is corrupted).
    /// - JSON deserialization fails (file format changed between versions).
    ///
    /// Returning null rather than throwing allows callers to simply fall through
    /// to an interactive login prompt without needing try/catch everywhere.
    /// </summary>
    /// <param name="path">Absolute path to the token cache file.</param>
    /// <returns>The deserialized token data, or null if the file is missing/unreadable.</returns>
    public static TokenData? Load(string path)
    {
        if (!File.Exists(path)) return null;
        try
        {
            var encrypted = File.ReadAllBytes(path);
            // Decrypt with DPAPI — will throw CryptographicException if the file
            // was encrypted by a different Windows user or is corrupt.
            var bytes = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
            var json = Encoding.UTF8.GetString(bytes);
            return JsonSerializer.Deserialize<TokenData>(json);
        }
        catch
        {
            // Swallow all exceptions: corrupt file, wrong user, format change, etc.
            // The caller will treat null as "no cached token available".
            return null;
        }
    }
}
