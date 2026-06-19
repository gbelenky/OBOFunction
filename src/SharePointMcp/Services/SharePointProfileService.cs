using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;
using Microsoft.Graph;
using SharePointMcp.Models;

namespace SharePointMcp.Services;

/// <summary>
/// Fetches a merged Microsoft Graph (<c>/me</c>) + SharePoint UPS
/// (PeopleManager/GetMyProperties) profile using the supplied <see cref="TokenCredential"/>.
/// The credential is chosen per-request by <see cref="RequestCredentialProvider"/>:
/// a user-delegated token (passthrough) or DefaultAzureCredential (local/app).
/// </summary>
public sealed class SharePointProfileService
{
    private static readonly HttpClient Http = new();

    private readonly TokenCredential _credential;
    private readonly string? _sharePointRootSiteUrl;
    private readonly bool _allowSharePointUps;
    private readonly string _resolvedVia;
    private readonly string[] _graphScopes;

    /// <param name="credential">The identity to act as for this request.</param>
    /// <param name="sharePointRootSiteUrl">SharePoint root site (e.g. https://contoso.sharepoint.com) or null to skip UPS.</param>
    /// <param name="allowSharePointUps">
    /// When false, the SharePoint UPS call is skipped entirely. Set false when the credential is a
    /// fixed-audience user token (it cannot mint a SharePoint-audience token), avoiding a guaranteed 401.
    /// </param>
    /// <param name="resolvedVia">Diagnostic marker recorded on the result ("user" or "app").</param>
    public SharePointProfileService(
        TokenCredential credential,
        string? sharePointRootSiteUrl,
        bool allowSharePointUps,
        string resolvedVia)
    {
        _credential = credential;
        _sharePointRootSiteUrl = sharePointRootSiteUrl?.TrimEnd('/');
        _allowSharePointUps = allowSharePointUps;
        _resolvedVia = resolvedVia;
        _graphScopes = ["https://graph.microsoft.com/.default"];
    }

    public async Task<UserProfile> GetMyProfileAsync(CancellationToken ct = default)
    {
        var graph = new GraphServiceClient(_credential, _graphScopes);

        var me = await graph.Me.GetAsync(rc =>
        {
            rc.QueryParameters.Select =
            [
                "id", "displayName", "givenName", "surname", "userPrincipalName",
                "mail", "jobTitle", "department", "officeLocation",
                "preferredLanguage", "mobilePhone", "businessPhones"
            ];
        }, ct).ConfigureAwait(false);

        SharePointProfile? sp = null;
        if (_allowSharePointUps && !string.IsNullOrWhiteSpace(_sharePointRootSiteUrl))
        {
            try
            {
                sp = await FetchSharePointProfileAsync(ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // SharePoint UPS is best-effort; the Graph profile is still returned.
                // Surface the reason to stderr so dev-loop failures (e.g. the Azure CLI
                // app's SharePoint-token 401, since it has no SP delegated consent) are
                // visible instead of silently producing a null SharePointProfile.
                Console.Error.WriteLine(
                    $"[SharePointProfileService] UPS fetch failed (resolvedVia={_resolvedVia}): " +
                    $"{ex.GetType().Name}: {ex.Message}");
                sp = null;
            }
        }

        return new UserProfile
        {
            Id = me?.Id,
            DisplayName = me?.DisplayName,
            GivenName = me?.GivenName,
            Surname = me?.Surname,
            UserPrincipalName = me?.UserPrincipalName,
            Mail = me?.Mail,
            JobTitle = me?.JobTitle,
            Department = me?.Department,
            OfficeLocation = me?.OfficeLocation,
            PreferredLanguage = me?.PreferredLanguage,
            MobilePhone = me?.MobilePhone,
            BusinessPhones = me?.BusinessPhones?.ToArray() ?? [],
            SharePointProfile = sp,
            ResolvedVia = _resolvedVia
        };
    }

    private async Task<SharePointProfile> FetchSharePointProfileAsync(CancellationToken ct)
    {
        var spResource = new Uri(_sharePointRootSiteUrl!).GetLeftPart(UriPartial.Authority);
        var token = await _credential
            .GetTokenAsync(new TokenRequestContext([$"{spResource}/.default"]), ct)
            .ConfigureAwait(false);

        var url = $"{_sharePointRootSiteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        req.Headers.Add("Accept", "application/json;odata=nometadata");

        using var resp = await Http.SendAsync(req, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();

        await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
        var root = doc.RootElement;

        string Str(string p) => root.TryGetProperty(p, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";
        string? StrOrNull(string p) => root.TryGetProperty(p, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

        IReadOnlyList<string> MultiValue(string propName)
        {
            if (!root.TryGetProperty("UserProfileProperties", out var arr) || arr.ValueKind != JsonValueKind.Array)
                return [];

            foreach (var kv in arr.EnumerateArray())
            {
                if (kv.TryGetProperty("Key", out var k) && string.Equals(k.GetString(), propName, StringComparison.OrdinalIgnoreCase))
                {
                    if (kv.TryGetProperty("Value", out var v) && v.ValueKind == JsonValueKind.String)
                    {
                        var raw = v.GetString();
                        if (string.IsNullOrWhiteSpace(raw)) return [];
                        return raw.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                    }
                }
            }
            return [];
        }

        string? aboutMe = StrOrNull("AboutMe");

        // Capture every UPS property by internal name so custom attributes
        // (created in Manage User Properties) surface without hard-coding each one.
        var extended = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (root.TryGetProperty("UserProfileProperties", out var allProps) && allProps.ValueKind == JsonValueKind.Array)
        {
            foreach (var kv in allProps.EnumerateArray())
            {
                if (kv.TryGetProperty("Key", out var k) && kv.TryGetProperty("Value", out var v))
                {
                    var key = k.GetString();
                    var val = v.ValueKind == JsonValueKind.String ? v.GetString() : v.ToString();
                    if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(val))
                        extended[key] = val;
                }
            }
        }

        aboutMe ??= extended.GetValueOrDefault("AboutMe");

        return new SharePointProfile
        {
            AccountName = Str("AccountName"),
            AboutMe = aboutMe,
            PersonalUrl = StrOrNull("PersonalUrl"),
            Skills = MultiValue("SPS-Skills"),
            Interests = MultiValue("SPS-Interests"),
            PastProjects = MultiValue("SPS-PastProjects"),
            Responsibilities = MultiValue("SPS-Responsibility"),
            Schools = MultiValue("SPS-School"),
            ExtendedProperties = extended
        };
    }
}
