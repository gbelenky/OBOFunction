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
                "displayName", "givenName", "surname", "mail", "userPrincipalName", "jobTitle"
            ];
        }, ct).ConfigureAwait(false);

        UpsFields ups = default;
        if (_allowSharePointUps && !string.IsNullOrWhiteSpace(_sharePointRootSiteUrl))
        {
            try
            {
                ups = await FetchSharePointProfileAsync(ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // SharePoint UPS is best-effort; the Graph name/email are still returned.
                // Surface the reason to stderr so dev-loop failures (e.g. a SharePoint-token
                // 401) are visible instead of silently dropping the UPS fields.
                Console.Error.WriteLine(
                    $"[SharePointProfileService] UPS fetch failed (resolvedVia={_resolvedVia}): " +
                    $"{ex.GetType().Name}: {ex.Message}");
                ups = default;
            }
        }

        return new UserProfile
        {
            Name = me?.DisplayName,
            FirstName = me?.GivenName,
            LastName = me?.Surname,
            Email = me?.Mail ?? me?.UserPrincipalName,
            JobTitle = ups.JobTitle ?? me?.JobTitle,
            Responsibilities = ups.Responsibilities ?? [],
            PastProjects = ups.PastProjects ?? [],
            Interests = ups.Interests ?? [],
            Country = ups.Country
        };
    }

    /// <summary>The subset of SharePoint UPS properties the MCP tool exposes.</summary>
    private readonly record struct UpsFields(
        string? JobTitle,
        IReadOnlyList<string>? Responsibilities,
        IReadOnlyList<string>? PastProjects,
        IReadOnlyList<string>? Interests,
        string? Country);

    private async Task<UpsFields> FetchSharePointProfileAsync(CancellationToken ct)
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

        // Index every UPS property by internal name (Key/Value pairs) so the named fields below
        // — including the custom IntranetCountry attribute — resolve regardless of ordering.
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (root.TryGetProperty("UserProfileProperties", out var arr) && arr.ValueKind == JsonValueKind.Array)
        {
            foreach (var kv in arr.EnumerateArray())
            {
                if (kv.TryGetProperty("Key", out var k) && kv.TryGetProperty("Value", out var v))
                {
                    var key = k.GetString();
                    var val = v.ValueKind == JsonValueKind.String ? v.GetString() : v.ToString();
                    if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(val))
                        props[key] = val!;
                }
            }
        }

        string? Single(string name) =>
            props.TryGetValue(name, out var s) && !string.IsNullOrWhiteSpace(s) ? s : null;

        IReadOnlyList<string> Multi(string name)
        {
            var raw = Single(name);
            return string.IsNullOrWhiteSpace(raw)
                ? []
                : raw.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }

        return new UpsFields(
            JobTitle: Single("SPS-JobTitle"),
            Responsibilities: Multi("SPS-Responsibility"),
            PastProjects: Multi("SPS-PastProjects"),
            Interests: Multi("SPS-Interests"),
            Country: Single("IntranetCountry"));
    }
}
