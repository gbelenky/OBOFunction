using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;
using Microsoft.Graph;
using SharePointAgent.Models;

namespace SharePointAgent.Services;

/// <summary>
/// Local-first profile fetch. Uses the developer/runtime <see cref="TokenCredential"/>
/// (DefaultAzureCredential: <c>az login</c> / VS Code / managed identity) to read the
/// signed-in user's Microsoft Graph <c>/me</c> and the SharePoint User Profile Service
/// (PeopleManager/GetMyProperties), then merges them into a <see cref="UserProfile"/>.
/// </summary>
/// <remarks>
/// There is no OAuth On-Behalf-Of here: a hosted agent has no inbound user assertion to
/// exchange. Locally this returns the developer's own profile, which is the right behavior
/// for build/debug. In production the agent identity (managed identity) is used.
/// </remarks>
public sealed class SharePointProfileService : ISharePointProfileService
{
    private static readonly HttpClient Http = new();

    private readonly TokenCredential _credential;
    private readonly string? _sharePointRootSiteUrl;
    private readonly string[] _graphScopes;

    public SharePointProfileService(TokenCredential credential, string? sharePointRootSiteUrl)
    {
        _credential = credential;
        _sharePointRootSiteUrl = sharePointRootSiteUrl?.TrimEnd('/');
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
        if (!string.IsNullOrWhiteSpace(_sharePointRootSiteUrl))
        {
            try
            {
                sp = await FetchSharePointProfileAsync(ct).ConfigureAwait(false);
            }
            catch
            {
                // SharePoint UPS is best-effort locally; Graph profile is still returned.
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
            SharePointProfile = sp
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
        if (aboutMe is null && root.TryGetProperty("UserProfileProperties", out var props) && props.ValueKind == JsonValueKind.Array)
        {
            foreach (var kv in props.EnumerateArray())
            {
                if (kv.TryGetProperty("Key", out var k) && string.Equals(k.GetString(), "AboutMe", StringComparison.OrdinalIgnoreCase)
                    && kv.TryGetProperty("Value", out var v) && v.ValueKind == JsonValueKind.String)
                {
                    aboutMe = v.GetString();
                    break;
                }
            }
        }

        return new SharePointProfile
        {
            AccountName = Str("AccountName"),
            AboutMe = aboutMe,
            PersonalUrl = StrOrNull("PersonalUrl"),
            Skills = MultiValue("SPS-Skills"),
            Interests = MultiValue("SPS-Interests"),
            PastProjects = MultiValue("SPS-PastProjects"),
            Responsibilities = MultiValue("SPS-Responsibility"),
            Schools = MultiValue("SPS-School")
        };
    }
}
