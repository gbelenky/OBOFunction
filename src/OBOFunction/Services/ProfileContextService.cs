using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using OBOFunction.Models;

namespace OBOFunction.Services;

/// <summary>
/// "Option A" profile resolver. OBO-exchanges the inbound SPFx user token to (1) SharePoint, to
/// read the user's own profile via <c>PeopleManager/GetMyProperties</c> (name, job title,
/// responsibilities, past projects, interests, and the custom <c>IntranetCountry</c> attribute),
/// merged with (2) Microsoft Graph <c>/me</c> (display/given name, mail, country fallback). Only
/// the caller's own profile is read (<c>GetMyProperties</c> / <c>/me</c>) — never other users.
/// <para>
/// Best-effort by design: any failure (missing consent, app-only identity, 401) is logged and the
/// corresponding fields are left empty, so the chat turn still completes (generic greeting,
/// global/default results). No profile is stored.
/// </para>
/// </summary>
public sealed class ProfileContextService : IProfileContextService
{
    private static readonly HttpClient Http = new();

    private readonly IConfidentialClientApplication _cca;
    private readonly string? _sharePointRootSiteUrl;
    private readonly ILogger<ProfileContextService> _logger;

    public ProfileContextService(IConfiguration config, ILogger<ProfileContextService> logger)
    {
        _logger = logger;
        _sharePointRootSiteUrl = config["SharePoint:RootSiteUrl"]?.TrimEnd('/');

        var tenantId = config["AzureAd:TenantId"]
            ?? throw new InvalidOperationException("AzureAd:TenantId is required for the profile OBO.");
        var clientId = config["AzureAd:ClientId"]
            ?? throw new InvalidOperationException("AzureAd:ClientId is required for the profile OBO.");
        var clientSecret = config["AzureAd:ClientSecret"]
            ?? throw new InvalidOperationException("AzureAd:ClientSecret is required for the profile OBO.");

        _cca = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
            .Build();
    }

    public async Task<UserProfileContext> GetProfileAsync(string userAssertion, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(userAssertion))
            return new UserProfileContext();

        var assertion = new UserAssertion(userAssertion);

        // 1) Primary: SharePoint UPS (GetMyProperties) — name + slim fields + IntranetCountry.
        UpsFields ups = default;
        if (!string.IsNullOrWhiteSpace(_sharePointRootSiteUrl))
        {
            try
            {
                var spResource = new Uri(_sharePointRootSiteUrl!).GetLeftPart(UriPartial.Authority);
                var spToken = await _cca
                    .AcquireTokenOnBehalfOf([$"{spResource}/.default"], assertion)
                    .ExecuteAsync(ct)
                    .ConfigureAwait(false);

                ups = await FetchUpsAsync(spToken.AccessToken, ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger.LogInformation(
                    ex, "SharePoint UPS profile lookup failed ({Type}); falling back to Graph.", ex.GetType().Name);
            }
        }

        // 2) Graph /me — display/given name, mail, job title, country fallback.
        GraphFields graph = default;
        try
        {
            var graphToken = await _cca
                .AcquireTokenOnBehalfOf(["https://graph.microsoft.com/.default"], assertion)
                .ExecuteAsync(ct)
                .ConfigureAwait(false);

            graph = await FetchGraphAsync(graphToken.AccessToken, ct).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogInformation(
                ex, "Graph profile lookup failed ({Type}); using UPS-only fields.", ex.GetType().Name);
        }

        return new UserProfileContext
        {
            Name = ups.PreferredName ?? graph.DisplayName,
            FirstName = graph.GivenName,
            Email = graph.Mail,
            JobTitle = ups.JobTitle ?? graph.JobTitle,
            Responsibilities = ups.Responsibilities ?? [],
            PastProjects = ups.PastProjects ?? [],
            Interests = ups.Interests ?? [],
            Country = ups.Country ?? graph.Country
        };
    }

    private readonly record struct UpsFields(
        string? PreferredName,
        string? JobTitle,
        IReadOnlyList<string>? Responsibilities,
        IReadOnlyList<string>? PastProjects,
        IReadOnlyList<string>? Interests,
        string? Country);

    private readonly record struct GraphFields(
        string? DisplayName,
        string? GivenName,
        string? Mail,
        string? JobTitle,
        string? Country);

    private async Task<UpsFields> FetchUpsAsync(string bearer, CancellationToken ct)
    {
        var url = $"{_sharePointRootSiteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearer);
        req.Headers.Add("Accept", "application/json;odata=nometadata");

        using var resp = await Http.SendAsync(req, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();

        await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
        var root = doc.RootElement;

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
            PreferredName: Single("PreferredName"),
            JobTitle: Single("SPS-JobTitle"),
            Responsibilities: Multi("SPS-Responsibility"),
            PastProjects: Multi("SPS-PastProjects"),
            Interests: Multi("SPS-Interests"),
            Country: Single("IntranetCountry"));
    }

    private static async Task<GraphFields> FetchGraphAsync(string bearer, CancellationToken ct)
    {
        const string url =
            "https://graph.microsoft.com/v1.0/me?$select=displayName,givenName,mail,userPrincipalName,jobTitle,country,usageLocation";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearer);
        req.Headers.Add("Accept", "application/json");

        using var resp = await Http.SendAsync(req, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();

        await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
        var root = doc.RootElement;

        string? Str(string name) =>
            root.TryGetProperty(name, out var e) && e.ValueKind == JsonValueKind.String
                ? (string.IsNullOrWhiteSpace(e.GetString()) ? null : e.GetString())
                : null;

        return new GraphFields(
            DisplayName: Str("displayName"),
            GivenName: Str("givenName"),
            Mail: Str("mail") ?? Str("userPrincipalName"),
            JobTitle: Str("jobTitle"),
            Country: Str("country") ?? Str("usageLocation"));
    }
}
