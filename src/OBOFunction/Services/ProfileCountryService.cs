using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;

namespace OBOFunction.Services;

/// <summary>
/// "Option A" country resolver. OBO-exchanges the inbound SPFx user token to (1) SharePoint, to
/// read the custom UPS attribute <c>IntranetCountry</c> via <c>PeopleManager/GetMyProperties</c>,
/// and falls back to (2) Microsoft Graph <c>/me</c> (<c>country</c>/<c>usageLocation</c>). The
/// proxy app registration carries the required delegated permissions (SharePoint
/// <c>AllSites.Read</c> + <c>User.Read.All</c>, Graph <c>User.Read</c>).
/// <para>
/// Best-effort by design: any failure (missing consent, app-only identity, 401) is logged and
/// returns <c>null</c>, so the chat turn still completes and country-filtered features fall back
/// to global/default results. The proxy never reads or stores the full profile — only the country.
/// </para>
/// </summary>
public sealed class ProfileCountryService : IProfileCountryService
{
    private static readonly HttpClient Http = new();

    private readonly IConfidentialClientApplication _cca;
    private readonly string? _sharePointRootSiteUrl;
    private readonly ILogger<ProfileCountryService> _logger;

    public ProfileCountryService(IConfiguration config, ILogger<ProfileCountryService> logger)
    {
        _logger = logger;
        _sharePointRootSiteUrl = config["SharePoint:RootSiteUrl"]?.TrimEnd('/');

        var tenantId = config["AzureAd:TenantId"]
            ?? throw new InvalidOperationException("AzureAd:TenantId is required for the profile country OBO.");
        var clientId = config["AzureAd:ClientId"]
            ?? throw new InvalidOperationException("AzureAd:ClientId is required for the profile country OBO.");
        var clientSecret = config["AzureAd:ClientSecret"]
            ?? throw new InvalidOperationException("AzureAd:ClientSecret is required for the profile country OBO.");

        _cca = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
            .Build();
    }

    public async Task<string?> GetCountryAsync(string userAssertion, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(userAssertion))
            return null;

        var assertion = new UserAssertion(userAssertion);

        // 1) Primary: SharePoint UPS custom attribute IntranetCountry (where the demo country lives).
        if (!string.IsNullOrWhiteSpace(_sharePointRootSiteUrl))
        {
            try
            {
                var spResource = new Uri(_sharePointRootSiteUrl!).GetLeftPart(UriPartial.Authority);
                var spToken = await _cca
                    .AcquireTokenOnBehalfOf([$"{spResource}/.default"], assertion)
                    .ExecuteAsync(ct)
                    .ConfigureAwait(false);

                var country = await FetchUpsCountryAsync(spToken.AccessToken, ct).ConfigureAwait(false);
                if (!string.IsNullOrWhiteSpace(country))
                    return country;
            }
            catch (Exception ex)
            {
                _logger.LogInformation(
                    ex, "SharePoint UPS country lookup failed ({Type}); falling back to Graph.", ex.GetType().Name);
            }
        }

        // 2) Fallback: Graph /me country / usageLocation.
        try
        {
            var graphToken = await _cca
                .AcquireTokenOnBehalfOf(["https://graph.microsoft.com/.default"], assertion)
                .ExecuteAsync(ct)
                .ConfigureAwait(false);

            return await FetchGraphCountryAsync(graphToken.AccessToken, ct).ConfigureAwait(false);
        }
        catch (Exception ex)
        {
            _logger.LogInformation(
                ex, "Graph country lookup failed ({Type}); no country resolved.", ex.GetType().Name);
            return null;
        }
    }

    private async Task<string?> FetchUpsCountryAsync(string bearer, CancellationToken ct)
    {
        var url = $"{_sharePointRootSiteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearer);
        req.Headers.Add("Accept", "application/json;odata=nometadata");

        using var resp = await Http.SendAsync(req, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();

        await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);

        if (!doc.RootElement.TryGetProperty("UserProfileProperties", out var arr)
            || arr.ValueKind != JsonValueKind.Array)
        {
            return null;
        }

        foreach (var kv in arr.EnumerateArray())
        {
            if (kv.TryGetProperty("Key", out var k)
                && string.Equals(k.GetString(), "IntranetCountry", StringComparison.OrdinalIgnoreCase)
                && kv.TryGetProperty("Value", out var v))
            {
                var val = v.ValueKind == JsonValueKind.String ? v.GetString() : v.ToString();
                return string.IsNullOrWhiteSpace(val) ? null : val;
            }
        }

        return null;
    }

    private static async Task<string?> FetchGraphCountryAsync(string bearer, CancellationToken ct)
    {
        const string url = "https://graph.microsoft.com/v1.0/me?$select=country,usageLocation";
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearer);
        req.Headers.Add("Accept", "application/json");

        using var resp = await Http.SendAsync(req, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();

        await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
        var root = doc.RootElement;

        if (root.TryGetProperty("country", out var c) && c.ValueKind == JsonValueKind.String)
        {
            var country = c.GetString();
            if (!string.IsNullOrWhiteSpace(country))
                return country;
        }

        if (root.TryGetProperty("usageLocation", out var u) && u.ValueKind == JsonValueKind.String)
        {
            var loc = u.GetString();
            if (!string.IsNullOrWhiteSpace(loc))
                return loc;
        }

        return null;
    }
}
