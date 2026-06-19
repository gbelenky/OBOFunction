using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions.Authentication;
using OBOFunction.Models;

namespace OBOFunction.Services;

public sealed class GraphProfileService : IGraphProfileService
{
    private readonly IConfidentialClientApplication _cca;
    private readonly string[] _graphScopes;
    private readonly string[] _sharePointScopes;
    private readonly string? _sharePointRootSiteUrl;
    private readonly ILogger<GraphProfileService> _logger;
    private static readonly HttpClient HttpClient = new();
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public GraphProfileService(IConfiguration config, ILogger<GraphProfileService> logger)
    {
        _logger = logger;

        var tenantId = config["AzureAd:TenantId"]
            ?? throw new InvalidOperationException("AzureAd:TenantId is required.");
        var clientId = config["AzureAd:ClientId"]
            ?? throw new InvalidOperationException("AzureAd:ClientId is required.");
        var clientSecret = config["AzureAd:ClientSecret"]
            ?? throw new InvalidOperationException("AzureAd:ClientSecret is required.");

        var graphScopesRaw = config["Graph:Scopes"] ?? "User.Read";
        _graphScopes = graphScopesRaw
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(s => $"https://graph.microsoft.com/{s}")
            .ToArray();

        _sharePointRootSiteUrl = config["SharePoint:RootSiteUrl"];
        var spScopesRaw = config["SharePoint:Scopes"] ?? "AllSites.Read User.Read.All";
        _sharePointScopes = string.IsNullOrWhiteSpace(_sharePointRootSiteUrl)
            ? Array.Empty<string>()
            : spScopesRaw
                .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(s => $"{_sharePointRootSiteUrl.TrimEnd('/')}/{s}")
                .ToArray();

        _cca = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
            .Build();
    }

    public async Task<UserProfile> GetMyProfileAsync(string userAssertion, CancellationToken ct)
    {
        var assertion = new UserAssertion(userAssertion);

        // OBO #1 — Graph token for /me + photo
        var graphResult = await _cca.AcquireTokenOnBehalfOf(_graphScopes, assertion)
            .ExecuteAsync(ct)
            .ConfigureAwait(false);

        var graph = new GraphServiceClient(new BearerTokenProvider(graphResult.AccessToken).ToHttpClient());

        var me = await graph.Me.GetAsync(cancellationToken: ct).ConfigureAwait(false)
            ?? throw new InvalidOperationException("Graph /me returned null.");

        string? photoB64 = null;
        string? photoCt = null;
        try
        {
            using var photo = await graph.Me.Photo.Content
                .GetAsync(cancellationToken: ct)
                .ConfigureAwait(false);

            if (photo is not null)
            {
                using var ms = new MemoryStream();
                await photo.CopyToAsync(ms, ct).ConfigureAwait(false);
                photoB64 = Convert.ToBase64String(ms.ToArray());
                photoCt = "image/jpeg";
            }
        }
        catch (Exception ex)
        {
            _logger.LogInformation(ex, "User has no profile photo or photo fetch failed.");
        }

        // OBO #2 — SharePoint token for UPS properties (optional)
        SharePointProfile? spProfile = null;
        if (!string.IsNullOrWhiteSpace(_sharePointRootSiteUrl) && _sharePointScopes.Length > 0)
        {
            try
            {
                var spResult = await _cca.AcquireTokenOnBehalfOf(_sharePointScopes, assertion)
                    .ExecuteAsync(ct)
                    .ConfigureAwait(false);

                spProfile = await FetchSharePointProfileAsync(spResult.AccessToken, ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "SharePoint UPS fetch failed; returning Graph-only profile.");
            }
        }

        return new UserProfile(
            Id: me.Id ?? string.Empty,
            DisplayName: me.DisplayName ?? string.Empty,
            GivenName: me.GivenName,
            Surname: me.Surname,
            UserPrincipalName: me.UserPrincipalName,
            Mail: me.Mail,
            JobTitle: me.JobTitle ?? spProfile?.ExtendedProperties.GetValueOrDefault("SPS-JobTitle"),
            Department: me.Department
                ?? spProfile?.ExtendedProperties.GetValueOrDefault("Department")
                ?? spProfile?.ExtendedProperties.GetValueOrDefault("SPS-Department"),
            OfficeLocation: me.OfficeLocation ?? spProfile?.ExtendedProperties.GetValueOrDefault("Office"),
            PreferredLanguage: me.PreferredLanguage,
            MobilePhone: me.MobilePhone,
            BusinessPhones: me.BusinessPhones?.ToArray() ?? Array.Empty<string>(),
            PhotoBase64: photoB64,
            PhotoContentType: photoCt,
            SharePointProfile: spProfile);
    }

    private async Task<SharePointProfile> FetchSharePointProfileAsync(string spAccessToken, CancellationToken ct)
    {
        var url = $"{_sharePointRootSiteUrl!.TrimEnd('/')}/_api/SP.UserProfiles.PeopleManager/GetMyProperties";

        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", spAccessToken);
        req.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        req.Headers.Add("Accept", "application/json;odata=nometadata");

        using var resp = await HttpClient.SendAsync(req, ct).ConfigureAwait(false);
        resp.EnsureSuccessStatusCode();

        await using var stream = await resp.Content.ReadAsStreamAsync(ct).ConfigureAwait(false);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct).ConfigureAwait(false);
        var root = doc.RootElement;

        string Str(string p) => root.TryGetProperty(p, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";
        string? StrOrNull(string p) => root.TryGetProperty(p, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

        IReadOnlyList<string> MultiValue(string propName)
        {
            if (!root.TryGetProperty("UserProfileProperties", out var arr) || arr.ValueKind != JsonValueKind.Array)
                return Array.Empty<string>();

            foreach (var kv in arr.EnumerateArray())
            {
                if (kv.TryGetProperty("Key", out var k) && string.Equals(k.GetString(), propName, StringComparison.OrdinalIgnoreCase))
                {
                    if (kv.TryGetProperty("Value", out var v) && v.ValueKind == JsonValueKind.String)
                    {
                        var raw = v.GetString();
                        if (string.IsNullOrWhiteSpace(raw)) return Array.Empty<string>();
                        return raw.Split('|', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                    }
                }
            }
            return Array.Empty<string>();
        }

        var extended = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (root.TryGetProperty("UserProfileProperties", out var props) && props.ValueKind == JsonValueKind.Array)
        {
            foreach (var kv in props.EnumerateArray())
            {
                if (kv.TryGetProperty("Key", out var k) && kv.TryGetProperty("Value", out var v))
                {
                    var key = k.GetString();
                    var val = v.ValueKind == JsonValueKind.String ? v.GetString() : v.ToString();
                    if (!string.IsNullOrEmpty(key) && val is not null)
                        extended[key] = val;
                }
            }
        }

        return new SharePointProfile(
            AccountName: Str("AccountName"),
            AboutMe: StrOrNull("AboutMe") ?? extended.GetValueOrDefault("AboutMe"),
            PictureUrl: StrOrNull("PictureUrl"),
            PersonalUrl: StrOrNull("PersonalUrl"),
            Skills: MultiValue("SPS-Skills"),
            Interests: MultiValue("SPS-Interests"),
            PastProjects: MultiValue("SPS-PastProjects"),
            Responsibilities: MultiValue("SPS-Responsibility"),
            Schools: MultiValue("SPS-School"),
            ExtendedProperties: extended);
    }

    private sealed class BearerTokenProvider : IAccessTokenProvider
    {
        private readonly string _token;
        public BearerTokenProvider(string token) => _token = token;

        public AllowedHostsValidator AllowedHostsValidator { get; } = new();

        public Task<string> GetAuthorizationTokenAsync(
            Uri uri,
            Dictionary<string, object>? additionalAuthenticationContext = null,
            CancellationToken cancellationToken = default)
            => Task.FromResult(_token);

        public HttpClient ToHttpClient()
        {
            var authProvider = new BaseBearerTokenAuthenticationProvider(this);
            return GraphClientFactory.Create(authProvider);
        }
    }
}
