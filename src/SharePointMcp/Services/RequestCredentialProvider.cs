using Azure.Core;
using Azure.Identity;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;

namespace SharePointMcp.Services;

/// <summary>
/// Resolves the <see cref="TokenCredential"/> for the CURRENT request. This is the single
/// identity model for the server — the same code runs locally, in Foundry, and behind the
/// SPFx proxy; only the inbound identity differs:
///
/// <list type="bullet">
/// <item>
///   <b>User identity (Foundry Playground / SPFx via the proxy):</b> the inbound request carries
///   <c>Authorization: Bearer &lt;user token&gt;</c> whose audience is THIS server's app
///   registration. It is exchanged via the OAuth 2.0 On-Behalf-Of flow
///   (<see cref="OnBehalfOfCredential"/>) for Microsoft Graph <i>and</i> SharePoint tokens, so
///   downstream calls act AS THE USER and the full profile — including custom SharePoint UPS
///   attributes — is returned. <see cref="OnBehalfOfCredential"/> performs a fresh OBO exchange
///   per requested scope, which is exactly why a single inbound token can reach both resources.
/// </item>
/// <item>
///   <b>Dev identity (local inner loop):</b> no bearer token is present, so
///   <see cref="DefaultAzureCredential"/> (your <c>az login</c> / VS Code identity) is used.
///   Downstream calls act as the developer — fast iteration without a tunnel or consent.
/// </item>
/// </list>
/// </summary>
public sealed class RequestCredentialProvider
{
    private readonly IHttpContextAccessor _httpContextAccessor;
    private readonly DefaultAzureCredential _defaultCredential;
    private readonly string? _tenantId;
    private readonly string? _clientId;
    private readonly string? _clientSecret;

    public RequestCredentialProvider(
        IHttpContextAccessor httpContextAccessor,
        DefaultAzureCredential defaultCredential,
        IConfiguration configuration)
    {
        _httpContextAccessor = httpContextAccessor;
        _defaultCredential = defaultCredential;

        // The MCP server's OWN confidential-client app registration (api://obo-sp-mcp-<env>).
        // Required only for the user-token (OBO) path; the local dev path never reads these.
        // The client secret comes from Key Vault (Mcp--ClientSecret -> Mcp:ClientSecret); tenant
        // and client id are plain app settings. AzureAd:* is accepted as a fallback for local dev.
        _tenantId = configuration["AzureAd:TenantId"];
        _clientId = configuration["AzureAd:ClientId"];
        _clientSecret = configuration["Mcp:ClientSecret"] ?? configuration["AzureAd:ClientSecret"];
    }

    /// <summary>
    /// Returns the credential for the current request and whether it represents a real user
    /// (OBO exchange of a passed-in user token) or the local developer identity.
    /// </summary>
    public (TokenCredential Credential, bool IsUserToken) Resolve()
    {
        var header = _httpContextAccessor.HttpContext?.Request.Headers.Authorization.ToString();

        if (!string.IsNullOrWhiteSpace(header) &&
            header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            var userAssertion = header["Bearer ".Length..].Trim();
            if (!string.IsNullOrWhiteSpace(userAssertion))
            {
                if (string.IsNullOrWhiteSpace(_tenantId) ||
                    string.IsNullOrWhiteSpace(_clientId) ||
                    string.IsNullOrWhiteSpace(_clientSecret))
                {
                    throw new InvalidOperationException(
                        "A user token was presented but the MCP server's OBO app registration is " +
                        "not configured. Set AzureAd:TenantId, AzureAd:ClientId and " +
                        "AzureAd:ClientSecret (the api://obo-sp-mcp app registration + secret).");
                }

                // OnBehalfOfCredential re-runs the OBO exchange for each requested scope, so the
                // same inbound user token yields a Graph-audience token for /me AND a
                // SharePoint-audience token for the User Profile Service.
                var obo = new OnBehalfOfCredential(_tenantId, _clientId, _clientSecret, userAssertion);
                return (obo, true);
            }
        }

        return (_defaultCredential, false);
    }
}
