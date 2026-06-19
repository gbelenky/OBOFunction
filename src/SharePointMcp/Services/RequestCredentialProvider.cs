using Azure.Core;
using Azure.Identity;
using Microsoft.AspNetCore.Http;

namespace SharePointMcp.Services;

/// <summary>
/// Resolves the <see cref="TokenCredential"/> to use for the CURRENT request, implementing
/// the server's two-mode authentication contract:
///
/// <list type="bullet">
/// <item>
///   <b>User mode (Layer B / production passthrough):</b> the inbound request carries
///   <c>Authorization: Bearer &lt;user token&gt;</c> — injected by Foundry's OAuth connection
///   layer. That token is wrapped via <see cref="StaticTokenCredential"/> and used as-is, so
///   downstream calls act AS THE USER. No OBO exchange happens here.
/// </item>
/// <item>
///   <b>App mode (Layer A / local dev bypass):</b> no bearer token is present, so
///   <see cref="DefaultAzureCredential"/> (your <c>az login</c> / VS Code / managed identity)
///   is used. Downstream calls act as the developer or the app — useful for fast inner-loop
///   testing without a tunnel or consent, but it only ever returns the developer's own data.
/// </item>
/// </list>
/// </summary>
public sealed class RequestCredentialProvider
{
    private readonly IHttpContextAccessor _httpContextAccessor;
    private readonly DefaultAzureCredential _defaultCredential;

    public RequestCredentialProvider(IHttpContextAccessor httpContextAccessor, DefaultAzureCredential defaultCredential)
    {
        _httpContextAccessor = httpContextAccessor;
        _defaultCredential = defaultCredential;
    }

    /// <summary>
    /// Returns the credential for the current request and whether it represents a real user
    /// (passthrough bearer token) or the app/developer identity.
    /// </summary>
    public (TokenCredential Credential, bool IsUserToken) Resolve()
    {
        var header = _httpContextAccessor.HttpContext?.Request.Headers.Authorization.ToString();

        if (!string.IsNullOrWhiteSpace(header) &&
            header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        {
            var token = header["Bearer ".Length..].Trim();
            if (!string.IsNullOrWhiteSpace(token))
                return (new StaticTokenCredential(token), true);
        }

        return (_defaultCredential, false);
    }
}
