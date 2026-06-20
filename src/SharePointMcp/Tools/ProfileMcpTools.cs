using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;
using SharePointMcp.Services;

namespace SharePointMcp.Tools;

/// <summary>
/// MCP tools exposed by this server. Each call resolves its identity per-request via
/// <see cref="RequestCredentialProvider"/> (user-token passthrough vs. DefaultAzureCredential),
/// so the same server serves both the production passthrough path and the local dev bypass.
/// </summary>
[McpServerToolType]
public sealed class ProfileMcpTools
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly RequestCredentialProvider _credentialProvider;
    private readonly string? _sharePointRootSiteUrl;

    public ProfileMcpTools(RequestCredentialProvider credentialProvider, IConfiguration configuration)
    {
        _credentialProvider = credentialProvider;

        var root = configuration["SHAREPOINT_ROOT_SITE_URL"];
        if (string.IsNullOrWhiteSpace(root))
        {
            var host = configuration["SHAREPOINT_TENANT_HOSTNAME"];
            if (!string.IsNullOrWhiteSpace(host))
                root = $"https://{host}";
        }
        _sharePointRootSiteUrl = root;
    }

    [McpServerTool(Name = "get_sharepoint_profile")]
    [Description("Fetch the current user's profile as a compact JSON object with exactly these fields: " +
                 "name, firstName, lastName, email (Microsoft Graph) and jobTitle, responsibilities, " +
                 "pastProjects, interests, country (SharePoint User Profile Service, including the custom " +
                 "IntranetCountry attribute). Acts on behalf of the signed-in user when a delegated token " +
                 "is present.")]
    public async Task<string> GetSharePointProfileAsync(CancellationToken cancellationToken = default)
    {
        var (credential, isUserToken) = _credentialProvider.Resolve();

        // The user path uses OnBehalfOfCredential (fresh OBO exchange per scope), so a single
        // inbound token mints both a Graph-audience token for /me and a SharePoint-audience token
        // for the User Profile Service. The dev path uses DefaultAzureCredential (az login), which
        // is also a delegated user and can call /me. Only a pure application identity (e.g. the
        // deployed UAMI when the Foundry passthrough delivered no user token) cannot call /me — the
        // service detects that and degrades gracefully instead of throwing, so the agent turn can
        // continue (e.g. run an FAQ search with no country filter).
        var service = new SharePointProfileService(
            credential,
            _sharePointRootSiteUrl,
            allowSharePointUps: true,
            resolvedVia: isUserToken ? "user" : "app");

        var profile = await service.GetMyProfileAsync(cancellationToken).ConfigureAwait(false);
        return JsonSerializer.Serialize(profile, JsonOptions);
    }
}
