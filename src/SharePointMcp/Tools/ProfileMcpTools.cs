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
    [Description("Fetch the current user's Microsoft Graph and SharePoint profile " +
                 "(name, job title, department, office, skills, interests, responsibilities, etc.). " +
                 "Acts on behalf of the signed-in user when a delegated token is present. " +
                 "Returns a JSON object.")]
    public async Task<string> GetSharePointProfileAsync(CancellationToken cancellationToken = default)
    {
        var (credential, isUserToken) = _credentialProvider.Resolve();

        // Both identities can mint a SharePoint-audience token now: the user path uses
        // OnBehalfOfCredential (fresh OBO exchange per scope) and the dev path uses
        // DefaultAzureCredential, so the SharePoint UPS call is always attempted and custom
        // attributes surface in every context.
        var service = new SharePointProfileService(
            credential,
            _sharePointRootSiteUrl,
            allowSharePointUps: true,
            resolvedVia: isUserToken ? "user" : "app");

        var profile = await service.GetMyProfileAsync(cancellationToken).ConfigureAwait(false);
        return JsonSerializer.Serialize(profile, JsonOptions);
    }
}
