using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using OBOFunction.Models;

namespace OBOFunction.Services;

/// <summary>
/// Invokes the deployed SharePoint hosted agent through the Foundry Responses API.
/// </summary>
/// <remarks>
/// <para>
/// <b>Shape 1 (default, GA):</b> when no toolbox is configured, the agent is called with the
/// Function's <see cref="DefaultAzureCredential"/> (managed identity in Azure / developer
/// credential locally). The signed-in user's profile — already fetched via OBO by
/// <c>GET /api/profile</c> — is injected as system context so the agent reasons over the real
/// user without ever seeing a user token. No agent secret reaches the SPFx browser.
/// </para>
/// <para>
/// <b>Shape 2 (opt-in, preview):</b> when <c>Foundry:ToolboxMcpUrl</c> is set, the proxy exchanges
/// the inbound SPFx user token (OBO) for an AI-audience user token and calls Responses
/// <i>as the signed-in user</i>, attaching the Foundry toolbox as an <c>mcp</c> tool. The Foundry
/// runtime then performs OAuth identity passthrough so the toolbox's MCP server (a managed
/// Microsoft 365 server, or this Function's own <c>/api/mcp</c> server) acts as the user. If the
/// user has not yet consented, the run returns an <c>oauth_consent_request</c>; this client
/// surfaces the <c>consent_link</c> so SPFx can complete the one-time consent ceremony and resume
/// via <c>previous_response_id</c>.
/// </para>
/// <para>
/// A direct HTTP call (rather than a preview SDK) keeps this stable across Agent Framework preview
/// churn. The endpoint, API version, scopes, and toolbox descriptor are all <c>Foundry:*</c>
/// settings because the data-plane surface is still evolving.
/// </para>
/// </remarks>
public sealed class AgentChatClient : IAgentChatClient
{
    private static readonly HttpClient Http = new();

    private readonly TokenCredential _appCredential;
    private readonly Uri _responsesUri;
    private readonly string _agentName;
    private readonly string _appTokenScope;
    private readonly ILogger<AgentChatClient> _logger;

    // Shape 2 (toolbox OAuth identity passthrough) — all optional / config-gated.
    private readonly bool _toolboxEnabled;
    private readonly string? _toolboxMcpUrl;
    private readonly string _toolboxServerLabel;
    private readonly string? _toolboxConnectionId;
    private readonly string _featuresHeader;
    private readonly string _userTokenScope;
    private readonly IConfidentialClientApplication? _cca;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public AgentChatClient(IConfiguration config, ILogger<AgentChatClient> logger)
    {
        _logger = logger;

        var projectEndpoint = config["Foundry:ProjectEndpoint"]
            ?? throw new InvalidOperationException("Foundry:ProjectEndpoint is required for the agent proxy.");
        _agentName = config["Foundry:AgentName"]
            ?? throw new InvalidOperationException("Foundry:AgentName is required for the agent proxy.");

        var apiVersion = config["Foundry:ApiVersion"] ?? "2025-05-01";
        _appTokenScope = config["Foundry:TokenScope"] ?? "https://ai.azure.com/.default";

        // Allow a full override; otherwise derive {endpoint}/responses?api-version=...
        var responsesUrl = config["Foundry:ResponsesUrl"];
        if (string.IsNullOrWhiteSpace(responsesUrl))
        {
            responsesUrl = $"{projectEndpoint.TrimEnd('/')}/responses?api-version={apiVersion}";
        }
        _responsesUri = new Uri(responsesUrl);

        _appCredential = new DefaultAzureCredential();

        // ---- Shape 2 wiring (only active when a toolbox MCP URL is configured) ----
        _toolboxMcpUrl = config["Foundry:ToolboxMcpUrl"];
        _toolboxEnabled = !string.IsNullOrWhiteSpace(_toolboxMcpUrl);
        _toolboxServerLabel = config["Foundry:ToolboxServerLabel"] ?? "SharePointProfile";
        _toolboxConnectionId = config["Foundry:ToolboxConnectionId"];
        _featuresHeader = config["Foundry:FeaturesHeader"] ?? "Toolboxes=V1Preview";
        _userTokenScope = config["Foundry:UserTokenScope"] ?? "https://ai.azure.com/.default";

        if (_toolboxEnabled)
        {
            // Build the confidential client used to OBO-exchange the SPFx user token for an
            // AI-audience user token, so the Responses run executes as the signed-in user.
            var tenantId = config["AzureAd:TenantId"]
                ?? throw new InvalidOperationException("AzureAd:TenantId is required when Foundry:ToolboxMcpUrl is set (Shape 2 OBO).");
            var clientId = config["AzureAd:ClientId"]
                ?? throw new InvalidOperationException("AzureAd:ClientId is required when Foundry:ToolboxMcpUrl is set (Shape 2 OBO).");
            var clientSecret = config["AzureAd:ClientSecret"]
                ?? throw new InvalidOperationException("AzureAd:ClientSecret is required when Foundry:ToolboxMcpUrl is set (Shape 2 OBO).");

            _cca = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();

            _logger.LogInformation(
                "Agent proxy: Shape 2 toolbox passthrough ENABLED (server_label={Label}).", _toolboxServerLabel);
        }
    }

    public async Task<AgentChatReply> ChatAsync(
        AgentChatRequest request, string? userAssertion = null, CancellationToken ct = default)
    {
        return _toolboxEnabled
            ? await ChatViaToolboxAsync(request, userAssertion, ct).ConfigureAwait(false)
            : await ChatViaManagedIdentityAsync(request, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Shape 1: call the agent as the Function (managed identity) and inject the pre-fetched
    /// user profile as system context. Unchanged, GA behaviour.
    /// </summary>
    private async Task<AgentChatReply> ChatViaManagedIdentityAsync(AgentChatRequest request, CancellationToken ct)
    {
        var token = await _appCredential
            .GetTokenAsync(new TokenRequestContext([_appTokenScope]), ct)
            .ConfigureAwait(false);

        var input = new List<object>();
        if (request.UserProfile is not null)
        {
            input.Add(new
            {
                role = "system",
                content = "The signed-in user's profile (JSON): " +
                          JsonSerializer.Serialize(request.UserProfile, JsonOpts)
            });
        }
        input.Add(new { role = "user", content = request.Message });

        var payload = new Dictionary<string, object?>
        {
            ["model"] = _agentName,
            ["input"] = input,
            ["previous_response_id"] = request.PreviousResponseId
        };

        var raw = await PostResponsesAsync(payload, token.Token, withFeaturesHeader: false, ct).ConfigureAwait(false);
        return ParseReply(raw);
    }

    /// <summary>
    /// Shape 2: OBO-exchange the user token for an AI-audience user token, then call the agent
    /// as the signed-in user with the Foundry toolbox attached as an <c>mcp</c> tool. Surfaces the
    /// OAuth consent link when the user has not yet granted delegated consent.
    /// </summary>
    private async Task<AgentChatReply> ChatViaToolboxAsync(AgentChatRequest request, string? userAssertion, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(userAssertion))
            throw new UnauthorizedAccessException("Shape 2 toolbox passthrough requires the signed-in user's token.");

        // OBO: user token (aud = this API) -> user token (aud = AI Foundry data plane).
        AuthenticationResult userAiToken;
        try
        {
            userAiToken = await _cca!
                .AcquireTokenOnBehalfOf([_userTokenScope], new UserAssertion(userAssertion))
                .ExecuteAsync(ct)
                .ConfigureAwait(false);
        }
        catch (MsalUiRequiredException ex)
        {
            // The user must consent to the AI data-plane scope before the agent can act as them.
            _logger.LogInformation(ex, "OBO to AI scope requires user consent.");
            throw new UnauthorizedAccessException(
                "The signed-in user must consent to the AI Foundry scope before Shape 2 can run. " +
                "Grant admin/user consent for the configured Foundry:UserTokenScope.");
        }

        // Foundry toolbox surfaced to the run as an MCP tool. The runtime performs OAuth identity
        // passthrough using the user identity on this (user-authenticated) Responses call.
        var mcpTool = new Dictionary<string, object?>
        {
            ["type"] = "mcp",
            ["server_label"] = _toolboxServerLabel,
            ["server_url"] = _toolboxMcpUrl,
            ["require_approval"] = "never"
        };
        if (!string.IsNullOrWhiteSpace(_toolboxConnectionId))
            mcpTool["project_connection_id"] = _toolboxConnectionId;

        var input = new List<object> { new { role = "user", content = request.Message } };

        var payload = new Dictionary<string, object?>
        {
            ["model"] = _agentName,
            ["input"] = input,
            ["tools"] = new[] { mcpTool },
            ["previous_response_id"] = request.PreviousResponseId
        };

        var raw = await PostResponsesAsync(payload, userAiToken.AccessToken, withFeaturesHeader: true, ct).ConfigureAwait(false);
        return ParseReply(raw);
    }

    private async Task<string> PostResponsesAsync(
        Dictionary<string, object?> payload, string bearer, bool withFeaturesHeader, CancellationToken ct)
    {
        using var httpReq = new HttpRequestMessage(HttpMethod.Post, _responsesUri);
        httpReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearer);
        if (withFeaturesHeader)
            httpReq.Headers.TryAddWithoutValidation("Foundry-Features", _featuresHeader);
        httpReq.Content = new StringContent(
            JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        using var resp = await Http.SendAsync(httpReq, ct).ConfigureAwait(false);
        var raw = await resp.Content.ReadAsStringAsync(ct).ConfigureAwait(false);

        if (!resp.IsSuccessStatusCode)
        {
            _logger.LogError("Agent Responses call failed ({Status}): {Body}", (int)resp.StatusCode, raw);
            throw new HttpRequestException($"Agent call failed with status {(int)resp.StatusCode}.");
        }

        return raw;
    }

    /// <summary>
    /// Parses a Responses payload into a reply. Delegates to <see cref="ResponsesPayloadParser"/>
    /// so the (preview, fragile) parsing is unit-testable in isolation.
    /// </summary>
    private static AgentChatReply ParseReply(string json) => ResponsesPayloadParser.ParseReply(json);
}
