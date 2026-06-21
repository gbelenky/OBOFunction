using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using OBOFunction.Models;

namespace OBOFunction.Services;

/// <summary>
/// A thin, <b>tool-agnostic</b> proxy that forwards a chat turn to the Foundry <b>hosted agent</b>
/// and returns its reply. The proxy is <b>completely unaware of the agent's tools</b>: it never
/// declares tools, instructions, or MCP servers — the agent owns all of that. The proxy's only
/// jobs are (1) authenticate the SPFx caller and (2) call the agent <i>as that user</i> so the
/// agent's own tools run with the user's identity.
/// </summary>
/// <remarks>
/// <para><b>Flow:</b> SPFx → Proxy → Agent → (the agent's own) tools.</para>
/// <para>
/// The agent (<c>SharePointProfileAgent</c>) owns its tools itself, e.g. the local in-process
/// <c>search_faq</c> tool. This client never references them.
/// </para>
/// <para><b>Identity (one user token, no app token):</b></para>
/// <list type="number">
/// <item><description>SPFx sends the user's token (audience = this proxy's app registration).</description></item>
/// <item><description>The proxy OBO-exchanges it for a token whose audience is the Foundry data
/// plane (<c>https://ai.azure.com</c>) — still the <i>same user</i>.</description></item>
/// <item><description>The proxy POSTs the turn to the agent's Responses endpoint with that user
/// token. No user token is ever stored.</description></item>
/// </list>
/// <para><b>Country context:</b> Foundry OAuth identity passthrough cannot deliver the per-user
/// SharePoint profile through this custom SPFx → proxy → agent chain (hosted-agent tool discovery
/// runs under the agent's managed identity, so the passthrough tool is dropped and no consent is
/// surfaced — see <c>docs/foundry-oauth-passthrough-findings.md</c>). As a result the proxy
/// resolves the user's country itself (Option A, <see cref="IProfileCountryService"/>) and the
/// endpoint injects it as plain conversation context before calling this client. This client
/// remains tool-agnostic; it just forwards the (possibly enriched) message.</para>
/// </remarks>
public sealed class AgentChatClient : IAgentChatClient
{
    private static readonly HttpClient Http = new();

    private readonly Uri _agentResponsesUri;
    private readonly string _aiPlaneScope;
    private readonly IConfidentialClientApplication _cca;
    private readonly ILogger<AgentChatClient> _logger;

    public AgentChatClient(IConfiguration config, ILogger<AgentChatClient> logger)
    {
        _logger = logger;

        // The Foundry hosted-agent Responses endpoint. Prefer an explicit override; otherwise
        // derive {projectEndpoint}/agents/{agentName}/endpoint/protocols/openai/responses?api-version=v1.
        var agentResponsesUrl = config["Foundry:AgentResponsesUrl"];
        if (string.IsNullOrWhiteSpace(agentResponsesUrl))
        {
            var projectEndpoint = config["Foundry:ProjectEndpoint"]
                ?? throw new InvalidOperationException(
                    "Foundry:AgentResponsesUrl or Foundry:ProjectEndpoint is required for the agent proxy.");
            var agentName = config["Foundry:AgentName"] ?? "SharePointProfileAgent";
            agentResponsesUrl =
                $"{projectEndpoint.TrimEnd('/')}/agents/{agentName}/endpoint/protocols/openai/responses?api-version=v1";
        }
        _agentResponsesUri = new Uri(agentResponsesUrl);

        // OBO target: the Foundry data-plane audience. Calling the agent as the user is what lets
        // the agent's UserEntraToken connection forward the user's identity to its tools.
        _aiPlaneScope = config["Foundry:TokenScope"] ?? "https://ai.azure.com/.default";

        // Confidential client for the OBO exchange (SPFx user token -> AI-plane token, same user).
        var tenantId = config["AzureAd:TenantId"]
            ?? throw new InvalidOperationException("AzureAd:TenantId is required for the agent proxy OBO.");
        var clientId = config["AzureAd:ClientId"]
            ?? throw new InvalidOperationException("AzureAd:ClientId is required for the agent proxy OBO.");
        var clientSecret = config["AzureAd:ClientSecret"]
            ?? throw new InvalidOperationException("AzureAd:ClientSecret is required for the agent proxy OBO.");

        _cca = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
            .Build();

        _logger.LogInformation("Agent proxy ready: agent_endpoint={Endpoint}.", _agentResponsesUri);
    }

    public async Task<AgentChatReply> ChatAsync(
        AgentChatRequest request, string? userAssertion = null, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(userAssertion))
            throw new UnauthorizedAccessException("The agent proxy requires the signed-in user's token.");

        // OBO the SPFx user token to the Foundry data-plane audience — still the same user.
        string aiPlaneUserToken;
        try
        {
            var result = await _cca
                .AcquireTokenOnBehalfOf([_aiPlaneScope], new UserAssertion(userAssertion))
                .ExecuteAsync(ct)
                .ConfigureAwait(false);
            aiPlaneUserToken = result.AccessToken;
        }
        catch (MsalUiRequiredException ex)
        {
            _logger.LogInformation(ex, "OBO to the Foundry data plane requires user/admin consent.");
            throw new UnauthorizedAccessException(
                "The signed-in user must consent to the Azure AI (Foundry) data-plane scope before the " +
                "agent can be called on their behalf. Grant admin/user consent for the configured " +
                "Foundry:TokenScope.");
        }

        // Pure proxy → agent: just the conversation. NO tools, NO instructions, NO MCP wiring —
        // the agent owns all of that. previous_response_id threads the multi-turn conversation and
        // is only sent when present — the endpoint rejects an explicit null (it must be a string).
        var payload = new Dictionary<string, object?>
        {
            ["input"] = new List<object> { new { role = "user", content = request.Message } }
        };

        if (!string.IsNullOrWhiteSpace(request.PreviousResponseId))
            payload["previous_response_id"] = request.PreviousResponseId;

        var raw = await PostAgentAsync(payload, aiPlaneUserToken, ct).ConfigureAwait(false);
        return ParseReply(raw);
    }

    private async Task<string> PostAgentAsync(
        Dictionary<string, object?> payload, string bearer, CancellationToken ct)
    {
        using var httpReq = new HttpRequestMessage(HttpMethod.Post, _agentResponsesUri);
        httpReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearer);
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
