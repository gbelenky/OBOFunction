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
/// Invokes the Foundry <b>model</b> Responses endpoint (OpenAI v1) with the standalone
/// SharePointMcp server attached as an <c>mcp</c> tool, so the model reads the signed-in
/// user's Graph + SharePoint profile as that user.
/// </summary>
/// <remarks>
/// <para>
/// The proxy targets the raw model deployment (e.g. gpt-4.1-mini), NOT the named hosted agent:
/// hosted agents <b>reject request-supplied tools</b> ("Not allowed"), but the per-user MCP
/// <c>authorization</c> must be attached per call. The hosted agent remains for the autonomous
/// Foundry Playground demo (baked-in URL tool, app-only/empty profile there). The agent persona
/// is carried by the <c>instructions</c> field on every request.
/// </para>
/// <para>Two independent tokens travel in one Responses request:</para>
/// <list type="bullet">
/// <item>
/// <description>
/// <b>Leg ① — call the model:</b> the request's <c>Authorization</c> header is the proxy's own
/// app identity (<see cref="DefaultAzureCredential"/> — managed identity in Azure, developer
/// credential locally) scoped to the AI data plane. It carries no user context.
/// </description>
/// </item>
/// <item>
/// <description>
/// <b>Leg ② — call the MCP server as the user:</b> the inbound SPFx user token is OBO-exchanged
/// for a token whose audience is the MCP server's app registration, and placed in the
/// <c>mcp</c> tool's <c>authorization</c> field. Foundry forwards it as
/// <c>Authorization: Bearer</c> when it dials the MCP <c>/mcp</c> endpoint; the MCP server then
/// performs its own OBO to Microsoft Graph + SharePoint. No user token is ever stored.
/// </description>
/// </item>
/// </list>
/// </remarks>
public sealed class AgentChatClient : IAgentChatClient
{
    private static readonly HttpClient Http = new();

    private readonly TokenCredential _appCredential;
    private readonly Uri _responsesUri;
    private readonly string _model;
    private readonly string _appTokenScope;
    private readonly ILogger<AgentChatClient> _logger;

    // MCP tool (leg ②) — the SharePointMcp server attached to every run as the signed-in user.
    private readonly string _mcpServerUrl;
    private readonly string _mcpServerLabel;
    private readonly string _mcpUserScope;
    private readonly IConfidentialClientApplication _cca;

    // System instruction injected on every run (there is no baked-in agent on the raw-model
    // Responses endpoint, so the proxy carries the agent persona itself).
    private const string Instructions =
        "You are the SharePoint Profile Assistant. At the START of every conversation, FIRST call " +
        "the `get_sharepoint_profile` tool to load the signed-in user's Microsoft Graph + SharePoint " +
        "profile, then use that context (job title, department, office, skills, interests, " +
        "responsibilities) in your answers. Never ask for information already present in the profile. " +
        "If a field is null, work with what is available and do not invent values.";

    public AgentChatClient(IConfiguration config, ILogger<AgentChatClient> logger)
    {
        _logger = logger;

        var projectEndpoint = config["Foundry:ProjectEndpoint"]
            ?? throw new InvalidOperationException("Foundry:ProjectEndpoint is required for the agent proxy.");

        _appTokenScope = config["Foundry:TokenScope"] ?? "https://ai.azure.com/.default";

        // The model deployment to drive (e.g. gpt-4.1-mini). The proxy calls the raw-model
        // OpenAI Responses endpoint — NOT the hosted agent — because hosted agents REJECT
        // request-supplied tools ("Not allowed"), and the per-user MCP authorization MUST be
        // attached per call. The hosted agent remains for the autonomous Playground demo.
        _model = FirstNonBlank(
                config["Foundry:ModelDeployment"],
                config["AZURE_AI_MODEL_DEPLOYMENT_NAME"],
                Environment.GetEnvironmentVariable("AZURE_AI_MODEL_DEPLOYMENT_NAME"))
            ?? "gpt-4.1-mini";

        // Allow a full override; otherwise derive the account-level OpenAI v1 responses endpoint
        // ({scheme}://{host}/openai/v1/responses) from the project endpoint host. The /v1 path does
        // NOT take an api-version query parameter.
        var responsesUrl = config["Foundry:ResponsesUrl"];
        if (string.IsNullOrWhiteSpace(responsesUrl))
        {
            var host = new Uri(projectEndpoint);
            responsesUrl = $"{host.Scheme}://{host.Authority}/openai/v1/responses";
        }
        _responsesUri = new Uri(responsesUrl);

        _appCredential = new DefaultAzureCredential();

        // ---- MCP tool wiring (leg ②) ----
        _mcpServerUrl = config["Mcp:ServerUrl"]
            ?? throw new InvalidOperationException("Mcp:ServerUrl is required (the SharePointMcp /mcp endpoint).");
        _mcpServerLabel = config["Mcp:ServerLabel"] ?? "SharePointMcp";
        // OBO target: a token whose audience is the MCP server's app registration.
        _mcpUserScope = config["Mcp:UserTokenScope"]
            ?? throw new InvalidOperationException(
                "Mcp:UserTokenScope is required (e.g. api://<mcp-app-id>/.default) so the proxy can OBO " +
                "the user token to the MCP server's audience.");

        // Confidential client used to OBO-exchange the SPFx user token (aud = this API) for a token
        // whose audience is the MCP server. This is the proxy's own app registration.
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

        _logger.LogInformation(
            "Agent proxy ready: model={Model}, mcp_server_label={Label}.", _model, _mcpServerLabel);
    }

    private static string? FirstNonBlank(params string?[] values) =>
        values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));

    public async Task<AgentChatReply> ChatAsync(
        AgentChatRequest request, string? userAssertion = null, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(userAssertion))
            throw new UnauthorizedAccessException("The agent proxy requires the signed-in user's token.");

        // Leg ②: OBO the SPFx user token to the MCP server's audience.
        string mcpUserToken;
        try
        {
            var result = await _cca
                .AcquireTokenOnBehalfOf([_mcpUserScope], new UserAssertion(userAssertion))
                .ExecuteAsync(ct)
                .ConfigureAwait(false);
            mcpUserToken = result.AccessToken;
        }
        catch (MsalUiRequiredException ex)
        {
            _logger.LogInformation(ex, "OBO to the MCP scope requires user/admin consent.");
            throw new UnauthorizedAccessException(
                "The signed-in user must consent to the SharePoint MCP API scope before the agent can read " +
                "their profile. Grant admin/user consent for the configured Mcp:UserTokenScope.");
        }

        // Leg ①: authorise the Responses call itself with the proxy's app identity.
        var appToken = await _appCredential
            .GetTokenAsync(new TokenRequestContext([_appTokenScope]), ct)
            .ConfigureAwait(false);

        // The SharePointMcp server attached as an MCP tool, authorised AS THE USER via leg ②.
        var mcpTool = new Dictionary<string, object?>
        {
            ["type"] = "mcp",
            ["server_label"] = _mcpServerLabel,
            ["server_url"] = _mcpServerUrl,
            ["authorization"] = mcpUserToken,
            ["require_approval"] = "never",
            ["allowed_tools"] = new[] { "get_sharepoint_profile" }
        };

        var payload = new Dictionary<string, object?>
        {
            ["model"] = _model,
            ["instructions"] = Instructions,
            ["input"] = new List<object> { new { role = "user", content = request.Message } },
            ["tools"] = new[] { mcpTool },
            ["previous_response_id"] = request.PreviousResponseId
        };

        var raw = await PostResponsesAsync(payload, appToken.Token, ct).ConfigureAwait(false);
        return ParseReply(raw);
    }

    private async Task<string> PostResponsesAsync(
        Dictionary<string, object?> payload, string bearer, CancellationToken ct)
    {
        using var httpReq = new HttpRequestMessage(HttpMethod.Post, _responsesUri);
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
