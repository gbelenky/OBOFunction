using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OBOFunction.Models;
using System.Net.Http.Headers;
using System.Text.Json;

namespace OBOFunction.Services;

/// <summary>
/// Calls the Foundry HOSTED agent as the signed-in user (OBO identity pass-through).
/// 
/// <para><b>Flow:</b> SPFx (user token) → Proxy validates JWT and extracts user assertion → 
/// Creates OBO credential from app registration + user assertion → Calls hosted agent Responses endpoint
/// as the user → Agent's Toolbox connection resolves THIS user's tokens via OBO.</para>
///
/// <para><b>Identity:</b> The user's Bearer token (assertion) is used to create an OnBehalfOfCredential.
/// When the OnBehalfOfCredential needs to acquire a token for the Foundry scope (https://ai.azure.com/.default),
/// it performs an OBO exchange to get a token for the user. This token is then used to call the agent.</para>
///
/// <para><b>Conversation state:</b> Multi-turn conversations are tied to <c>response_id</c> (conversation ID)
/// which the client stores and passes back as <c>previous_response_id</c> on follow-up turns.</para>
/// </summary>
public sealed class AgentChatClient : IAgentChatClient
{
    private readonly Uri _projectEndpoint;
    private readonly string _agentName;
    private readonly string _tenantId;
    private readonly string _clientId;
    private readonly string _clientSecret;
    private readonly string _tokenScope;
    private readonly HttpClient _httpClient;
    private readonly ILogger<AgentChatClient> _logger;

    public AgentChatClient(IConfiguration config, ILogger<AgentChatClient> logger)
    {
        _logger = logger;

        _projectEndpoint = new Uri(config["Foundry:ProjectEndpoint"]
            ?? throw new InvalidOperationException(
                "Foundry:ProjectEndpoint is required (format: https://<resource>.services.ai.azure.com/api/projects/<projectName>)."));

        _agentName = config["Foundry:AgentName"]
            ?? throw new InvalidOperationException("Foundry:AgentName is required (the hosted agent name/id).");

        _tenantId = config["AzureAd:TenantId"]
            ?? throw new InvalidOperationException("AzureAd:TenantId is required.");

        _clientId = config["AzureAd:ClientId"]
            ?? throw new InvalidOperationException("AzureAd:ClientId is required.");

        _clientSecret = config["AzureAd:ClientSecret"]
            ?? throw new InvalidOperationException("AzureAd:ClientSecret is required.");

        _tokenScope = config["Foundry:TokenScope"] ?? "https://ai.azure.com/.default";

        _httpClient = new HttpClient();
        _logger.LogInformation("Agent proxy initialized: project={Project}, agent={Agent}.", 
            _projectEndpoint, _agentName);
    }

    public async Task<AgentChatReply> ChatAsync(
        AgentChatRequest request, string? userAssertion = null, CancellationToken ct = default)
    {
        System.Diagnostics.Activity.Current?.SetTag(
            OBOFunction.Observability.AgentTelemetry.Attr.AgentName, _agentName);

        if (string.IsNullOrWhiteSpace(userAssertion))
            throw new UnauthorizedAccessException("The agent proxy requires the signed-in user's token.");

        // Create OBO credential: this will perform OBO exchanges when requesting Foundry tokens.
        var oboCredential = new OnBehalfOfCredential(
            tenantId: _tenantId,
            clientId: _clientId,
            clientSecret: _clientSecret,
            userAssertion: userAssertion);

        try
        {
            // Acquire a Foundry-scoped token AS THE USER via OBO.
            var accessToken = await oboCredential.GetTokenAsync(
                new Azure.Core.TokenRequestContext(new[] { _tokenScope }),
                ct).ConfigureAwait(false);

            _logger.LogInformation("OBO token acquired for user; calling agent endpoint.");

            // POST to the Foundry agent endpoint.
            // Format: POST https://<resource>.services.ai.azure.com/api/projects/<projectName>/responses/<agentName>
            var responseUrl = $"{_projectEndpoint}/responses/{_agentName}";

            var requestBody = new
            {
                message = request.Message ?? string.Empty,
                previous_response_id = request.PreviousResponseId
            };

            var json = JsonSerializer.Serialize(requestBody);
            _logger.LogInformation("Calling agent endpoint: {Url} with body: {Body}",
                responseUrl, json);
            var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

            var httpRequest = new HttpRequestMessage(HttpMethod.Post, responseUrl)
            {
                Content = content
            };
            httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken.Token);

            var response = await _httpClient.SendAsync(httpRequest, ct).ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                var errorBody = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
                _logger.LogError("Agent call failed: {StatusCode} {ReasonPhrase}. Endpoint: {Endpoint}. Response body: {Body}", 
                    response.StatusCode, response.ReasonPhrase, responseUrl, errorBody);
                throw new AgentResponsesException(
                    $"Agent call failed with status {response.StatusCode}.",
                    (int)response.StatusCode,
                    errorBody);
            }

            var responseText = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
            var result = JsonSerializer.Deserialize<ResponseBody>(responseText);

            _logger.LogInformation("Agent response received: {Length} chars, response_id={ResponseId}.", 
                result?.output?.Length ?? 0, result?.response_id);

            return new AgentChatReply(
                Reply: result?.output ?? string.Empty,
                ResponseId: result?.response_id ?? string.Empty,
                Status: "success");
        }
        catch (Azure.Identity.AuthenticationFailedException ex)
        {
            _logger.LogError(ex, "OBO token acquisition failed. User assertion may be invalid or permissions insufficient.");
            throw;
        }
        catch (AgentResponsesException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Agent call failed.");
            throw;
        }
    }

    private record ResponseBody(string? output, string? response_id);
}

/// <summary>
/// Represents a non-success response from the Foundry agent endpoint.
/// </summary>
public sealed class AgentResponsesException : Exception
{
    public AgentResponsesException(string message, int statusCode, string body)
        : base(message)
    {
        StatusCode = statusCode;
        Body = body;
    }

    public int StatusCode { get; }
    public string Body { get; }

    /// <summary>
    /// True when the error indicates the continued conversation has a function tool call with no
    /// submitted output — i.e. <c>previous_response_id</c> points at an unfinishable turn.
    /// </summary>
    public bool IsDanglingToolCall =>
        StatusCode == 400 &&
        Body.Contains("No tool output found for function call", StringComparison.OrdinalIgnoreCase);
}
