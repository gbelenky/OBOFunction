using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OBOFunction.Models;
using System.Net.Http.Headers;
using System.Text.Json.Nodes;
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
    private readonly string? _agentResponsesUrl;
    private readonly string _apiVersion;
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

        _agentResponsesUrl = config["Foundry:AgentResponsesUrl"];
        _apiVersion = config["Foundry:ApiVersion"] ?? "v1";

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

            var parsedReply = await SendAsync(request, accessToken.Token, allowFreshRetry: true, ct).ConfigureAwait(false);

            _logger.LogInformation("Agent response received: {Length} chars, response_id={ResponseId}, status={Status}.",
                parsedReply.Reply.Length, parsedReply.ResponseId, parsedReply.Status);

            return new AgentChatReply(
                Reply: parsedReply.Reply,
                ResponseId: parsedReply.ResponseId,
                Status: string.Equals(parsedReply.Status, "failed", StringComparison.OrdinalIgnoreCase) ? "failed" : "success");
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

    private async Task<AgentChatReply> SendAsync(
        AgentChatRequest request,
        string bearerToken,
        bool allowFreshRetry,
        CancellationToken ct)
    {
        var responseUrl = ResolveAgentResponsesUrl();
        var requestBody = new JsonObject
        {
            ["input"] = BuildInput(request)
        };

        if (!string.IsNullOrWhiteSpace(request.PreviousResponseId))
        {
            requestBody["previous_response_id"] = request.PreviousResponseId;
        }

        var json = JsonSerializer.Serialize(requestBody);
        _logger.LogInformation("Calling agent endpoint: {Url} with body: {Body}", responseUrl, json);

        using var httpRequest = new HttpRequestMessage(HttpMethod.Post, responseUrl)
        {
            Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json")
        };
        httpRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
        httpRequest.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        using var response = await _httpClient.SendAsync(httpRequest, ct).ConfigureAwait(false);
        var responseText = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);

        if (!response.IsSuccessStatusCode)
        {
            _logger.LogError("Agent call failed: {StatusCode} {ReasonPhrase}. Endpoint: {Endpoint}. Response body: {Body}",
                response.StatusCode, response.ReasonPhrase, responseUrl, responseText);

            var failure = new AgentResponsesException((int)response.StatusCode, responseText);
            if (allowFreshRetry &&
                !string.IsNullOrWhiteSpace(request.PreviousResponseId) &&
                failure.IsDanglingToolCall)
            {
                _logger.LogWarning("Retrying agent call without previousResponseId after dangling tool-call error.");
                return await SendAsync(request with { PreviousResponseId = null }, bearerToken, allowFreshRetry: false, ct).ConfigureAwait(false);
            }

            throw failure;
        }

        var parsedReply = ResponsesPayloadParser.ParseReply(responseText);
        if (allowFreshRetry &&
            !string.IsNullOrWhiteSpace(request.PreviousResponseId) &&
            IsFailedDanglingToolCall(parsedReply))
        {
            _logger.LogWarning("Retrying agent call without previousResponseId after failed continuation response.");
            return await SendAsync(request with { PreviousResponseId = null }, bearerToken, allowFreshRetry: false, ct).ConfigureAwait(false);
        }

        return parsedReply;
    }

    private string ResolveAgentResponsesUrl()
    {
        if (!string.IsNullOrWhiteSpace(_agentResponsesUrl))
            return _agentResponsesUrl;

        var baseUrl = $"{_projectEndpoint.ToString().TrimEnd('/')}/agents/{Uri.EscapeDataString(_agentName)}/endpoint/protocols/openai/responses";
        return $"{baseUrl}?api-version={Uri.EscapeDataString(_apiVersion)}";
    }

    private static string BuildInput(AgentChatRequest request) =>
        request.Greeting
            ? "The chat was just opened. Greet the signed-in user in one short sentence by first name and offer help. Do not dump their profile."
            : request.Message ?? string.Empty;

    private static bool IsFailedDanglingToolCall(AgentChatReply reply) =>
        string.Equals(reply.Status, "failed", StringComparison.OrdinalIgnoreCase) &&
        reply.Reply.Contains("No tool output found for function call", StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// Represents a non-success response from the Foundry agent endpoint.
/// </summary>
public sealed class AgentResponsesException : Exception
{
    public AgentResponsesException(int statusCode, string body)
        : this($"Agent call failed with status {statusCode}.", statusCode, body)
    {
    }

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
