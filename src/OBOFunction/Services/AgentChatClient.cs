using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OBOFunction.Models;

namespace OBOFunction.Services;

/// <summary>
/// Calls the Foundry HOSTED agent as the signed-in user (OBO identity pass-through).
/// 
/// <para><b>Flow:</b> SPFx (user token) → Proxy (validates JWT) → HTTP POST to Responses endpoint
/// with user's token in Authorization header → Foundry agent → Toolbox (user OBO for profile).</para>
///
/// <para><b>Identity:</b> The user's Bearer token is passed through to the Responses endpoint.
/// The agent's Toolbox connection uses this token to perform OBO calls for the user's
/// SharePoint/Graph profile.</para>
///
/// <para><b>Conversation state:</b> Multi-turn conversations are tied to <c>response_id</c> (conversation ID)
/// which the client stores and passes back as <c>previous_response_id</c> on follow-up turns.</para>
/// </summary>
public sealed class AgentChatClient : IAgentChatClient
{
    private readonly HttpClient _httpClient;
    private readonly Uri _responsesEndpoint;
    private readonly string _agentName;
    private readonly ILogger<AgentChatClient> _logger;

    public AgentChatClient(IConfiguration config, ILogger<AgentChatClient> logger)
    {
        _logger = logger;

        // Foundry Responses endpoint format: https://<resource>.services.ai.azure.com/projects/<projectName>/responses
        string responsesUrl = config["Foundry:ResponsesUrl"]
            ?? throw new InvalidOperationException(
                "Foundry:ResponsesUrl is required (format: https://<resource>.services.ai.azure.com/projects/<projectName>/responses).");
        _responsesEndpoint = new Uri(responsesUrl.TrimEnd('/'));

        _agentName = config["Foundry:AgentName"]
            ?? throw new InvalidOperationException("Foundry:AgentName is required (the hosted agent name/id).");

        _httpClient = new HttpClient();
        _logger.LogInformation("Agent proxy ready: endpoint={Endpoint}, agent={Agent}.", _responsesEndpoint, _agentName);
    }

    public async Task<AgentChatReply> ChatAsync(
        AgentChatRequest request, string? userAssertion = null, CancellationToken ct = default)
    {
        System.Diagnostics.Activity.Current?.SetTag(
            OBOFunction.Observability.AgentTelemetry.Attr.AgentName, _agentName);

        if (string.IsNullOrWhiteSpace(userAssertion))
            throw new UnauthorizedAccessException("The agent proxy requires the signed-in user's token.");

        // POST to the Foundry Responses endpoint. Format:
        // POST https://<resource>.services.ai.azure.com/projects/<projectName>/responses/<agentName>
        var requestUrl = $"{_responsesEndpoint}/{_agentName}";

        var requestBody = new
        {
            message = request.Message,
            previous_response_id = request.PreviousResponseId
        };

        var json = System.Text.Json.JsonSerializer.Serialize(requestBody);
        var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

        // Pass the user's Bearer token in the Authorization header.
        // The agent's Toolbox connection will use this token to perform OBO on behalf of the user.
        var request2 = new HttpRequestMessage(HttpMethod.Post, requestUrl)
        {
            Content = content
        };
        request2.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", userAssertion);

        try
        {
            var response = await _httpClient.SendAsync(request2, ct).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();

            var responseText = await response.Content.ReadAsStringAsync(ct).ConfigureAwait(false);
            var result = System.Text.Json.JsonSerializer.Deserialize<ResponseBody>(responseText);

            _logger.LogInformation("Agent response received: {Length} chars.", result?.output?.Length ?? 0);

            return new AgentChatReply(
                Reply: result?.output ?? string.Empty,
                ResponseId: result?.response_id ?? string.Empty,
                Status: "success");
        }
        catch (HttpRequestException ex)
        {
            _logger.LogError(ex, "Agent HTTP call failed: {RequestUrl}", requestUrl);
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
/// Carries a non-success response from the Foundry agent Responses endpoint so the caller can
/// inspect the status and body (e.g. to detect a dangling-tool-call continuation error and retry).
/// </summary>
public sealed class AgentResponsesException : Exception
{
    public AgentResponsesException(int statusCode, string body)
        : base($"Agent call failed with status {statusCode}.")
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
