using OBOFunction.Models;

namespace OBOFunction.Services;

/// <summary>
/// Server-side client that invokes the Foundry hosted agent (via <c>agent_reference</c>) as the
/// signed-in user. Keeps all agent credentials off the browser. The agent owns its tools; the
/// proxy stays tool-agnostic.
/// </summary>
public interface IAgentChatClient
{
    /// <param name="request">The chat turn from the SPFx web part (may carry injected profile context).</param>
    /// <param name="userAssertion">
    /// The validated inbound SPFx user token. It is OBO-exchanged for a Foundry data-plane token
    /// (<c>https://ai.azure.com/.default</c>) so the agent call executes as the real signed-in user.
    /// Required — the proxy rejects the call without it.
    /// </param>
    Task<AgentChatReply> ChatAsync(AgentChatRequest request, string? userAssertion = null, CancellationToken ct = default);
}
