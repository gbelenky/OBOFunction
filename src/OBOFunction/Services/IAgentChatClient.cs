using OBOFunction.Models;

namespace OBOFunction.Services;

/// <summary>
/// Server-side client that invokes the Foundry model Responses endpoint with the
/// SharePointMcp server attached as a per-user <c>mcp</c> tool. Keeps all agent
/// credentials off the browser.
/// </summary>
public interface IAgentChatClient
{
    /// <param name="request">The chat turn from the SPFx web part.</param>
    /// <param name="userAssertion">
    /// The validated inbound SPFx user token. It is OBO-exchanged for a token whose audience is
    /// the MCP server, attached to the <c>mcp</c> tool's <c>authorization</c> field so the MCP
    /// server (and its downstream Graph/SharePoint OBO) executes as the real signed-in user.
    /// Required — the proxy rejects the call without it.
    /// </param>
    Task<AgentChatReply> ChatAsync(AgentChatRequest request, string? userAssertion = null, CancellationToken ct = default);
}
