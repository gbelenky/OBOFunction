using OBOFunction.Models;

namespace OBOFunction.Services;

/// <summary>
/// Server-side client that invokes the deployed SharePoint hosted agent over the
/// Foundry Responses API using the Function's managed identity. Keeps all agent
/// credentials off the browser.
/// </summary>
public interface IAgentChatClient
{
    /// <param name="request">The chat turn from the SPFx web part.</param>
    /// <param name="userAssertion">
    /// The validated inbound SPFx user token. Used only by the Shape 2 path: it is exchanged
    /// (OBO) for an AI-audience user token so the Responses run — and therefore the Foundry
    /// toolbox's OAuth identity passthrough — executes as the real signed-in user. Ignored by
    /// the default Shape 1 path, which calls the agent with the Function's managed identity.
    /// </param>
    Task<AgentChatReply> ChatAsync(AgentChatRequest request, string? userAssertion = null, CancellationToken ct = default);
}
