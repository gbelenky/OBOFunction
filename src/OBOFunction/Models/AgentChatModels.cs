namespace OBOFunction.Models;

/// <summary>
/// Chat request from the SPFx web part to the agent proxy.
/// <paramref name="PreviousResponseId"/> carries the prior turn's id so the hosted
/// agent can continue a multi-turn conversation (server-side state, no thread on the client).
/// <paramref name="UserProfile"/> is optional: when supplied (e.g. fetched via
/// <c>GET /api/profile</c>), the proxy injects it as context so the agent reasons over the
/// real signed-in user instead of its own managed-identity profile.
/// </summary>
public sealed record AgentChatRequest(
    string Message,
    string? PreviousResponseId = null,
    object? UserProfile = null);

/// <summary>
/// Reply returned to the web part.
/// <para>
/// <paramref name="Status"/> is <c>"completed"</c> for a normal answer, or
/// <c>"consent_required"</c> when a Shape 2 OAuth identity-passthrough tool (Foundry
/// toolbox / managed MCP server) needs the signed-in user to grant delegated consent.
/// In that case <paramref name="ConsentUrl"/> holds the one-time consent link the SPFx
/// web part must open; after the user consents, the web part re-sends the same message
/// with <see cref="AgentChatRequest.PreviousResponseId"/> set to <paramref name="ResponseId"/>
/// to resume the run.
/// </para>
/// </summary>
public sealed record AgentChatReply(
    string Reply,
    string ResponseId,
    string Status = "completed",
    string? ConsentUrl = null);
