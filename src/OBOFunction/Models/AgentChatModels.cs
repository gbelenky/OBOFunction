namespace OBOFunction.Models;

/// <summary>
/// Chat request from the SPFx web part to the agent proxy.
/// <paramref name="PreviousResponseId"/> carries the prior turn's id so the run can
/// continue a multi-turn conversation (server-side state, no thread on the client).
/// <paramref name="UserProfile"/> is optional and normally unused: the proxy already
/// gives the agent the real signed-in user's profile via the per-user MCP tool, so the
/// client does not need to pass it. Kept for callers that want to supply extra context.
/// <paramref name="SystemContext"/> is set server-side only (never by the client): the
/// proxy puts the signed-in user's profile here so it is delivered to the agent as a
/// separate developer-role input item rather than mixed into the user's message — this
/// prevents the model from echoing the raw profile back on a simple greeting.
/// </summary>
public sealed record AgentChatRequest(
    string Message,
    string? PreviousResponseId = null,
    object? UserProfile = null,
    string? SystemContext = null);

/// <summary>
/// Reply returned to the web part.
/// <para>
/// <paramref name="Status"/> is <c>"completed"</c> for a normal answer, or
/// <c>"consent_required"</c> when an OAuth identity-passthrough MCP tool needs the signed-in
/// user to grant delegated consent.
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
