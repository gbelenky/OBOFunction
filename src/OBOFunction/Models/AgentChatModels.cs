namespace OBOFunction.Models;

/// <summary>
/// Chat request from the SPFx web part to the agent proxy.
/// <paramref name="PreviousResponseId"/> carries the prior turn's id so the run can
/// continue a multi-turn conversation (server-side state, no thread on the client).
/// <paramref name="SystemContext"/> is set server-side only (never by the client): the
/// proxy puts the signed-in user's profile here so it is delivered to the agent as a
/// separate developer-role input item rather than mixed into the user's message — this
/// prevents the model from echoing the raw profile back on a simple greeting.
/// <paramref name="Greeting"/> lets the (agnostic) client request the opening turn without
/// prescribing any wording: it just signals "the chat was opened — greet me". The proxy
/// owns the greeting trigger and the agent owns the greeting text, so the front-end carries
/// no greeting/profile logic. When true, <paramref name="Message"/> may be empty.
/// </summary>
public sealed record AgentChatRequest(
    string Message,
    string? PreviousResponseId = null,
    string? SystemContext = null,
    bool Greeting = false);

/// <summary>
/// Reply returned to the web part. <paramref name="Status"/> is <c>"completed"</c> for a
/// normal answer or <c>"failed"</c> when the agent run could not complete.
/// </summary>
public sealed record AgentChatReply(
    string Reply,
    string ResponseId,
    string Status = "completed");
