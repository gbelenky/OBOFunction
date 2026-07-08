namespace OBOFunction.Models;

/// <summary>
/// Chat request from the SPFx web part to the agent proxy.
/// <paramref name="Message"/> is the user's query text (or empty string if greeting=true).
/// <paramref name="Greeting"/> if true, triggers an opening greeting from the agent (overrides Message).
/// <paramref name="PreviousResponseId"/> carries the prior turn's conversation id so the run can
/// continue a multi-turn conversation on the same <c>ProjectConversation</c>.
/// </summary>
public sealed record AgentChatRequest(
    string Message,
    bool Greeting = false,
    string? PreviousResponseId = null);

/// <summary>
/// Reply returned to the web part.
/// <paramref name="ResponseId"/> is the conversation id; the client should pass it back as
/// <c>PreviousResponseId</c> on the next turn to continue the conversation.
/// <paramref name="Status"/> is <c>"success"</c> for a normal answer or <c>"failed"</c> when
/// the agent run could not complete.
/// </summary>
public sealed record AgentChatReply(
    string Reply,
    string ResponseId,
    string Status = "success");
