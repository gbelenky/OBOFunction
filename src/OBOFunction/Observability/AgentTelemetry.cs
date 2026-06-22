using System.Diagnostics;

namespace OBOFunction.Observability;

/// <summary>
/// Central <see cref="ActivitySource"/> for the proxy plus helpers to tag a chat-turn span using
/// OpenTelemetry GenAI semantic conventions (https://opentelemetry.io/docs/specs/semconv/gen-ai/).
/// <para>
/// Spans emitted here are exported to Application Insights via Azure Monitor and stitch the proxy
/// (SPFx → proxy → OBO) onto the SAME distributed trace as the Foundry hosted-agent server-side
/// traces — so a single trace covers SPFx → proxy → agent → search_faq.
/// </para>
/// <para>
/// PRIVACY: only non-sensitive correlation values are tagged — never the user JWT, any access
/// token, or raw profile fields (name/email/country). Identity is recorded as the Entra object id
/// (<c>oid</c>) only, and profile resolution as a boolean.
/// </para>
/// </summary>
public static class AgentTelemetry
{
    /// <summary>Name registered with <c>AddSource(...)</c> in <c>Program.cs</c>.</summary>
    public const string SourceName = "OBOFunction.AgentChat";

    public static readonly ActivitySource Source = new(SourceName);

    // GenAI semantic-convention attribute names (stable subset) + a few proxy-specific ones.
    public static class Attr
    {
        public const string Operation = "gen_ai.operation.name";       // e.g. "chat"
        public const string System = "gen_ai.system";                  // e.g. "azure.ai.foundry"
        public const string AgentName = "gen_ai.agent.name";
        public const string ResponseId = "gen_ai.response.id";         // == Foundry responseId
        public const string PreviousResponseId = "gen_ai.previous_response.id";
        public const string ResponseStatus = "gen_ai.response.status";

        // Proxy-specific correlation (privacy-safe).
        public const string UserOid = "enduser.id";                    // Entra oid — NOT name/email
        public const string IsFirstTurn = "chat.is_first_turn";
        public const string IsGreeting = "chat.is_greeting";
        public const string ProfileResolved = "chat.profile_resolved"; // bool, no PII
        public const string RecoveredDangling = "chat.recovered_dangling_tool_call";
    }

    /// <summary>Starts the per-turn span for the chat endpoint.</summary>
    public static Activity? StartChatTurn(bool isFirstTurn, bool isGreeting, string? userOid)
    {
        var act = Source.StartActivity("agent.chat", ActivityKind.Server);
        if (act is null) return null;

        act.SetTag(Attr.Operation, "chat");
        act.SetTag(Attr.System, "azure.ai.foundry");
        act.SetTag(Attr.IsFirstTurn, isFirstTurn);
        act.SetTag(Attr.IsGreeting, isGreeting);
        if (!string.IsNullOrWhiteSpace(userOid))
            act.SetTag(Attr.UserOid, userOid);
        return act;
    }
}
