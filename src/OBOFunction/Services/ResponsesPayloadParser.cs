using System.Text;
using System.Text.Json;
using OBOFunction.Models;

namespace OBOFunction.Services;

/// <summary>
/// Pure, dependency-free parsing of a Foundry Responses payload into an
/// <see cref="AgentChatReply"/>. Kept separate from <see cref="AgentChatClient"/> (which owns the
/// HTTP/auth concerns) so the fragile, preview-surface parsing can be unit-tested without
/// constructing a credentialed client.
/// </summary>
public static class ResponsesPayloadParser
{
    /// <summary>
    /// Parses a Responses payload. Detects an OAuth/MCP consent request and surfaces its
    /// link; otherwise prefers <c>output_text</c>, falling back to <c>output[].content[].text</c>.
    /// </summary>
    public static AgentChatReply ParseReply(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        var id = root.TryGetProperty("id", out var idEl) ? idEl.GetString() ?? "" : "";

        // --- Surface a consent ceremony if the run is awaiting delegated consent. ---
        if (TryFindConsentLink(root, out var consentUrl))
        {
            return new AgentChatReply(
                Reply: "Additional sign-in is required to access your profile. " +
                       "Open the consent link, approve access, then resend your message.",
                ResponseId: id,
                Status: "consent_required",
                ConsentUrl: consentUrl);
        }

        if (root.TryGetProperty("output_text", out var ot) && ot.ValueKind == JsonValueKind.String)
        {
            return new AgentChatReply(ot.GetString() ?? "", id);
        }

        var sb = new StringBuilder();
        if (root.TryGetProperty("output", out var output) && output.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in output.EnumerateArray())
            {
                if (!item.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (var c in content.EnumerateArray())
                {
                    if (c.TryGetProperty("text", out var text) && text.ValueKind == JsonValueKind.String)
                        sb.Append(text.GetString());
                    else if (c.TryGetProperty("type", out var t) && t.GetString() == "output_text"
                             && c.TryGetProperty("text", out var t2))
                        sb.Append(t2.GetString());
                }
            }
        }

        return new AgentChatReply(sb.ToString(), id);
    }

    /// <summary>
    /// Scans the Responses <c>output[]</c> for an OAuth identity-passthrough consent request and
    /// extracts its link. The preview surface is not finalised, so this is intentionally tolerant:
    /// it matches any output item whose <c>type</c> mentions "consent" or "approval" and pulls the
    /// first link-shaped string (<c>consent_link</c>, <c>consent_url</c>, <c>url</c>,
    /// <c>approval_url</c>, or a nested <c>action.url</c>).
    /// </summary>
    public static bool TryFindConsentLink(JsonElement root, out string consentUrl)
    {
        consentUrl = string.Empty;
        if (!root.TryGetProperty("output", out var output) || output.ValueKind != JsonValueKind.Array)
            return false;

        foreach (var item in output.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object) continue;

            var type = item.TryGetProperty("type", out var t) ? t.GetString() ?? "" : "";
            var looksLikeConsent =
                type.Contains("consent", StringComparison.OrdinalIgnoreCase) ||
                type.Contains("approval", StringComparison.OrdinalIgnoreCase);
            if (!looksLikeConsent) continue;

            foreach (var key in new[] { "consent_link", "consent_url", "url", "approval_url" })
            {
                if (item.TryGetProperty(key, out var v) && v.ValueKind == JsonValueKind.String)
                {
                    var s = v.GetString();
                    if (!string.IsNullOrWhiteSpace(s)) { consentUrl = s!; return true; }
                }
            }

            if (item.TryGetProperty("action", out var action) && action.ValueKind == JsonValueKind.Object
                && action.TryGetProperty("url", out var au) && au.ValueKind == JsonValueKind.String)
            {
                var s = au.GetString();
                if (!string.IsNullOrWhiteSpace(s)) { consentUrl = s!; return true; }
            }
        }

        return false;
    }
}
