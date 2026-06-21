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
    /// Parses a Responses payload. Surfaces a failed run's error message; otherwise prefers
    /// <c>output_text</c>, falling back to <c>output[].content[].text</c>.
    /// </summary>
    public static AgentChatReply ParseReply(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        var id = root.TryGetProperty("id", out var idEl) ? idEl.GetString() ?? "" : "";

        // --- Surface a failed run (e.g. a tool error) instead of returning an empty reply. ---
        if (root.TryGetProperty("status", out var st) && st.ValueKind == JsonValueKind.String
            && string.Equals(st.GetString(), "failed", StringComparison.OrdinalIgnoreCase))
        {
            var message = "The agent run failed.";
            if (root.TryGetProperty("error", out var err) && err.ValueKind == JsonValueKind.Object
                && err.TryGetProperty("message", out var em) && em.ValueKind == JsonValueKind.String)
            {
                var s = em.GetString();
                if (!string.IsNullOrWhiteSpace(s)) message = s!;
            }
            return new AgentChatReply(message, id, Status: "failed");
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
}
