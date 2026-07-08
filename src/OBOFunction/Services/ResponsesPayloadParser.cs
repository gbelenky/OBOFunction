using OBOFunction.Models;
using System.Text.Json;

namespace OBOFunction.Services;

/// <summary>
/// Tolerant parser for the Foundry/OpenAI Responses payload returned by the hosted-agent endpoint.
/// Handles both normal completions and failed runs that still return HTTP 200.
/// </summary>
public static class ResponsesPayloadParser
{
    public static AgentChatReply ParseReply(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;

        var responseId = GetString(root, "id") ?? GetString(root, "response_id") ?? string.Empty;
        var status = GetString(root, "status") ?? "completed";

        if (string.Equals(status, "failed", StringComparison.OrdinalIgnoreCase))
        {
            var errorMessage =
                root.TryGetProperty("error", out var error) ? GetString(error, "message") : null;

            return new AgentChatReply(
                Reply: string.IsNullOrWhiteSpace(errorMessage) ? "The agent run failed." : errorMessage,
                ResponseId: responseId,
                Status: status);
        }

        var reply =
            GetString(root, "output_text") ??
            ParseOutputArray(root) ??
            string.Empty;

        return new AgentChatReply(
            Reply: reply,
            ResponseId: responseId,
            Status: status);
    }

    private static string? ParseOutputArray(JsonElement root)
    {
        if (!root.TryGetProperty("output", out var output) || output.ValueKind != JsonValueKind.Array)
            return null;

        var parts = new List<string>();

        foreach (var item in output.EnumerateArray())
        {
            if (!item.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Array)
                continue;

            foreach (var contentItem in content.EnumerateArray())
            {
                var type = GetString(contentItem, "type");
                if (string.Equals(type, "output_text", StringComparison.OrdinalIgnoreCase))
                {
                    var text = GetString(contentItem, "text");
                    if (!string.IsNullOrWhiteSpace(text))
                        parts.Add(text);
                    continue;
                }

                var fallbackText = GetString(contentItem, "text");
                if (!string.IsNullOrWhiteSpace(fallbackText))
                    parts.Add(fallbackText);
            }
        }

        return parts.Count == 0 ? null : string.Concat(parts);
    }

    private static string? GetString(JsonElement element, string propertyName) =>
        element.TryGetProperty(propertyName, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;
}
