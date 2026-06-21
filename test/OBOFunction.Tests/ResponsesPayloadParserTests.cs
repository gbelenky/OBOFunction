using OBOFunction.Services;
using Xunit;

namespace OBOFunction.Tests;

/// <summary>
/// Tests for <see cref="ResponsesPayloadParser"/> — the fragile, preview-surface parsing of
/// Foundry Responses payloads, including the OAuth consent ceremony. These cover the
/// paths that have never executed at runtime, so the parsing tolerance is the contract under test.
/// </summary>
public class ResponsesPayloadParserTests
{
    // ---------- normal completion ----------

    [Fact]
    public void ParseReply_UsesOutputText_WhenPresent()
    {
        var json = """{"id":"resp_1","output_text":"Hello Genady"}""";

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("Hello Genady", reply.Reply);
        Assert.Equal("resp_1", reply.ResponseId);
        Assert.Equal("completed", reply.Status);
        Assert.Null(reply.ConsentUrl);
    }

    [Fact]
    public void ParseReply_ConcatenatesOutputContentText_WhenNoOutputText()
    {
        var json = """
        {
          "id":"resp_2",
          "output":[
            {"content":[{"type":"output_text","text":"Hel"},{"type":"output_text","text":"lo"}]}
          ]
        }
        """;

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("Hello", reply.Reply);
        Assert.Equal("resp_2", reply.ResponseId);
        Assert.Equal("completed", reply.Status);
    }

    [Fact]
    public void ParseReply_FallsBackToBareTextProperty_InContent()
    {
        var json = """
        {
          "id":"resp_3",
          "output":[ {"content":[{"text":"plain text"}]} ]
        }
        """;

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("plain text", reply.Reply);
    }

    [Fact]
    public void ParseReply_ReturnsEmptyReply_WhenNoTextAnywhere()
    {
        var json = """{"id":"resp_4","output":[]}""";

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("", reply.Reply);
        Assert.Equal("resp_4", reply.ResponseId);
        Assert.Equal("completed", reply.Status);
    }

    [Fact]
    public void ParseReply_HandlesMissingId()
    {
        var json = """{"output_text":"hi"}""";

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("hi", reply.Reply);
        Assert.Equal("", reply.ResponseId);
    }

    // ---------- consent ceremony ----------

    [Theory]
    [InlineData("oauth_consent_request", "consent_link")]
    [InlineData("mcp_oauth_consent", "consent_url")]
    [InlineData("mcp_approval_request", "approval_url")]
    [InlineData("consent", "url")]
    public void ParseReply_DetectsConsent_AcrossTypeAndLinkVariants(string type, string linkKey)
    {
        var json = $$"""
        {
          "id":"resp_consent",
          "output":[
            {"type":"{{type}}","{{linkKey}}":"https://login.microsoftonline.com/consent/abc"}
          ]
        }
        """;

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("consent_required", reply.Status);
        Assert.Equal("https://login.microsoftonline.com/consent/abc", reply.ConsentUrl);
        Assert.Equal("resp_consent", reply.ResponseId);
        Assert.Contains("consent link", reply.Reply, System.StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ParseReply_DetectsConsent_FromNestedActionUrl()
    {
        var json = """
        {
          "id":"resp_consent2",
          "output":[
            {"type":"oauth_consent_request","action":{"url":"https://consent.example/xyz"}}
          ]
        }
        """;

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("consent_required", reply.Status);
        Assert.Equal("https://consent.example/xyz", reply.ConsentUrl);
    }

    [Fact]
    public void ParseReply_ConsentTakesPrecedence_OverTextOutput()
    {
        // A run can carry both a partial message and a pending consent item; consent must win
        // so SPFx drives the ceremony rather than showing a half-answer.
        var json = """
        {
          "id":"resp_mixed",
          "output":[
            {"content":[{"type":"output_text","text":"partial"}]},
            {"type":"oauth_consent_request","consent_link":"https://consent.example/win"}
          ]
        }
        """;

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("consent_required", reply.Status);
        Assert.Equal("https://consent.example/win", reply.ConsentUrl);
    }

    [Fact]
    public void ParseReply_IgnoresConsentTypeWithoutLink()
    {
        // A consent-typed item with no usable link must NOT be treated as a consent ceremony
        // (otherwise SPFx would open an empty URL); fall through to normal parsing.
        var json = """
        {
          "id":"resp_nolink",
          "output":[
            {"type":"oauth_consent_request"},
            {"content":[{"type":"output_text","text":"answer"}]}
          ]
        }
        """;

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("completed", reply.Status);
        Assert.Null(reply.ConsentUrl);
        Assert.Equal("answer", reply.Reply);
    }

    [Fact]
    public void ParseReply_SurfacesFailedRun_WithErrorMessage()
    {
        // The agent endpoint returns HTTP 200 with a status:"failed" body when a continuation
        // points at a turn whose tool call never got its output. The parser must surface the error
        // message (not an empty reply) so the endpoint can detect it and restart fresh.
        var json = """
        {
          "id":"resp_failed",
          "status":"failed",
          "error":{"message":"No tool output found for function call call_abc123."}
        }
        """;

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("failed", reply.Status);
        Assert.Equal("resp_failed", reply.ResponseId);
        Assert.Contains("No tool output found for function call", reply.Reply);
    }

    [Fact]
    public void ParseReply_FailedRunWithoutErrorMessage_UsesGenericText()
    {
        var json = """{"id":"resp_failed2","status":"failed"}""";

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("failed", reply.Status);
        Assert.Equal("The agent run failed.", reply.Reply);
    }

    [Fact]
    public void ParseReply_DoesNotFalsePositive_OnNonConsentType()
    {
        // A normal item that happens to carry a "url" must not be mistaken for consent.
        var json = """
        {
          "id":"resp_normal",
          "output":[
            {"type":"message","url":"https://example.com/not-consent",
             "content":[{"type":"output_text","text":"hi"}]}
          ]
        }
        """;

        var reply = ResponsesPayloadParser.ParseReply(json);

        Assert.Equal("completed", reply.Status);
        Assert.Null(reply.ConsentUrl);
        Assert.Equal("hi", reply.Reply);
    }
}
