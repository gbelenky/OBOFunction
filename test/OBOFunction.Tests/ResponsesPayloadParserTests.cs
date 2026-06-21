using OBOFunction.Services;
using Xunit;

namespace OBOFunction.Tests;

/// <summary>
/// Tests for <see cref="ResponsesPayloadParser"/> — the fragile, preview-surface parsing of
/// Foundry Responses payloads (normal completion + failed-run recovery). These cover the
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

    // ---------- failed run ----------

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
}
