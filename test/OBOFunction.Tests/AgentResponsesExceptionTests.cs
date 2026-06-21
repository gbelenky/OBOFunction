using OBOFunction.Services;
using Xunit;

namespace OBOFunction.Tests;

/// <summary>
/// Tests for <see cref="AgentResponsesException.IsDanglingToolCall"/> — the signal the proxy uses to
/// auto-recover a continuation whose <c>previous_response_id</c> points at a turn with an unsatisfied
/// function tool call (e.g. leftover state from an earlier agent version with a removed tool).
/// </summary>
public class AgentResponsesExceptionTests
{
    [Fact]
    public void IsDanglingToolCall_True_For400WithNoToolOutputBody()
    {
        var body =
            "{\"error\":{\"type\":\"invalid_request_error\",\"message\":\"\"}}\n" +
            "No tool output found for function call call_LJ6fJ0yWwCOjRGbVVXpZIR7D.";

        var ex = new AgentResponsesException(400, body);

        Assert.True(ex.IsDanglingToolCall);
    }

    [Fact]
    public void IsDanglingToolCall_False_WhenStatusIsNot400()
    {
        var ex = new AgentResponsesException(502, "No tool output found for function call call_x.");

        Assert.False(ex.IsDanglingToolCall);
    }

    [Fact]
    public void IsDanglingToolCall_False_ForOther400Errors()
    {
        var ex = new AgentResponsesException(400, "{\"error\":{\"message\":\"invalid input\"}}");

        Assert.False(ex.IsDanglingToolCall);
    }

    [Fact]
    public void Exception_CarriesStatusAndBody()
    {
        var ex = new AgentResponsesException(400, "boom");

        Assert.Equal(400, ex.StatusCode);
        Assert.Equal("boom", ex.Body);
    }
}
