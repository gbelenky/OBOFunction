using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using OBOFunction.Auth;
using OBOFunction.Models;
using OBOFunction.Services;

namespace OBOFunction.Functions;

/// <summary>
/// Server-side proxy that lets an SPFx web part chat with the SharePoint hosted agent.
/// The web part authenticates with its user token (validated here, same audience as
/// <c>GET /api/profile</c>); this proxy forwards the message to the hosted agent using the
/// Function's managed identity, so no agent credential reaches the browser.
/// </summary>
public sealed class AgentChatFunction
{
    private readonly IAgentChatClient _agent;
    private readonly IUserTokenAccessor _tokens;
    private readonly ILogger<AgentChatFunction> _logger;

    public AgentChatFunction(IAgentChatClient agent, IUserTokenAccessor tokens, ILogger<AgentChatFunction> logger)
    {
        _agent = agent;
        _tokens = tokens;
        _logger = logger;
    }

    [Function("AgentChat")]
    [Authorize]
    public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "agent/chat")] HttpRequest req,
        CancellationToken ct)
    {
        try
        {
            // Enforce a validated SharePoint user token before invoking the agent.
            var assertion = _tokens.GetBearerToken(req);

            var body = await req.ReadFromJsonAsync<AgentChatRequest>(cancellationToken: ct);
            if (body is null || string.IsNullOrWhiteSpace(body.Message))
            {
                return new BadRequestObjectResult(new ProblemDetails
                {
                    Status = StatusCodes.Status400BadRequest,
                    Title = "Invalid request",
                    Detail = "A non-empty 'message' is required."
                });
            }

            var reply = await _agent.ChatAsync(body, assertion, ct);
            return new OkObjectResult(reply);
        }
        catch (UnauthorizedAccessException ex)
        {
            _logger.LogWarning(ex, "Unauthorized agent chat request.");
            return new ObjectResult(new ProblemDetails
            {
                Status = StatusCodes.Status401Unauthorized,
                Title = "Unauthorized",
                Detail = ex.Message
            }) { StatusCode = StatusCodes.Status401Unauthorized };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Agent chat failed.");
            return new ObjectResult(new ProblemDetails
            {
                Status = StatusCodes.Status502BadGateway,
                Title = "Agent call failed",
                Detail = ex.Message
            }) { StatusCode = StatusCodes.Status502BadGateway };
        }
    }
}
