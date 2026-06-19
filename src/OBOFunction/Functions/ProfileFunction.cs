using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using OBOFunction.Auth;
using OBOFunction.Services;

namespace OBOFunction.Functions;

public sealed class ProfileFunction
{
    private readonly IGraphProfileService _graph;
    private readonly IUserTokenAccessor _tokens;
    private readonly ILogger<ProfileFunction> _logger;

    public ProfileFunction(IGraphProfileService graph, IUserTokenAccessor tokens, ILogger<ProfileFunction> logger)
    {
        _graph = graph;
        _tokens = tokens;
        _logger = logger;
    }

    [Function("GetMyProfile")]
    [Authorize]
    public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "profile")] HttpRequest req,
        CancellationToken ct)
    {
        try
        {
            var assertion = _tokens.GetBearerToken(req);
            var profile = await _graph.GetMyProfileAsync(assertion, ct);
            return new OkObjectResult(profile);
        }
        catch (UnauthorizedAccessException ex)
        {
            _logger.LogWarning(ex, "Unauthorized profile request.");
            return new ObjectResult(new ProblemDetails
            {
                Status = StatusCodes.Status401Unauthorized,
                Title = "Unauthorized",
                Detail = ex.Message
            }) { StatusCode = StatusCodes.Status401Unauthorized };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Profile fetch failed.");
            return new ObjectResult(new ProblemDetails
            {
                Status = StatusCodes.Status502BadGateway,
                Title = "Graph call failed",
                Detail = ex.Message
            }) { StatusCode = StatusCodes.Status502BadGateway };
        }
    }
}
