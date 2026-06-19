using Azure.Identity;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using OBOFunction.Auth;
using OBOFunction.Models;
using OBOFunction.Services;

var builder = WebApplication.CreateBuilder(args);

// ---------------------------------------------------------------------------
// OBOFunction — SharePoint Profile API (Shape 1 OBO proxy), hosted on App Service.
//
// Refactored from an Azure Functions isolated worker to a plain ASP.NET Core Web API
// so it can run as an App Service Web App alongside the SharePointMcp API. Endpoints,
// auth (Microsoft.Identity.Web JWT + MSAL OBO), and DTOs are unchanged:
//   GET  /api/profile     -> validated user JWT -> OBO -> Graph /me + SharePoint UPS
//   POST /api/agent/chat  -> validated user JWT -> Foundry hosted agent (managed identity)
// ---------------------------------------------------------------------------

// Key Vault as a configuration source (secrets referenced by name, MSI to read).
var kvUri = builder.Configuration["KeyVault:Uri"];
if (!string.IsNullOrWhiteSpace(kvUri))
{
    builder.Configuration.AddAzureKeyVault(new Uri(kvUri), new DefaultAzureCredential());
}

// Observability: Application Insights (connection string from config/env).
builder.Services.AddApplicationInsightsTelemetry();

// AuthN/Z: validate the inbound SPFx user JWT (issuer + audience = api://<client-id>).
builder.Services
    .AddMicrosoftIdentityWebApiAuthentication(builder.Configuration, "AzureAd")
    .EnableTokenAcquisitionToCallDownstreamApi()
    .AddInMemoryTokenCaches();
builder.Services.AddAuthorization();

// CORS: allow only the customer's SharePoint Online origin.
var sharePointOrigin = ResolveSharePointOrigin(builder.Configuration);
const string CorsPolicy = "SharePointOrigin";
builder.Services.AddCors(o => o.AddPolicy(CorsPolicy, p =>
{
    if (sharePointOrigin is not null)
        p.WithOrigins(sharePointOrigin).AllowAnyHeader().AllowAnyMethod().AllowCredentials();
}));

builder.Services.AddScoped<IUserTokenAccessor, UserTokenAccessor>();
builder.Services.AddScoped<IGraphProfileService, GraphProfileService>();
builder.Services.AddSingleton<IAgentChatClient, AgentChatClient>();

var app = builder.Build();

app.UseCors(CorsPolicy);
app.UseAuthentication();
app.UseAuthorization();

// Health probes (anonymous) for App Service / load balancers.
app.MapGet("/liveness", () => Results.Ok("Healthy")).AllowAnonymous();
app.MapGet("/readiness", () => Results.Ok("Ready")).AllowAnonymous();

// GET /api/profile — the signed-in user's Graph + SharePoint profile (OBO).
app.MapGet("/api/profile", [Authorize] async (
    HttpRequest req,
    IGraphProfileService graph,
    IUserTokenAccessor tokens,
    ILoggerFactory loggerFactory,
    CancellationToken ct) =>
{
    var logger = loggerFactory.CreateLogger("ProfileEndpoint");
    try
    {
        var assertion = tokens.GetBearerToken(req);
        var profile = await graph.GetMyProfileAsync(assertion, ct);
        return Results.Ok(profile);
    }
    catch (UnauthorizedAccessException ex)
    {
        logger.LogWarning(ex, "Unauthorized profile request.");
        return Problem(ex.Message, "Unauthorized", StatusCodes.Status401Unauthorized);
    }
    catch (Exception ex)
    {
        logger.LogError(ex, "Profile fetch failed.");
        return Problem(ex.Message, "Graph call failed", StatusCodes.Status502BadGateway);
    }
});

// POST /api/agent/chat — server-side proxy to the Foundry hosted agent.
app.MapPost("/api/agent/chat", [Authorize] async (
    HttpRequest req,
    [FromBody] AgentChatRequest? body,
    IAgentChatClient agent,
    IUserTokenAccessor tokens,
    ILoggerFactory loggerFactory,
    CancellationToken ct) =>
{
    var logger = loggerFactory.CreateLogger("AgentChatEndpoint");
    try
    {
        var assertion = tokens.GetBearerToken(req);

        if (body is null || string.IsNullOrWhiteSpace(body.Message))
            return Problem("A non-empty 'message' is required.", "Invalid request", StatusCodes.Status400BadRequest);

        var reply = await agent.ChatAsync(body, assertion, ct);
        return Results.Ok(reply);
    }
    catch (UnauthorizedAccessException ex)
    {
        logger.LogWarning(ex, "Unauthorized agent chat request.");
        return Problem(ex.Message, "Unauthorized", StatusCodes.Status401Unauthorized);
    }
    catch (Exception ex)
    {
        logger.LogError(ex, "Agent chat failed.");
        return Problem(ex.Message, "Agent call failed", StatusCodes.Status502BadGateway);
    }
});

app.Run();

static IResult Problem(string detail, string title, int status) =>
    Results.Problem(detail: detail, title: title, statusCode: status);

// Derives the SharePoint Online origin (https://<tenant>.sharepoint.com) for the CORS
// allow-list from SharePoint:TenantHostname, falling back to SharePoint:RootSiteUrl.
static string? ResolveSharePointOrigin(IConfiguration config)
{
    var host = config["SharePoint:TenantHostname"];
    if (!string.IsNullOrWhiteSpace(host))
        return $"https://{host.Trim().TrimEnd('/')}";

    var root = config["SharePoint:RootSiteUrl"];
    if (!string.IsNullOrWhiteSpace(root) && Uri.TryCreate(root, UriKind.Absolute, out var uri))
        return uri.GetLeftPart(UriPartial.Authority);

    return null;
}
