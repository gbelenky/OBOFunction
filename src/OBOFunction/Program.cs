using Azure.Identity;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using OBOFunction.Auth;
using OBOFunction.Models;
using OBOFunction.Services;

var builder = WebApplication.CreateBuilder(args);

// ---------------------------------------------------------------------------
// OBOFunction — agent-chat proxy (OBO), hosted on App Service.
//
// The proxy ONLY delegates profile data retrieval to the Foundry agent: it OBO-exchanges
// the inbound SPFx user token to the SharePointMcp audience and attaches it to the agent's
// `mcp` tool, so the agent (model + MCP) reads the profile AS THE USER. The proxy never
// reads Graph/SharePoint itself.
//   POST /api/agent/chat  -> validated user JWT -> OBO -> Foundry model + per-user MCP tool
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
builder.Services.AddSingleton<IAgentChatClient, AgentChatClient>();

var app = builder.Build();

app.UseCors(CorsPolicy);
app.UseAuthentication();
app.UseAuthorization();

// Health probes (anonymous) for App Service / load balancers.
app.MapGet("/liveness", () => Results.Ok("Healthy")).AllowAnonymous();
app.MapGet("/readiness", () => Results.Ok("Ready")).AllowAnonymous();

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
