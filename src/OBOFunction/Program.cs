using Azure.Identity;
using Azure.Monitor.OpenTelemetry.AspNetCore;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using OBOFunction.Auth;
using OBOFunction.Models;
using OBOFunction.Observability;
using OBOFunction.Services;

var builder = WebApplication.CreateBuilder(args);

// ---------------------------------------------------------------------------
// OBOFunction — agent-chat proxy (OBO), hosted on App Service.
//
// The proxy validates the SPFx caller's JWT, extracts the user assertion, and calls the
// Foundry hosted agent AS THAT USER using OnBehalfOfCredential (OBO identity pass-through).
// The agent owns all profile resolution via its Toolbox connection (SharePointProfile).
//   POST /api/agent/chat  -> validated user JWT -> OBO to ai.azure.com -> hosted agent
//                               (agent calls Toolbox for user's profile as needed)
// ---------------------------------------------------------------------------

// Key Vault as a configuration source (secrets referenced by name, MSI to read).
var kvUri = builder.Configuration["KeyVault:Uri"];
if (!string.IsNullOrWhiteSpace(kvUri))
{
    builder.Configuration.AddAzureKeyVault(new Uri(kvUri), new DefaultAzureCredential());
}

// Observability: OpenTelemetry → Azure Monitor (Foundry-native, GenAI semantic conventions).
// Exports to Application Insights via APPLICATIONINSIGHTS_CONNECTION_STRING.
builder.Services.AddOpenTelemetry().UseAzureMonitor();
builder.Services.AddOpenTelemetry().WithTracing(t =>
    t.AddSource(OBOFunction.Observability.AgentTelemetry.SourceName));

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

        body ??= new AgentChatRequest(string.Empty);

        if (string.IsNullOrWhiteSpace(body.Message))
            return Problem("A non-empty 'message' is required.", "Invalid request", StatusCodes.Status400BadRequest);

        var isFirstTurn = string.IsNullOrWhiteSpace(body.PreviousResponseId);

        // One distributed span per chat turn (privacy-safe tags only — no token/PII).
        // Correlates the proxy hop with the agent's server-side traces via the returned responseId.
        using var span = AgentTelemetry.StartChatTurn(isFirstTurn, false, GetUserOid(req));
        if (!string.IsNullOrWhiteSpace(body.PreviousResponseId))
            span?.SetTag(AgentTelemetry.Attr.PreviousResponseId, body.PreviousResponseId);

        // Call the agent as the user (OBO). The agent owns:
        // - Profile resolution (Toolbox connection with per-user OAuth tokens)
        // - Multi-turn conversation state (ProjectConversation)
        // - Tool execution (local search_faq and Toolbox-based profile fetch)
        var reply = await agent.ChatAsync(body, assertion, ct);

        // Tag the outcome so a whole conversation is discoverable by responseId.
        span?.SetTag(AgentTelemetry.Attr.ResponseId, reply.ResponseId);
        span?.SetTag(AgentTelemetry.Attr.ResponseStatus, reply.Status);

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

static string? GetUserOid(HttpRequest req) =>
    req.HttpContext.User?.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier")?.Value
    ?? req.HttpContext.User?.FindFirst("oid")?.Value;

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
