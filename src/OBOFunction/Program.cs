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
// Option A: resolve the signed-in user's slim profile via OBO so the proxy can inject it as context.
builder.Services.AddSingleton<IProfileContextService, ProfileContextService>();

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
    IProfileContextService profileContext,
    ILoggerFactory loggerFactory,
    CancellationToken ct) =>
{
    var logger = loggerFactory.CreateLogger("AgentChatEndpoint");
    try
    {
        var assertion = tokens.GetBearerToken(req);

        if (body is null || string.IsNullOrWhiteSpace(body.Message))
            return Problem("A non-empty 'message' is required.", "Invalid request", StatusCodes.Status400BadRequest);

        // Option A — inject the signed-in user's slim profile as conversation context on the FIRST
        // turn only (follow-ups inherit it via previous_response_id). The proxy stays unaware of the
        // agent's tools; it only prepends profile data the agent uses to greet the user by name and
        // to drive country-filtered features (search_faq). Best-effort: empty profile is omitted.
        var outgoing = body;
        if (string.IsNullOrWhiteSpace(body.PreviousResponseId))
        {
            var profile = await profileContext.GetProfileAsync(assertion, ct);
            if (profile.HasAny)
            {
                outgoing = body with { Message = BuildProfileContext(profile) + body.Message };
                logger.LogInformation("Injected profile context into the first agent turn.");
            }
            else
            {
                logger.LogInformation("No profile resolved for the user; proceeding without profile context.");
            }
        }

        var reply = await agent.ChatAsync(outgoing, assertion, ct);
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

// Builds the host-provided profile context block prepended to the first agent turn. The agent uses
// it to greet the user by name and to drive country-filtered features (search_faq) without asking.
static string BuildProfileContext(OBOFunction.Models.UserProfileContext p)
{
    var greetingName = !string.IsNullOrWhiteSpace(p.FirstName) ? p.FirstName : p.Name;

    // Emit the profile as a single-line JSON object (NOT a markdown list) so the model is far less
    // likely to parrot it back. The surrounding directive forbids any echo/display of the data.
    var fields = new List<string>();
    if (!string.IsNullOrWhiteSpace(greetingName))
        fields.Add($"\"firstName\":{JsonString(greetingName)}");
    if (!string.IsNullOrWhiteSpace(p.Name))
        fields.Add($"\"name\":{JsonString(p.Name)}");
    if (!string.IsNullOrWhiteSpace(p.Email))
        fields.Add($"\"email\":{JsonString(p.Email)}");
    if (!string.IsNullOrWhiteSpace(p.JobTitle))
        fields.Add($"\"jobTitle\":{JsonString(p.JobTitle)}");
    if (p.Responsibilities.Count > 0)
        fields.Add($"\"responsibilities\":{JsonArray(p.Responsibilities)}");
    if (p.PastProjects.Count > 0)
        fields.Add($"\"pastProjects\":{JsonArray(p.PastProjects)}");
    if (p.Interests.Count > 0)
        fields.Add($"\"interests\":{JsonArray(p.Interests)}");
    if (!string.IsNullOrWhiteSpace(p.Country))
        fields.Add($"\"country\":{JsonString(p.Country)}");

    var json = "{" + string.Join(",", fields) + "}";

    return
        "SYSTEM/HOST DIRECTIVE — NOT a user message. The signed-in user's profile is provided below as " +
        "JSON metadata about the signed-in user.\n" +
        "RULES (follow exactly):\n" +
        "1. Do NOT volunteer or dump this profile UNSOLICITED, and NEVER append a bulleted profile list to " +
        "a greeting or to an unrelated answer.\n" +
        "2. If the user EXPLICITLY asks about their own profile (e.g. \"what is my profile?\", \"show my " +
        "details\", \"what's my job title/country/interests?\"), DO answer directly using these values — " +
        "it is the user asking about THEIR OWN data, so never refuse for privacy reasons.\n" +
        "3. If the user's message is just a greeting, reply with one short sentence that greets them by " +
        "first name and offers help (e.g. \"Hello " + (greetingName ?? "there") + "! How can I help you today?\").\n" +
        "4. For any other message, greet by first name in your first sentence, then DIRECTLY answer — you " +
        "MAY use the profile field values to inform your answer.\n" +
        "5. Silently use \"country\" as authoritative for country-filtered features such as search_faq.\n" +
        "6. Never ask the user for their name or country.\n" +
        "USER_PROFILE_JSON=" + json + "\n\n" +
        "The user's actual message follows:\n";
}

static string JsonString(string s) => System.Text.Json.JsonSerializer.Serialize(s);

static string JsonArray(IEnumerable<string> items) => System.Text.Json.JsonSerializer.Serialize(items);

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
