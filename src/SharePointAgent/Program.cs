using Azure.AI.Projects;
using Azure.Identity;
using Microsoft.Agents.AI;
using Microsoft.Agents.AI.Foundry.Hosting;
using Microsoft.Extensions.AI;
using SharePointAgent.Services;
using SharePointAgent.Tools;

// ----------------------------------------------------------------------------
// SharePointAgent — Microsoft Foundry Hosted Agent (.NET 10)
//
// Profile path (CORE OBO goal) — STANDARD FOUNDRY TOOLBOX PATTERN:
//   The SharePointMcp server's `get_sharepoint_profile` tool is reached through a
//   Foundry Toolbox bound to a project OAuth-Identity-Passthrough connection
//   (`project_connection_id`). We register the toolbox with the hosting layer via
//   `AddFoundryToolboxes` (Microsoft.Agents.AI.Foundry.Hosting). The hosting layer
//   connects to the Foundry Toolboxes MCP proxy and injects the toolbox's tools as
//   host-executed MCP tools wrapped in a CONSENT-AWARE function: the first call for
//   a new user surfaces an `McpConsentInfo.ConsentUrl`, Foundry runs the per-user
//   OAuth flow (token acquisition, refresh, consent), and the MCP server then does
//   OBO to Graph + SharePoint UPS and returns the full profile (incl. country). The
//   agent process carries NO auth logic — Foundry resolves the connection credential
//   when it proxies each MCP call. This mirrors the official C# sample
//   `hosted-agents/agent-framework/toolbox-auth-paths`.
//
//   StrictMode is disabled so a consent-pending / transiently-unreachable toolbox
//   source never bricks host startup — the agent still serves the LOCAL search_faq
//   tool, and the profile tool surfaces consent on first per-user invocation.
//
// Fallback (local dev / degraded): if TOOLBOX_NAME is unset, declare the MCP server
//   directly by URL (HostedMcpServerTool). HostedMcpServerTool CANNOT emit
//   `project_connection_id`, so without a per-call user token the MCP server falls
//   back to its own managed identity (app-only) → `profileAvailable: false`.
//
// Hosting uses Microsoft.Agents.AI.Foundry.Hosting (`AddFoundryResponses` /
// `MapFoundryResponses` / `AddFoundryToolboxes`) so it runs both locally
// (`azd ai agent run`) and as a Foundry hosted agent unchanged.
// ----------------------------------------------------------------------------

string projectEndpoint =
    Environment.GetEnvironmentVariable("FOUNDRY_PROJECT_ENDPOINT")
    ?? Environment.GetEnvironmentVariable("AZURE_AI_PROJECT_ENDPOINT")
    ?? throw new InvalidOperationException(
        "FOUNDRY_PROJECT_ENDPOINT (or AZURE_AI_PROJECT_ENDPOINT) must be set.");

string deploymentName =
    FirstNonBlank(
        Environment.GetEnvironmentVariable("AZURE_AI_MODEL_DEPLOYMENT_NAME"),
        Environment.GetEnvironmentVariable("FOUNDRY_MODEL_DEPLOYMENT_NAME"))
    ?? "gpt-4.1-mini";

static string? FirstNonBlank(params string?[] values) =>
    values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));

var credential = new DefaultAzureCredential();
var projectClient = new AIProjectClient(new Uri(projectEndpoint), credential);

// ---- SharePoint profile tool (server-side, per-user OBO via Foundry Toolbox) ----
// When TOOLBOX_NAME is set the profile tool is registered with the HOSTING layer below
// (AddFoundryToolboxes) — NOT added to this in-process `tools` list. When it is unset we
// fall back to a direct HostedMcpServerTool (local dev / degraded app-only).
string? toolboxName =
    FirstNonBlank(
        Environment.GetEnvironmentVariable("TOOLBOX_NAME"),
        Environment.GetEnvironmentVariable("SHAREPOINT_TOOLBOX_NAME"));

string? mcpUserAuthorization =
    Environment.GetEnvironmentVariable("MCP_USER_AUTHORIZATION");

IList<AITool> tools = [];

if (string.IsNullOrWhiteSpace(toolboxName))
{
    // FALLBACK: no toolbox configured — declare the MCP server directly by URL.
    string mcpServerUrl =
        FirstNonBlank(
            Environment.GetEnvironmentVariable("MCP_SERVER_URL"),
            Environment.GetEnvironmentVariable("SHAREPOINT_MCP_URL"))
        ?? throw new InvalidOperationException(
            "Either TOOLBOX_NAME (Foundry Toolbox, recommended for per-user OBO) or " +
            "MCP_SERVER_URL (direct MCP /mcp endpoint) must be set.");

    var mcpTool = new HostedMcpServerTool("SharePointMcp", new Uri(mcpServerUrl))
    {
        AllowedTools = ["get_sharepoint_profile"],
        ApprovalMode = HostedMcpServerToolApprovalMode.NeverRequire,
    };

    if (!string.IsNullOrWhiteSpace(mcpUserAuthorization))
    {
        string headerValue = mcpUserAuthorization.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase)
            ? mcpUserAuthorization
            : $"Bearer {mcpUserAuthorization}";
        mcpTool.Headers ??= new Dictionary<string, string>();
        mcpTool.Headers["Authorization"] = headerValue;
    }

    tools.Add(mcpTool);
    Console.Error.WriteLine(
        $"[startup] TOOLBOX_NAME unset — using direct MCP tool ({mcpServerUrl}, app-only/degraded).");
}

// ---- Local FAQ / Q&A tool (no OBO, no user identity) ----
// The agent owns this tool: it queries Azure AI Search for FAQ entries filtered by the user's
// country (the index `Location` field). It needs only the country string from the profile, so it
// runs with the agent's OWN identity (a read-only query key when configured, otherwise the
// agent's Managed Identity / developer credential). Registered only when an endpoint is set, so
// the agent still runs if the FAQ index is not configured.
string? searchEndpoint =
    FirstNonBlank(
        Environment.GetEnvironmentVariable("AZURE_SEARCH_ENDPOINT"),
        Environment.GetEnvironmentVariable("SEARCH_ENDPOINT"));

if (!string.IsNullOrWhiteSpace(searchEndpoint))
{
    string searchIndex =
        FirstNonBlank(Environment.GetEnvironmentVariable("AZURE_SEARCH_INDEX_NAME")) ?? "faq-index";
    string countryField =
        FirstNonBlank(Environment.GetEnvironmentVariable("AZURE_SEARCH_COUNTRY_FIELD")) ?? "Location";
    string? searchApiKey =
        FirstNonBlank(
            Environment.GetEnvironmentVariable("AZURE_SEARCH_API_KEY"),
            Environment.GetEnvironmentVariable("SEARCH_API_KEY"));
    bool includeGlobal =
        !string.Equals(
            Environment.GetEnvironmentVariable("AZURE_SEARCH_INCLUDE_GLOBAL"),
            "false", StringComparison.OrdinalIgnoreCase);

    var faqService = new FaqSearchService(searchEndpoint, searchIndex, countryField, searchApiKey, includeGlobal);
    var faqTools = new FaqTools(faqService);
    tools.Add(AIFunctionFactory.Create(faqTools.SearchFaqAsync, name: "search_faq"));
}

const string instructions =
    "You are the SharePoint Profile Assistant. " +
    "At the START of every new conversation, FIRST call the `get_sharepoint_profile` tool " +
    "to load the signed-in user's profile, then greet them by name and use that context " +
    "(job title, department, office, skills, interests, responsibilities) in your answers. " +
    "Never ask the user for information already present in their profile. " +
    "If the tool returns null fields, work with what is available and do not invent values. " +
    "When the user asks a policy, IT, HR, finance, facilities or general how-to question — or asks " +
    "which FAQs apply to them — call the `search_faq` tool, passing the user's `country` from their " +
    "profile so results are filtered to their location (globally-applicable entries are always " +
    "included). Answer ONLY from the returned FAQ entries and cite each FAQ's Title; do not invent " +
    "answers. If the profile has no country, ask the user for their country/region, or proceed with " +
    "the globally-applicable entries. " +
    "If the profile tool returns `profileAvailable: false`, explain to the user — clearly and briefly — " +
    "that no signed-in user is present in this context (for example the Foundry Playground runs " +
    "autonomously), so per-user profile data requires invoking the agent with a user token via the " +
    "proxy's POST /api/agent/chat (OBO), an M365/Teams SSO channel, or an MCP OAuth identity-passthrough " +
    "connection. Relay the `detail` field; do not invent profile values. In that case you can still " +
    "answer FAQ questions if the user tells you their country/region.";

AIAgent agent = projectClient
    .AsAIAgent(
        model: deploymentName,
        name: "SharePointProfileAgent",
        instructions: instructions,
        tools: [.. tools]);

string port =
    Environment.GetEnvironmentVariable("DEFAULT_AD_PORT")
    ?? Environment.GetEnvironmentVariable("PORT")
    ?? "8088";

var builder = WebApplication.CreateBuilder(args);
builder.WebHost.UseUrls($"http://+:{port}");
builder.Services.AddFoundryResponses(agent);

// STANDARD FOUNDRY TOOLBOX PATTERN (per-user OBO via OAuth identity passthrough).
// The hosting layer connects to the Foundry Toolboxes MCP proxy, discovers the toolbox's
// tools, and injects them as host-executed, consent-aware MCP tools at request time.
if (!string.IsNullOrWhiteSpace(toolboxName))
{
    // Hosted containers inject FOUNDRY_AGENT_TOOLSET_ENDPOINT (the toolbox MCP proxy base).
    // When absent (local dev), derive it from the project endpoint so AddFoundryToolboxes
    // can still reach the toolbox proxy (mirrors the official toolbox-auth-paths sample).
    if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("FOUNDRY_AGENT_TOOLSET_ENDPOINT")))
    {
        Environment.SetEnvironmentVariable(
            "FOUNDRY_AGENT_TOOLSET_ENDPOINT",
            $"{projectEndpoint.TrimEnd('/')}/toolboxes");
    }

    builder.Services.AddFoundryToolboxes(
        options =>
        {
            options.ApiVersion = "v1";
            // Do NOT brick startup if the toolbox source needs per-user consent or is
            // transiently unreachable — the agent still serves the LOCAL search_faq tool,
            // and the profile tool surfaces its consent URL on first per-user invocation.
            options.StrictMode = false;
        },
        toolboxName);

    Console.Error.WriteLine($"[startup] Registered Foundry Toolbox '{toolboxName}' (consent-aware, StrictMode=false).");
}

var app = builder.Build();
app.MapFoundryResponses();
app.MapGet("/readiness", () => Results.Ok("Ready"));
app.MapGet("/liveness", () => Results.Ok("Healthy"));

app.Run();
