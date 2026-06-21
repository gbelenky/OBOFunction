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

// ---- Profile: host-injected (Option A) ----
// The agent no longer calls any `get_sharepoint_profile` MCP/toolbox tool. The proxy
// (POST /api/agent/chat) resolves the signed-in user's profile via OBO and prepends it to
// the first turn as a USER_PROFILE_JSON directive block. The agent treats that as private
// background knowledge. Only the LOCAL search_faq tool runs in-process here.

IList<AITool> tools = [];

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
    "The signed-in user's profile is supplied to you by the host as a USER_PROFILE_JSON directive " +
    "block at the start of the conversation. Treat it as PRIVATE background knowledge for yourself only. " +
    "NEVER repeat, list, quote, summarize, or display the user's profile fields back to them, and NEVER " +
    "append a bulleted profile list to any reply. " +
    "If the user's message is just a greeting, reply with one short sentence that greets them by first " +
    "name and offers help. For any other message, greet by first name in your first sentence, then " +
    "directly answer — you MAY use the profile values (job title, responsibilities, interests, country) " +
    "to inform your answer. Never ask the user for information already present in their profile. " +
    "When the user asks a policy, IT, HR, finance, facilities or general how-to question — or asks " +
    "which FAQs apply to them — call the `search_faq` tool, passing the user's `country` from the " +
    "profile so results are filtered to their location (globally-applicable entries are always " +
    "included). Answer ONLY from the returned FAQ entries and cite each FAQ's Title; do not invent " +
    "answers. FAQ entries may be written in the user's local language (e.g. German 'Urlaubsantrag' or " +
    "French 'Demande de congé' for a vacation/leave request); match the user's question to the most " +
    "relevant entry REGARDLESS of language, and answer in the user's language, translating the entry " +
    "if needed. If the tool result has \"broadened\": true it returned the full FAQ set for the region " +
    "because the keyword search found no direct match — pick the relevant entries from it yourself. " +
    "If the profile has no country, proceed with the globally-applicable entries or ask the " +
    "user for their country/region. " +
    "If no USER_PROFILE_JSON block was provided, answer generally and you may ask the user for their " +
    "country/region to filter FAQs; do not invent profile values.";

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

var app = builder.Build();
app.MapFoundryResponses();
app.MapGet("/readiness", () => Results.Ok("Ready"));
app.MapGet("/liveness", () => Results.Ok("Healthy"));

app.Run();
