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
// Single architecture: the agent declares the standalone SharePointMcp server as a
// Foundry-native `mcp` tool (HostedMcpServerTool). Foundry resolves and calls the
// MCP server server-side and forwards the caller's identity to it (the user's OBO
// token injected by the proxy for the SPFx path, or the Foundry OAuth connection
// token for the Playground). The MCP server then performs OBO to Graph + SharePoint.
// The agent itself never touches Graph/SharePoint and holds no embedded profile tool.
//
// Hosting uses Microsoft.Agents.AI.Foundry.Hosting (`AddFoundryResponses` /
// `MapFoundryResponses`) so it runs both locally (`azd ai agent run`) and as a
// Foundry hosted agent unchanged.
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

// The agent's ONLY tool is the standalone SharePointMcp server, declared as a
// Foundry-native MCP tool. Foundry calls it server-side and forwards the caller
// identity; the MCP server does the OBO exchange to Graph + SharePoint.
string mcpServerUrl =
    FirstNonBlank(
        Environment.GetEnvironmentVariable("MCP_SERVER_URL"),
        Environment.GetEnvironmentVariable("SHAREPOINT_MCP_URL"))
    ?? throw new InvalidOperationException(
        "MCP_SERVER_URL must be set to the SharePointMcp server /mcp endpoint " +
        "(e.g. https://app-mcp-<token>.azurewebsites.net/mcp).");

// IMPORTANT (verified by decompiling Microsoft.Extensions.AI.OpenAI):
// HostedMcpServerTool is converted to an OpenAI McpTool that can carry EITHER a
// `server_url` (+ optional `authorization` header) OR a built-in `connector_id`.
// It does NOT read AdditionalProperties and CANNOT emit a custom Foundry
// `project_connection_id`. Therefore a hosted (code) agent cannot bind its
// in-process MCP tool to a project "OAuth Identity Passthrough" connection.
//
// User-identity OBO to the custom MCP server is delivered per-call via the
// `authorization` header (the proxy/SPFx path sets it to the user's OBO token).
// If MCP_USER_AUTHORIZATION is present in this process (e.g. a static dev token),
// it is forwarded; otherwise the tool is URL-only and the MCP server falls back to
// its own managed identity (app-only) — correct for local dev, empty profile in
// the autonomous Playground. The SPFx production path does NOT rely on this agent
// for auth; the proxy calls the model Responses endpoint with a per-user
// `authorization` token (see OBOFunction AgentChatClient).
string? mcpUserAuthorization =
    Environment.GetEnvironmentVariable("MCP_USER_AUTHORIZATION");

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

IList<AITool> tools = [mcpTool];

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

AIAgent agent = new AIProjectClient(new Uri(projectEndpoint), credential)
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
