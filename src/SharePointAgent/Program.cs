using Azure.AI.Projects;
using Azure.Identity;
using Microsoft.Agents.AI;
using Microsoft.Agents.AI.Foundry.Hosting;
using Microsoft.Extensions.AI;

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

var mcpTool = new HostedMcpServerTool("SharePointMcp", new Uri(mcpServerUrl))
{
    AllowedTools = ["get_sharepoint_profile"],
    ApprovalMode = HostedMcpServerToolApprovalMode.NeverRequire,
};

IList<AITool> tools = [mcpTool];

const string instructions =
    "You are the SharePoint Profile Assistant. " +
    "At the START of every new conversation, FIRST call the `get_sharepoint_profile` tool " +
    "to load the signed-in user's profile, then greet them by name and use that context " +
    "(job title, department, office, skills, interests, responsibilities) in your answers. " +
    "Never ask the user for information already present in their profile. " +
    "If the tool returns null fields, work with what is available and do not invent values. " +
    "If the tool returns `profileAvailable: false`, explain to the user — clearly and briefly — " +
    "that no signed-in user is present in this context (for example the Foundry Playground runs " +
    "autonomously), so per-user profile data requires invoking the agent with a user token via the " +
    "proxy's POST /api/agent/chat (OBO), an M365/Teams SSO channel, or an MCP OAuth identity-passthrough " +
    "connection. Relay the `detail` field; do not invent profile values.";

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
