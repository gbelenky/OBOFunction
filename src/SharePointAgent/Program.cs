using Azure.AI.Projects;
using Azure.Identity;
using Microsoft.Agents.AI;
using Microsoft.Agents.AI.Foundry.Hosting;
using Microsoft.Extensions.AI;
using ModelContextProtocol.Client;
using SharePointAgent.Services;
using SharePointAgent.Tools;

// ----------------------------------------------------------------------------
// SharePointAgent — Microsoft Foundry Hosted Agent (.NET 10)
//
// Exposes a single chat agent whose embedded `get_sharepoint_profile` tool fetches
// the signed-in user's Graph + SharePoint profile at the start of every session.
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
    Environment.GetEnvironmentVariable("AZURE_AI_MODEL_DEPLOYMENT_NAME")
    ?? Environment.GetEnvironmentVariable("FOUNDRY_MODEL_DEPLOYMENT_NAME")
    ?? "gpt-4.1-mini";

string? sharePointRootSiteUrl = Environment.GetEnvironmentVariable("SHAREPOINT_ROOT_SITE_URL");
if (string.IsNullOrWhiteSpace(sharePointRootSiteUrl))
{
    var host = Environment.GetEnvironmentVariable("SHAREPOINT_TENANT_HOSTNAME");
    if (!string.IsNullOrWhiteSpace(host))
        sharePointRootSiteUrl = $"https://{host}";
}

var credential = new DefaultAzureCredential();

// Tool sourcing is flag-gated so the agent can be tested two ways:
//   * MCP_SERVER_URL set   -> consume tools from the standalone SharePointMcp server
//                             over MCP (Layer A local test, or a deployed/tunnelled server).
//   * MCP_SERVER_URL unset -> use the embedded in-process get_sharepoint_profile tool
//                             (original self-contained behaviour).
IList<AITool> tools;
string? mcpServerUrl = Environment.GetEnvironmentVariable("MCP_SERVER_URL");

if (!string.IsNullOrWhiteSpace(mcpServerUrl))
{
    var transport = new HttpClientTransport(new HttpClientTransportOptions
    {
        Endpoint = new Uri(mcpServerUrl),
        Name = "SharePointMcp",
        TransportMode = HttpTransportMode.AutoDetect,
    });

    // Kept alive for the process lifetime; tool invocations flow through this client.
    McpClient mcpClient = await McpClient.CreateAsync(transport);
    IList<McpClientTool> mcpTools = await mcpClient.ListToolsAsync();
    tools = [.. mcpTools];
}
else
{
    var profileService = new SharePointProfileService(credential, sharePointRootSiteUrl);
    var profileTools = new ProfileTools(profileService);
    tools = [profileTools.CreateTool()];
}

const string instructions =
    "You are the SharePoint Profile Assistant. " +
    "At the START of every new conversation, FIRST call the `get_sharepoint_profile` tool " +
    "to load the signed-in user's profile, then greet them by name and use that context " +
    "(job title, department, office, skills, interests, responsibilities) in your answers. " +
    "Never ask the user for information already present in their profile. " +
    "If the tool returns null fields, work with what is available and do not invent values.";

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
app.MapGet("/liveness", () => Results.Ok("Healthy"));
app.MapGet("/readiness", () => Results.Ok("Ready"));

app.Run();
