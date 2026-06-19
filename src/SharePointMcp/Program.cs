using Azure.Identity;
using SharePointMcp.Services;
using SharePointMcp.Tools;

// ----------------------------------------------------------------------------
// SharePointMcp — standalone Model Context Protocol (MCP) server (.NET 8, ASP.NET Core)
//
// Exposes `get_sharepoint_profile` over MCP Streamable HTTP at /mcp using the official
// ModelContextProtocol.AspNetCore SDK. Two-mode auth (see RequestCredentialProvider):
//   - User mode (prod): inbound `Authorization: Bearer <user token>` -> OBO -> acts AS THE USER
//     (the proxy attaches this token to the Foundry mcp tool; Foundry forwards it here).
//   - App mode (dev):   no bearer -> DefaultAzureCredential -> acts as the developer/managed identity.
//
// Run locally:  dotnet run  ->  http://localhost:8089/mcp
// Transport is Stateless (required: Foundry's MCP client does not retain Mcp-Session-Id).
// ----------------------------------------------------------------------------

var builder = WebApplication.CreateBuilder(args);

// Key Vault as a configuration source (Managed Identity). Secrets surface as config keys with
// '--' mapped to ':' — e.g. Mcp--ClientSecret -> Mcp:ClientSecret (the OBO app secret).
var kvUri = builder.Configuration["KeyVault:Uri"];
if (!string.IsNullOrWhiteSpace(kvUri))
{
    builder.Configuration.AddAzureKeyVault(new Uri(kvUri), new DefaultAzureCredential());
}

string port =
    Environment.GetEnvironmentVariable("PORT")
    ?? "8089";
builder.WebHost.UseUrls($"http://+:{port}");

builder.Services.AddHttpContextAccessor();
builder.Services.AddSingleton(_ => new DefaultAzureCredential());
builder.Services.AddScoped<RequestCredentialProvider>();
builder.Services.AddScoped<ProfileMcpTools>();

builder.Services
    .AddMcpServer()
    .WithHttpTransport(o => o.Stateless = true)
    .WithTools<ProfileMcpTools>();

var app = builder.Build();

app.MapMcp("/mcp");
app.MapGet("/liveness", () => Results.Ok("Healthy"));
app.MapGet("/readiness", () => Results.Ok("Ready"));

app.Run();
