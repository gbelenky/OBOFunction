# SharePointMcp — SharePoint profile MCP server

A standalone **Model Context Protocol (MCP)** server (.NET 8, ASP.NET Core) exposing one tool,
`get_sharepoint_profile`, over MCP **Streamable HTTP** at `/mcp`. It returns the caller's Microsoft
Graph `/me` profile **and** selected SharePoint User Profile Service (UPS) fields — including custom
attributes (Country, Responsibilities, Past projects, Interests).

Consumed two ways, **same code, same identity model**:
- the **proxy** (`src/OBOFunction`, `POST /api/agent/chat`) attaches a per-user token to the Foundry
  `mcp` tool, which Foundry forwards here;
- the **Foundry hosted agent** (`src/SharePointAgent`) references it via `HostedMcpServerTool` for the
  Playground/autonomous demo (needs a Foundry OAuth identity-passthrough connection to supply the user token).

## Identity model (one `if`)

`RequestCredentialProvider.Resolve()` inspects the inbound `Authorization: Bearer` header:

| Inbound | Acts as | `ResolvedVia` | SharePoint UPS |
|---|---|---|---|
| **User token** (audience = this server, `api://<mcp-app>`) | the signed-in user via `OnBehalfOfCredential` (OBO) | `user` | ✅ returned (OBO re-exchanges for both Graph and SharePoint scopes) |
| **No token** (local inner loop) | `DefaultAzureCredential` (your `az login` / managed identity) | `app` | best-effort; app-only/CLI tokens usually lack SharePoint delegated consent → empty |

`OnBehalfOfCredential` runs a fresh OBO exchange **per requested scope**, so a single inbound user
token reaches **both** Graph (`/me`) and SharePoint (`PeopleManager/GetMyProperties`) — which is why
custom UPS fields surface in user mode. Proven by `scripts/test-spfx-chain.ps1`.

The HTTP transport is **Stateless** (`WithHttpTransport(o => o.Stateless = true)`) — required, because
Foundry's MCP client does not retain `Mcp-Session-Id`.

## Why a separate server (not the proxy, not a Function)

The official `ModelContextProtocol.AspNetCore` SDK implements the MCP Streamable HTTP transport
correctly via `app.MapMcp("/mcp")`. Keeping it separate from the OBO REST proxy lets Foundry attach it
as a first-class `mcp` tool and keeps its app registration (audience) distinct, which is what enables a
clean two-hop OBO chain (proxy audience → MCP audience → Graph/SharePoint).

## Configuration

| Key | Purpose |
|---|---|
| `AzureAd:TenantId`, `AzureAd:ClientId` | the MCP server's own app registration (`api://<mcp-app>`) — the OBO confidential client. |
| `Mcp:ClientSecret` | OBO client secret (Key Vault reference `Mcp--ClientSecret`). |
| `KeyVault:Uri`, `AZURE_CLIENT_ID` | Key Vault as a config source via managed identity. |
| `SharePoint:RootSiteUrl`, `SharePoint:TenantHostname`, `SharePoint:Scopes` | UPS root + SharePoint OBO scope. |
| `PORT` | listen port (default `8089`). |

The OBO config is required only for the user-token path; the local-dev path never reads it.

## Local run + smoke test

```powershell
cd src/SharePointMcp
az login          # DefaultAzureCredential picks this up
dotnet run        # http://localhost:8089/mcp
```

Use `../../http/mcp.http` (REST Client) or the stateless one-shot below:

```powershell
$h = @{ "Content-Type"="application/json"; "Accept"="application/json, text/event-stream" }
Invoke-WebRequest http://localhost:8089/mcp -Method POST -Headers $h `
  -Body '{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{}}'
Invoke-WebRequest http://localhost:8089/mcp -Method POST -Headers $h `
  -Body '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"get_sharepoint_profile","arguments":{}}}'
```

Health probes: `GET /liveness`, `GET /readiness`.

To exercise the **user (OBO)** path end-to-end without SharePoint UI, run `scripts/test-spfx-chain.ps1`
(it mints a proxy-audience user token, OBO-exchanges to this server's audience, then calls `tools/call`).

## Deploy

Wired into `azure.yaml` as service **`sharepoint-mcp`** (`host: appservice`). `azd deploy sharepoint-mcp`
publishes the .NET app to its App Service. `infra/` provisions the web app, its managed identity, and the
Key Vault reference for `Mcp--ClientSecret`.
