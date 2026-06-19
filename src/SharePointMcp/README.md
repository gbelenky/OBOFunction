# SharePointMcp — 2-mode SharePoint profile MCP server

A standalone **Model Context Protocol (MCP)** server (.NET 8, ASP.NET Core) that exposes a
single tool, `get_sharepoint_profile`, over MCP **Streamable HTTP** at `/mcp`. It is consumed
by the Foundry hosted agent — directly during local development, or via a Foundry **toolbox
OAuth connection** in the cloud (Shape 2).

## Why a separate server (not a Function)

The official `ModelContextProtocol.AspNetCore` SDK implements the MCP Streamable HTTP transport
(handshake, `Mcp-Session-Id`, SSE framing) correctly via `app.MapMcp()`. The preview Functions
MCP extension uses system-key auth on a fixed SSE path, which fights reading a forwarded user
bearer token. The OBO REST API (`src/OBOFunction`) stays a Function; the MCP server is ASP.NET Core.

## The 2-mode credential contract

`RequestCredentialProvider.Resolve()` inspects the inbound `Authorization: Bearer` header:

| Mode | Trigger | Acts as | `ResolvedVia` | SharePoint UPS |
| --- | --- | --- | --- | --- |
| **Layer A** (dev/bypass) | no bearer token | `DefaultAzureCredential` (your `az login`) | `app` | attempted (best-effort) |
| **Layer B** (passthrough) | inbound user bearer token | the **signed-in user** (`StaticTokenCredential`) | `user` | **skipped** (see limitation) |

One `if` lets the same server serve both loops with no code change.

### SharePoint UPS limitation in passthrough

The injected user token has a **fixed audience** (Graph). It cannot be exchanged for a
SharePoint-audience token, so the User Profile Service (`PeopleManager/GetMyProperties`) call is
**skipped in Layer B**. The Graph profile (`/me`) is still returned. Full SP-in-passthrough would
need either a SharePoint-scoped Foundry connection or an OBO exchange inside the MCP server.

## Local run + smoke test (Layer A)

```powershell
cd src/SharePointMcp
az login                # DefaultAzureCredential picks this up
dotnet run              # http://localhost:8089/mcp
```

MCP handshake is multi-step (Streamable HTTP requires the `Mcp-Session-Id` header on follow-ups):

```powershell
$h = @{ "Content-Type"="application/json"; "Accept"="application/json, text/event-stream" }
$init = Invoke-WebRequest http://localhost:8089/mcp -Method POST -Headers $h `
  -Body '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"smoke","version":"1.0"}}}'
$sid = ([string[]]$init.Headers['Mcp-Session-Id'])[0]
$h2 = $h + @{ "Mcp-Session-Id"=$sid }
Invoke-WebRequest http://localhost:8089/mcp -Method POST -Headers $h2 -Body '{"jsonrpc":"2.0","method":"notifications/initialized"}' | Out-Null
# list tools
Invoke-WebRequest http://localhost:8089/mcp -Method POST -Headers $h2 -Body '{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}'
# call the tool (returns YOUR profile via az login)
Invoke-WebRequest http://localhost:8089/mcp -Method POST -Headers $h2 -Body '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"get_sharepoint_profile","arguments":{}}}'
```

Health probes: `GET /liveness`, `GET /readiness`.

## Expose to Foundry cloud (Layer B)

```powershell
devtunnel host -p 8089 --allow-anonymous     # -> https://<id>.devtunnels.ms
```

Point the Foundry toolbox OAuth connection's `generic_mcp` server URL at `https://<id>.devtunnels.ms/mcp`.
The connection injects the per-user token as `Authorization: Bearer`, flipping the server into Layer B.

## Configuration (env vars)

| Variable | Purpose | Default |
| --- | --- | --- |
| `PORT` | listen port | `8089` |
| `SHAREPOINT_ROOT_SITE_URL` | SP root site for UPS | derived from hostname |
| `SHAREPOINT_TENANT_HOSTNAME` | e.g. `contoso.sharepoint.com` | — |

## Deploy

Wired into `azure.yaml` as service `sharepoint-mcp` (`host: containerapp`, multi-stage `Dockerfile`).
`azd deploy sharepoint-mcp` builds the image and pushes to the Container App.
