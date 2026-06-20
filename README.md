# OBOFunction — SharePoint profile for a Foundry agent, signed-in user identity end-to-end

Surfaces the **signed-in SharePoint user's** Microsoft Graph + SharePoint profile (including custom
User Profile Service fields — Country, Responsibilities, Past projects, Interests…) to a
**Microsoft Foundry agent**, with the user's identity flowing the whole way down via the
**OAuth 2.0 On-Behalf-Of (OBO)** flow. An **SPFx web part** is the auth boundary; no user token is
ever stored or logged. The proxy holds **no profile logic** — it delegates all retrieval to the agent.

> 📐 Diagram: [`docs/architecture.excalidraw`](./docs/architecture.excalidraw) (open in <https://aka.ms/excalidraw>).
> 🧭 Design contract for AI coding agents: [`.github/copilot-instructions.md`](./.github/copilot-instructions.md).
> 🔎 Deeper identity/security rationale: [`ARCHITECTURE.md`](./ARCHITECTURE.md).

## Components

| Component | Path | Host | What it is |
|---|---|---|---|
| **Proxy** | `src/OBOFunction` | App Service (.NET 8, ASP.NET Core) | The single backend the browser talks to. Validates the SPFx user JWT, OBO-exchanges it to the MCP audience, then **delegates all profile retrieval to the Foundry agent**. Holds no Graph/SharePoint logic. Endpoint: `POST /api/agent/chat`. |
| **MCP server** | `src/SharePointMcp` | App Service (.NET 8, ASP.NET Core) | Standalone **Model Context Protocol** server exposing `get_sharepoint_profile` over Streamable HTTP at `/mcp`. Does its own OBO → Graph + SharePoint UPS as the caller. |
| **Hosted agent** | `src/SharePointAgent` | Foundry Hosted Agent (.NET 10, Microsoft Agent Framework) | Chat agent for the **Foundry Playground / autonomous** demo. Declares the MCP server as a Foundry-native `mcp` tool (`HostedMcpServerTool`). Holds no embedded profile logic. |

## Architecture (one flow)

```
SPFx web part (user JWT, aud api://<proxy-app>)
  │
  └─ POST /api/agent/chat ──► Proxy
                               ├─ leg ① proxy app token authorizes the Foundry Responses call
                               └─ leg ② OBO user→MCP-audience token, attached to the mcp tool
                                      ▼
                               Foundry model Responses (gpt-4.1-mini)
                                      │  forwards the per-user token to the mcp tool
                                      ▼
                               SharePointMcp /mcp ──OBO──► Graph + SharePoint UPS (as the user)

Playground / autonomous: Foundry Hosted Agent ──HostedMcpServerTool──► SharePointMcp
   (no user token present → MCP server falls back to its managed identity → per-user fields empty)
```

The proxy never reads Graph or SharePoint itself — the **MCP server** is the only component that
touches the profile store, and it always does so as the signed-in user (OBO). The agent composes the
reply (e.g. the greeting + formatted profile list); the MCP server only returns raw fields.

**Why a proxy?** SPFx is a public, third-party SharePoint client and can't be a confidential
token broker; the Foundry data plane sends no CORS headers to a `*.sharepoint.com` origin; and an
agent key in the browser would leak. The proxy is the one confidential client that holds the OBO
secret and keeps all credentials off the client. See [`spfx-sample/README.md`](./spfx-sample/README.md).

### Endpoints

| Endpoint | Auth | Purpose |
|---|---|---|
| `POST /api/agent/chat` | `Authorization: Bearer <user JWT>` (aud `api://<proxy-app>`) | Server-side proxy to the Foundry model Responses endpoint with the MCP server attached as a per-user `mcp` tool. The agent retrieves and formats the profile. Body `{ "message": "...", "previousResponseId": null }` → `{ "reply", "responseId", "status", "consentUrl?" }`. |
| `GET  /liveness`, `GET /readiness` | anonymous | Health probes. |

## App registrations (exactly two)

Both are single-tenant. Secrets live in Key Vault — never in app settings or `local.settings.json`.

### 1. Proxy API — `api://<proxy-app>`
- **Expose an API** → scope `access_as_user` (admin + user consent).
- **Delegated permissions:** `openid`, `profile`, `offline_access`, and the **MCP API** scope `api://<mcp-app>/access_as_user` (so the proxy can OBO-exchange the inbound user token to the MCP audience). The proxy no longer calls Graph/SharePoint directly, so it needs **no** Graph/SharePoint delegated permissions. Grant admin consent.
- **Pre-authorized client applications** (for `access_as_user`): SharePoint Online Client Extensibility Web Application Principal `08e18876-6177-487e-b8b5-cf950c1e598c` (so SPFx can mint the inbound token); optionally Azure CLI `04b07795-8ddb-461a-bbee-02f9e1bf7b46` for local testing via `az`/`scripts/test-spfx-chain.ps1`.
- **Client secret** → Key Vault secret `AzureAd--ClientSecret`.

### 2. MCP API — `api://<mcp-app>`
- **Expose an API** → scope `access_as_user`.
- **Delegated permissions:** the same Graph + SharePoint set as above (the MCP server does its own OBO to Graph + SharePoint). Grant admin consent.
- **Pre-authorized client applications** (for `access_as_user`): the **Proxy API** app (so the proxy's OBO token is accepted); the SharePoint Online Client Extensibility principal `08e18876-...` (for the Foundry/SPFx passthrough path).
- **Client secret** → Key Vault secret `Mcp--ClientSecret`.

> Nothing else is required. Earlier exploration created a Foundry *toolbox* and extra project
> *connections* (`SharePointProfileTools`, `sp-profile-obo`, a `UserEntraToken` `SharePointMcp`
> connection) — none are used by this architecture and can be deleted from the Foundry project.

### Why `User.Read.All` requires admin consent

Entra classifies the SharePoint Online **`User.Read.All`** ("Read user profiles") scope as **admin-consent-required** because of what the scope *can* reach, independent of how this app uses it:

| Permission | Consent | Reach |
|---|---|---|
| `User.Read` | User can self-consent | Only the **signed-in user's own** profile. |
| `User.Read.All` | **Tenant admin only** | **All users'** profiles in the tenant. |

Any `*.All` directory-read scope can, by definition, read other people's data, so a regular user is not allowed to approve it — only a Global Admin / Privileged Role Admin can, and they grant it **once tenant-wide** (after which individual users never see a consent prompt).

It is still required here because SharePoint's User Profile Service (`PeopleManager/GetMyProperties`, incl. the custom `IntranetCountry` UPS field) is gated behind `User.Read.All` — there is no narrower "read my own UPS profile" delegated scope.

**Runtime is still per-user.** This is a **delegated + OBO** flow, so every call is bounded by the signed-in user's own effective permissions: `GetMyProperties` returns only *their* profile. Admin consent grants the *capability* to call the profile store; it does **not** turn this into app-only, read-everyone access (that would require the **application** permission `User.Read.All`, which this design deliberately does not use).

## Prerequisites

- Azure subscription + an Entra tenant where you can create app registrations and grant admin consent.
- A Microsoft Foundry project with a deployed chat model (`gpt-4.1-mini`).
- Tools: **.NET 8 SDK** (proxy + MCP), **.NET 10 SDK** (agent), Azure Developer CLI (`azd`), PowerShell 7.

## Deploy (`azd up`)

```pwsh
cd c:\src\OBOFunction
azd auth login
azd env new obo-dev

azd env set AAD_TENANT_ID              <tenant-guid>
azd env set AAD_CLIENT_ID              <proxy-app-client-id>
azd env set --secret AAD_CLIENT_SECRET                 # proxy app secret, paste when prompted
azd env set MCP_CLIENT_ID              <mcp-app-client-id>
azd env set --secret MCP_CLIENT_SECRET                 # mcp app secret
azd env set SHAREPOINT_TENANT_HOSTNAME contoso.sharepoint.com
azd env set FOUNDRY_PROJECT_ENDPOINT   https://<acct>.services.ai.azure.com/api/projects/<proj>

azd up
```

`azure.yaml` deploys three services: **`api`** (proxy), **`sharepoint-mcp`** (MCP server) — both
App Service Linux .NET 8 on a shared plan — and **`sharepoint-agent`** (Foundry hosted agent).
`infra/` provisions the resource group, both web apps, Storage, Key Vault, Log Analytics,
App Insights and a User-Assigned Managed Identity via [Azure Verified Modules](https://aka.ms/avm).
Both app secrets land in Key Vault; app settings use Key Vault references.

### Key app settings (wired by `infra/`)

| Setting | Component | Purpose |
|---|---|---|
| `AzureAd__TenantId` / `AzureAd__ClientId` / `AzureAd__Audience` | proxy | Validate the inbound SPFx JWT; OBO client. |
| `AzureAd__ClientSecret` | proxy | KV reference — OBO secret. |
| `SharePoint__RootSiteUrl`, `SharePoint__TenantHostname` | proxy | UPS root host + CORS origin (the proxy only uses these to build the allowed SharePoint origin). |
| `Foundry__ProjectEndpoint` | proxy | Foundry project endpoint (derives the OpenAI v1 Responses URL). |
| `Mcp__ServerUrl`, `Mcp__ServerLabel`, `Mcp__UserTokenScope` | proxy | The MCP `/mcp` endpoint + OBO target `api://<mcp-app>/.default` for leg ②. |
| `Mcp__ClientSecret` | MCP | KV reference — the MCP server's own OBO secret. |
| `KeyVault__Uri`, `AZURE_CLIENT_ID` | both | KV as a config source via the managed identity. |

## SPFx web part (consumer)

Reference docs + sample in [`spfx-sample/`](./spfx-sample/README.md). In `package-solution.json`:

```json
"webApiPermissionRequests": [
  { "resource": "api://<proxy-app-client-id>", "scope": "access_as_user" }
]
```

In the web part — call the chat endpoint (the web part auto-greets on load and renders the agent's
formatted reply):

```ts
import { AadHttpClient } from "@microsoft/sp-http";

const client = await this.context.aadHttpClientFactory.getClient("api://<proxy-app-client-id>");
const res = await client.post(
  "https://<proxy-hostname>/api/agent/chat",
  AadHttpClient.configurations.v1,
  {
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ message: "Greet me by name and list my profile.", previousResponseId: null })
  }
);
const { reply, responseId } = await res.json();
```

After uploading the `.sppkg`, a SharePoint admin approves the request at
**SharePoint Admin Center → Advanced → API access**.

## Local development & testing

```pwsh
# Proxy (POST /api/agent/chat)
cd src\OBOFunction
dotnet run

# MCP server
cd ..\SharePointMcp; az login; dotnet run     # http://localhost:8089/mcp
```

- **`http/agent-chat.http`** and **`http/mcp.http`** — drive the endpoints with the
  [REST Client](https://marketplace.visualstudio.com/items?itemName=humao.rest-client) extension.
- **`tools/TokenHelper`** — `dotnet run --project tools\TokenHelper` mints an interactive user token
  (aud `api://<proxy-app>`) for `POST /api/agent/chat` without device-code flow.
- **`scripts/test-spfx-chain.ps1`** — reproduces the **full SPFx identity chain** end-to-end without
  SharePoint UI or Foundry: user token → proxy OBO → MCP-audience token → MCP `tools/call` → OBO →
  Graph + SharePoint UPS. Run `az login --tenant <tenant>` then the script; a successful run prints
  the real profile with `"ResolvedVia": "user"`.

> For `az`/script-based local testing, pre-authorize Azure CLI `04b07795-8ddb-461a-bbee-02f9e1bf7b46`
> on the proxy app's **Expose an API** blade for `access_as_user`.

## CI/CD

```pwsh
azd pipeline config      # GitHub Actions with federated credentials — no long-lived secrets
```

## Teardown

```pwsh
azd down --purge --force      # also purges Key Vault soft-delete
```

## Project layout

```
src/OBOFunction       .NET 8 App Service — agent-only proxy (POST /api/agent/chat, OBO to MCP audience)
src/SharePointMcp     .NET 8 App Service — MCP server (get_sharepoint_profile, OBO)
src/SharePointAgent   .NET 10 Foundry hosted agent (HostedMcpServerTool → MCP server)
infra                 Bicep (main.bicep + modules/resources.bicep, AVM-based)
spfx-sample           Reference SPFx consumer (docs)
scripts               test-spfx-chain.ps1, refresh-token.ps1
http                  REST Client requests (agent-chat.http, mcp.http)
tools/TokenHelper     Interactive user-token minter for local testing
test/OBOFunction.Tests Unit tests (Responses payload parser)
azure.yaml            azd service mapping (api, sharepoint-mcp, sharepoint-agent)
docs/architecture.excalidraw   The single architecture diagram
```

## Security

- Managed identity → Key Vault (Secrets User); secrets only as Key Vault references.
- Strict JWT audience + issuer validation (`Microsoft.Identity.Web`).
- CORS limited to the configured SharePoint tenant origin.
- HTTPS only, TLS 1.2+, FTPS disabled.
- The user JWT and OBO access tokens are never logged or stored.
