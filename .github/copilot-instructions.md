# Project: OBOFunction — SharePoint profile for a Foundry agent, user identity end-to-end (OBO)

Design contract for AI coding agents. Keep changes consistent with the **single architecture** below.
For setup/run see [`../README.md`](../README.md); for identity rationale see [`../ARCHITECTURE.md`](../ARCHITECTURE.md).

## Goal
Surface the **signed-in SharePoint user's** Microsoft Graph + SharePoint (UPS, incl. custom attributes)
profile to a **Microsoft Foundry agent**, with the user's identity flowing end-to-end via the
**OAuth 2.0 On-Behalf-Of (OBO)** flow. SPFx is the auth boundary; no user token is stored or logged.

## The three components (single architecture)
1. **Proxy** — `src/OBOFunction`, **App Service, .NET 8 ASP.NET Core** (minimal API; *not* Functions isolated worker). Holds **no** Graph/SharePoint logic — delegates all profile retrieval to the agent.
   - `POST /api/agent/chat` — validate SPFx JWT → OBO to MCP audience → call the Foundry **model** Responses endpoint with the MCP server attached as a per-user `mcp` tool. Body `{ message, previousResponseId }` → `{ reply, responseId, status, consentUrl? }`.
   - Health: `/liveness`, `/readiness`.
2. **MCP server** — `src/SharePointMcp`, **App Service, .NET 8**. Tool `get_sharepoint_profile`; `RequestCredentialProvider` → `OnBehalfOfCredential` (user token present) else `DefaultAzureCredential` (app-only). Stateless HTTP transport. Endpoint `/mcp`.
3. **Hosted agent** — `src/SharePointAgent`, **Foundry hosted agent, .NET 10** (Microsoft Agent Framework). Declares the MCP server via `HostedMcpServerTool` (URL-bound). Playground/autonomous demo only. No embedded profile tool.

## Identity flow
- **Chat (single path):** proxy validates the SPFx user JWT (aud `api://<proxy-app>`); leg ① app token (`DefaultAzureCredential`, `https://ai.azure.com/.default`) authorizes the Responses call; leg ② OBO user→MCP-audience token in the `mcp` tool's `authorization` → Foundry forwards it to the MCP server → MCP server OBO → Graph + SharePoint as the user. The proxy never reads Graph/SharePoint itself.
- The proxy targets the **raw model** (gpt-4.1-mini), not the named hosted agent (hosted agents reject request-supplied tools).

## Exactly two app registrations
1. **Proxy API** `api://<proxy-app>`, scope `access_as_user`. Delegated `openid`/`profile`/`offline_access` + the **MCP API** scope `api://<mcp-app>/access_as_user` (the proxy OBOs only to the MCP audience — **no** Graph/SharePoint delegated perms). Pre-authorize SharePoint Online Client Extensibility `08e18876-6177-487e-b8b5-cf950c1e598c` (+ Azure CLI `04b07795-...` for local). Secret in KV (`AzureAd--ClientSecret`).
2. **MCP API** `api://<mcp-app>`, scope `access_as_user`. Same delegated Graph + SharePoint perms. Pre-authorize the **Proxy API** app + the SharePoint Online Client Extensibility principal. Secret in KV (`Mcp--ClientSecret`).

> No third app registration, no Foundry toolbox, no `/api/ask`, no embedded agent tool. Any leftover Foundry toolbox/connections are orphans.

## Stack
- `Microsoft.Identity.Web` + `Microsoft.Identity.Client` (OBO). `Microsoft.Graph` v5+ is used by the **MCP server** (not the proxy). `ModelContextProtocol.AspNetCore` (pin `ModelContextProtocol` **1.2.0**), `Microsoft.Agents.AI.Foundry.Hosting`.
- App Insights + OpenTelemetry; Key Vault references via managed identity (no secrets in app settings).
- CORS allow-list limited to `https://<tenant>.sharepoint.com`.

## azd / infra
- `azure.yaml` services: `api` (`host: appservice`), `sharepoint-mcp` (`host: appservice`), `sharepoint-agent` (`host: azure.ai.agent`, runtime `dotnet_10`).
- `infra/` = `main.bicep` (`targetScope='subscription'`) + `modules/resources.bicep`, composing **Azure Verified Modules** (`br/public:avm/res/...`) for two web apps on a shared plan, Storage, Key Vault, App Insights, Log Analytics, UAMI. No `infra/hooks`.
- Tags on every resource: `azd-env-name`, `purpose=demo`, `owner=gbelenky`, `project=obo-function`.
- `azd up` end-to-end; `azd down --purge --force` cleans up (incl. KV soft-delete); `azd pipeline config` = federated creds.

## Config keys
- **Proxy:** `AzureAd:{TenantId,ClientId,Audience,ClientSecret}`, `SharePoint:{RootSiteUrl,TenantHostname}` (CORS only), `Foundry:{ProjectEndpoint,TokenScope}`, `Mcp:{ServerUrl,ServerLabel,UserTokenScope}`, `KeyVault:Uri`.
- **MCP server:** `AzureAd:{...,ClientSecret}`, `Graph:Scopes`, `SharePoint:{RootSiteUrl,Scopes,TenantHostname}`, `KeyVault:Uri`.

## Code quality
- Constructor DI; no service locators. RFC 7807 `ProblemDetails` on errors. No `TODO` placeholders.
- Never log the user JWT or any access token. Validate `aud` + `iss` strictly. HTTPS only, TLS 1.2+, FTPS disabled.
- Keep the `AgentChatReply.Status`/`ConsentUrl` consent contract (MCP OAuth passthrough can emit consent links).

## Avoid
- Functions isolated/in-process model (this is App Service ASP.NET Core now).
- Foundry OpenAPI/embedded tool for profile data; passing the user token *through* the agent.
- ADAL (MSAL only); storing user tokens; calling Graph as the app (defeats OBO).
- Reintroducing "Shape 1/2", "Toolbox", `/api/ask`, or an embedded agent profile tool.

## Definition of done
- [ ] `azd up` deploys cleanly; `azd down --purge --force` cleans up fully.
- [ ] `scripts/test-spfx-chain.ps1` returns the full profile with `ResolvedVia: "user"`.
- [ ] SPFx sample calls `POST /api/agent/chat` (auto-greets on load); tenant approval documented.
- [ ] App Insights shows end-to-end trace: SPFx → proxy → Graph/SharePoint → Foundry → MCP.
- [ ] README + ARCHITECTURE + component READMEs reflect the single architecture and the two app registrations.
