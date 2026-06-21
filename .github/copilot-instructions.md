# Project: OBOFunction — SharePoint profile for a Foundry agent, user identity end-to-end (OBO)

Design contract for AI coding agents. Keep changes consistent with the **single architecture** below.
For setup/run see [`../README.md`](../README.md); for identity rationale see [`../ARCHITECTURE.md`](../ARCHITECTURE.md).

## Goal
Surface the **signed-in SharePoint user's** Microsoft Graph + SharePoint (UPS, incl. custom attributes
such as `IntranetCountry`) profile to a **Microsoft Foundry hosted agent**, with the user's identity
flowing via the **OAuth 2.0 On-Behalf-Of (OBO)** flow. SPFx is the auth boundary; no user token is
stored or logged. The agent uses the profile (esp. country) to answer **country-filtered FAQ** queries.

## The two components (single architecture — "Option A")
1. **Proxy** — `src/OBOFunction`, **App Service, .NET 8 ASP.NET Core** (minimal API; *not* Functions isolated worker). The only confidential client; performs **all** OBO.
   - `POST /api/agent/chat` — validate SPFx JWT; on the FIRST turn resolve the user's profile via OBO (SharePoint UPS `GetMyProperties` + Graph `/me`), inject it as a `USER_PROFILE_JSON` developer-role context item, then OBO to `https://ai.azure.com/.default` and call the hosted **agent** (via `agent_reference`) as the user. Body `{ message?, greeting?, previousResponseId? }` → `{ reply, responseId, status }`.
   - Health: `/liveness`, `/readiness`.
2. **Hosted agent** — `src/SharePointAgent`, **Foundry hosted agent, .NET 10** (Microsoft Agent Framework). Owns one **local in-process tool** `search_faq` (Azure AI Search, agent identity). No profile tool — the profile arrives as host-injected context.

## Identity flow
- **Single path:** proxy validates the SPFx user JWT (aud `api://<proxy-app>`). On the first turn it does three OBO exchanges from the same user assertion: user→SharePoint (UPS profile), user→Graph (`/me`), user→`https://ai.azure.com/.default` (to call the agent). The profile is injected as a developer-role item; the agent is invoked **as the user** via `agent_reference` (no `model` param).
- Follow-up turns inherit context via `previousResponseId`; the profile is resolved/injected only on the first turn.
- `search_faq` runs under the **agent's own identity** (query key locally / Managed Identity + `Search Index Data Reader` in prod) — it only needs the country string, never a user token.

## Exactly one app registration
1. **Proxy API** `api://<proxy-app>`, scope `access_as_user`. Delegated Microsoft Graph (`User.Read`, `openid`, `profile`, `offline_access`) + Office 365 SharePoint Online (`AllSites.Read`, `User.Read.All`) + Azure Machine Learning Services `user_impersonation` (for the OBO to `ai.azure.com`). Pre-authorize SharePoint Online Client Extensibility `08e18876-6177-487e-b8b5-cf950c1e598c` (+ Azure CLI `04b07795-...` for local). Secret in KV (`AzureAd--ClientSecret`).

> No second app registration, no MCP server, no Foundry toolbox/OAuth-passthrough connection, no `/api/ask`, no embedded/passthrough agent profile tool. Any leftover Foundry toolbox/connections or the old `app-mcp-*` web app are orphans.

## Stack
- **Proxy:** `Microsoft.Identity.Web` + `Microsoft.Identity.Client` (OBO); `Microsoft.Graph` v5+ and SharePoint REST for the profile; `Azure.AI.Projects` to call the agent. App Insights + OpenTelemetry; Key Vault references via managed identity (no secrets in app settings). CORS allow-list limited to `https://<tenant>.sharepoint.com`.
- **Agent:** `Microsoft.Agents.AI.Foundry.Hosting` (`AddFoundryResponses`/`MapFoundryResponses`), `Azure.AI.Projects`, `Azure.Search.Documents`. No `ModelContextProtocol` dependency.

## azd / infra
- `azure.yaml` services: `api` (`host: appservice`), `sharepoint-agent` (`host: azure.ai.agent`, runtime `dotnet_10`).
- `infra/` = `main.bicep` (`targetScope='subscription'`) + `modules/resources.bicep`, composing **Azure Verified Modules** (`br/public:avm/res/...`) for one web app (proxy) on a plan, Storage, Key Vault, App Insights, Log Analytics, UAMI. No `infra/hooks`.
- Tags on every resource: `azd-env-name`, `purpose=demo`, `owner=gbelenky`, `project=obo-function`.
- `azd up` end-to-end; `azd down --purge --force` cleans up (incl. KV soft-delete); `azd pipeline config` = federated creds.

## Config keys
- **Proxy:** `AzureAd:{TenantId,ClientId,Audience,ClientSecret}`, `SharePoint:{RootSiteUrl,TenantHostname}` (UPS OBO + CORS), `Foundry:{ProjectEndpoint,AgentName,AgentResponsesUrl?,TokenScope}`, `KeyVault:Uri`.
- **Agent:** `FOUNDRY_PROJECT_ENDPOINT`, `AZURE_AI_MODEL_DEPLOYMENT_NAME`, `AZURE_SEARCH_{ENDPOINT,INDEX_NAME,COUNTRY_FIELD,API_KEY?,INCLUDE_GLOBAL?}`.

## Code quality
- Constructor DI; no service locators. RFC 7807 `ProblemDetails` on errors. No `TODO` placeholders.
- Never log the user JWT or any access token. Validate `aud` + `iss` strictly. HTTPS only, TLS 1.2+, FTPS disabled.
- Keep the endpoint-level failed-dangling-tool-call recovery (`IsFailedDanglingToolCall` in `Program.cs`) and the client-side non-2xx retry in `AgentChatClient` — both guard multi-turn continuations.
- `search_faq` MUST stay a LOCAL tool in `src/SharePointAgent` (agent identity). The proxy MUST stay tool-agnostic — never enumerate/whitelist agent tools.

## Avoid
- Functions isolated/in-process model (this is App Service ASP.NET Core now).
- Reintroducing an MCP server, Foundry toolbox, OAuth identity-passthrough connection, or a second app registration.
- Passing the user token *through* the agent; an embedded/passthrough agent profile tool; Foundry OpenAPI tool for profile data.
- ADAL (MSAL only); storing user tokens; calling Graph as the app (defeats OBO).
- Reintroducing "Shape 1/2", "Toolbox", `/api/ask`, or `get_sharepoint_profile`.

## Definition of done
- [ ] `azd up` deploys cleanly; `azd down --purge --force` cleans up fully.
- [ ] SPFx sample posts `{ greeting: true }` on load and `{ message, previousResponseId }` for chat; tenant approval documented.
- [ ] Greeting returns a one-sentence name greeting (no profile dump); explicit "what is my profile?" returns the profile; a vacation question returns the country-filtered FAQ answer.
- [ ] App Insights shows end-to-end trace: SPFx → proxy → Graph/SharePoint + Foundry agent → search_faq.
- [ ] README + ARCHITECTURE + component READMEs reflect the single (Option A) architecture and the one app registration.
