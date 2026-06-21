# OBOFunction — SharePoint profile for a Foundry agent, signed-in user identity end-to-end

Brings the **signed-in SharePoint user's** Microsoft Graph + SharePoint profile (including custom
User Profile Service fields — Country, Responsibilities, Past projects, Interests…) to a
**Microsoft Foundry hosted agent**, with the user's identity flowing down via the **OAuth 2.0
On-Behalf-Of (OBO)** flow. An **SPFx web part** is the auth boundary; no user token is ever stored
or logged.

The agent greets the user by first name and answers FAQ questions filtered to the user's country.

> 📐 Diagram: [`docs/architecture.excalidraw`](./docs/architecture.excalidraw) (open in <https://aka.ms/excalidraw>).
> 🧭 Design contract for AI coding agents: [`.github/copilot-instructions.md`](./.github/copilot-instructions.md).
> 🔎 Deeper identity/security rationale: [`ARCHITECTURE.md`](./ARCHITECTURE.md).

## Components

| Component | Path | Host | What it is |
|---|---|---|---|
| **Proxy** | `src/OBOFunction` | App Service (.NET 8, ASP.NET Core) | The single backend the browser talks to. Validates the SPFx user JWT, resolves the user's profile via OBO, injects it as context, and calls the Foundry hosted agent **as the user**. Endpoint: `POST /api/agent/chat`. |
| **Hosted agent** | `src/SharePointAgent` | Foundry Hosted Agent (.NET 10, Microsoft Agent Framework) | Chat agent. Receives the profile as a host-injected `USER_PROFILE_JSON` context item, so it greets by first name. Owns one **local `search_faq` tool** (Azure AI Search, filtered by the profile's country — no OBO, agent identity). |

## Architecture (one flow, "Option A")

```
SPFx web part (user JWT, aud api://<proxy-app>)
  │
  └─ POST /api/agent/chat ──► Proxy
                               │  FIRST turn only:
                               ├─ OBO user → SharePoint UPS (GetMyProperties + custom IntranetCountry)
                               ├─ OBO user → Graph /me
                               │     → builds USER_PROFILE_JSON, injected as a developer-role item
                               └─ OBO user → https://ai.azure.com/.default
                                     ▼
                               Foundry Hosted Agent (SharePointProfileAgent), called AS THE USER
                                     │  profile arrives as context; greets by first name
                                     └─ local search_faq tool ──► Azure AI Search faq-index
                                        (filter Location = profile country + Global; agent identity, no OBO)
```

**Why "Option A"?** Foundry's OAuth identity passthrough cannot forward the calling user's Entra
token to a tool through a custom *SPFx → proxy → agent* chain (hosted-agent tool discovery runs under
the agent's managed identity, so a passthrough/MCP tool is dropped and no consent is surfaced). So the
**proxy** resolves the profile itself via OBO and injects it as plain conversation context *before*
invoking the agent. The agent never needs the user's token; it only sees the profile as background
knowledge. See [`ARCHITECTURE.md`](./ARCHITECTURE.md) for the full rationale.

**Why a proxy?** SPFx is a public, third-party SharePoint client and can't be a confidential token
broker; the Foundry data plane sends no CORS headers to a `*.sharepoint.com` origin; and an agent key
in the browser would leak. The proxy is the one confidential client that holds the OBO secret and keeps
all credentials off the client. See [`spfx-sample/README.md`](./spfx-sample/README.md).

### Endpoints

| Endpoint | Auth | Purpose |
|---|---|---|
| `POST /api/agent/chat` | `Authorization: Bearer <user JWT>` (aud `api://<proxy-app>`) | Resolves the profile via OBO on the first turn, injects it as context, and calls the Foundry hosted agent as the user. Body `{ "message": "...", "previousResponseId": null, "greeting": false }` → `{ "reply", "responseId", "status" }`. |
| `GET /liveness`, `GET /readiness` | anonymous | Health probes. |

## App registration (exactly one)

Single-tenant. The client secret lives in Key Vault — never in app settings or `local.settings.json`.

### Proxy API — `api://<proxy-app>`
- **Expose an API** → scope `access_as_user` (admin + user consent).
- **Delegated permissions** (the proxy does the OBO to Graph + SharePoint itself):
  - Microsoft Graph: `User.Read`, `openid`, `profile`, `offline_access`.
  - Office 365 SharePoint Online: `AllSites.Read`, `User.Read.All` (required for the SharePoint UPS).
  - Grant **admin consent**.
- **Pre-authorized client applications** (for `access_as_user`): SharePoint Online Client Extensibility
  Web Application Principal `08e18876-6177-487e-b8b5-cf950c1e598c` (so SPFx can mint the inbound token);
  optionally Azure CLI `04b07795-8ddb-461a-bbee-02f9e1bf7b46` for local testing via `az`.
- **Client secret** → Key Vault secret `AzureAd--ClientSecret`.

> One additional admin step (outside this app reg): the proxy OBO-exchanges the user token to the
> **Azure AI / Foundry data plane** (`https://ai.azure.com`). Grant the proxy app delegated
> `user_impersonation` on the **Azure Machine Learning Services** API (admin consent) so leg ② succeeds.

### Why `User.Read.All` requires admin consent

Entra classifies the SharePoint Online **`User.Read.All`** ("Read user profiles") scope as
**admin-consent-required** because of what the scope *can* reach, independent of how this app uses it:

| Permission | Consent | Reach |
|---|---|---|
| `User.Read` | User can self-consent | Only the **signed-in user's own** profile. |
| `User.Read.All` | **Tenant admin only** | **All users'** profiles in the tenant. |

SharePoint's User Profile Service (`PeopleManager/GetMyProperties`, incl. the custom `IntranetCountry`
UPS field) is gated behind `User.Read.All` — there is no narrower "read my own UPS profile" delegated
scope. **Runtime is still per-user:** this is a delegated + OBO flow, so `GetMyProperties` returns only
*their* profile. Admin consent grants the *capability*; it is **not** app-only read-everyone access.

## Prerequisites

- Azure subscription + an Entra tenant where you can create app registrations and grant admin consent.
- A Microsoft Foundry project with a deployed chat model (`gpt-4.1-mini`).
- An Azure AI Search service + `faq-index` (FAQ entries with a `Location` field). Grant the agent's
  identity **Search Index Data Reader** on the service.
- Tools: **.NET 8 SDK** (proxy), **.NET 10 SDK** (agent), Azure Developer CLI (`azd`), PowerShell 7.

## Deploy (`azd up`)

```pwsh
cd c:\src\OBOFunction
azd auth login
azd env new obo-dev

azd env set AAD_TENANT_ID              <tenant-guid>
azd env set AAD_CLIENT_ID              <proxy-app-client-id>
azd env set --secret AAD_CLIENT_SECRET                 # proxy app secret, paste when prompted
azd env set SHAREPOINT_TENANT_HOSTNAME contoso.sharepoint.com
azd env set FOUNDRY_PROJECT_ENDPOINT   https://<acct>.services.ai.azure.com/api/projects/<proj>
azd env set AZURE_AI_MODEL_DEPLOYMENT_NAME gpt-4.1-mini
azd env set AZURE_SEARCH_ENDPOINT      https://<search-service>.search.windows.net
azd env set AZURE_SEARCH_INDEX_NAME    faq-index
azd env set AZURE_SEARCH_COUNTRY_FIELD Location

azd up
```

`azure.yaml` deploys two services: **`api`** (proxy, App Service Linux .NET 8) and
**`sharepoint-agent`** (Foundry hosted agent, .NET 10). `infra/` provisions the resource group, the
proxy web app, Storage, Key Vault, Log Analytics, App Insights and a User-Assigned Managed Identity
via [Azure Verified Modules](https://aka.ms/avm). The proxy secret lands in Key Vault; app settings
use Key Vault references.

### Key app settings (wired by `infra/`)

| Setting | Component | Purpose |
|---|---|---|
| `AzureAd__TenantId` / `AzureAd__ClientId` / `AzureAd__Audience` | proxy | Validate the inbound SPFx JWT; OBO client. |
| `AzureAd__ClientSecret` | proxy | KV reference — OBO secret. |
| `SharePoint__RootSiteUrl`, `SharePoint__TenantHostname` | proxy | UPS root host + CORS origin. |
| `Foundry__ProjectEndpoint`, `Foundry__AgentName` | proxy | Foundry project endpoint + target agent. |
| `KeyVault__Uri`, `AZURE_CLIENT_ID` | proxy | KV as a config source via the managed identity. |

## SPFx web part (consumer)

Reference docs + sample in [`spfx-sample/`](./spfx-sample/README.md). In `package-solution.json`:

```json
"webApiPermissionRequests": [
  { "resource": "api://<proxy-app-client-id>", "scope": "access_as_user" }
]
```

The web part is **agnostic**: it calls `POST /api/agent/chat` with `{ greeting: true }` on load (no
wording) and renders the agent's reply; it carries no profile or greeting logic. After uploading the
`.sppkg`, a SharePoint admin approves the request at **SharePoint Admin Center → Advanced → API access**.

## Local development & testing

```pwsh
cd src\OBOFunction
dotnet run        # POST /api/agent/chat on http://localhost:5xxx
```

- **`http/agent-chat.http`** — drive the endpoint with the
  [REST Client](https://marketplace.visualstudio.com/items?itemName=humao.rest-client) extension.
- **`tools/TokenHelper`** — `dotnet run --project tools\TokenHelper` mints an interactive user token
  (aud `api://<proxy-app>`) for `POST /api/agent/chat` without device-code flow.
- **`scripts/refresh-token.ps1`** — refreshes a cached user token for repeated local calls.

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
src/OBOFunction       .NET 8 App Service — agent-chat proxy (POST /api/agent/chat, OBO profile + OBO to agent)
src/SharePointAgent   .NET 10 Foundry hosted agent (host-injected profile; local search_faq tool)
infra                 Bicep (main.bicep + modules/resources.bicep, AVM-based)
spfx-sample           Reference SPFx consumer (docs + sample web part)
scripts               refresh-token.ps1
http                  REST Client requests (agent-chat.http)
tools/TokenHelper     Interactive user-token minter for local testing
test/OBOFunction.Tests Unit tests (Responses payload parser)
azure.yaml            azd service mapping (api, sharepoint-agent)
docs/architecture.excalidraw   The single architecture diagram
```

## Security

- Managed identity → Key Vault (Secrets User); secrets only as Key Vault references.
- Strict JWT audience + issuer validation (`Microsoft.Identity.Web`).
- CORS limited to the configured SharePoint tenant origin.
- HTTPS only, TLS 1.2+, FTPS disabled.
- The user JWT and OBO access tokens are never logged or stored.
