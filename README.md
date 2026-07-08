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
| **Proxy** | `src/OBOFunction` | App Service (.NET 8, ASP.NET Core) | The single backend the browser talks to. Validates the SPFx user JWT, extracts the user's Bearer token, and calls the Foundry hosted agent **as the user** via OBO identity pass-through. Endpoint: `POST /api/agent/chat`. |
| **Hosted agent** | `src/SharePointAgent` | Foundry Hosted Agent (.NET 10, Microsoft Agent Framework) | Chat agent. Receives the user's OBO token (identity pass-through), so it executes as the signed-in user. On first use, the agent's **Toolbox connection** fetches the user's profile (name, email, country). Owns one **local `search_faq` tool** (Azure AI Search, filtered by the profile's country — agent identity, no OBO). |

## Architecture (one flow, "Option A")

```
SPFx web part (user JWT, aud api://<proxy-app>)
  │
  └─ POST /api/agent/chat ──► Proxy
                               │
                               ├─ Validates SPFx JWT (aud, iss, sig)
                               ├─ Extracts user's Bearer token
                               ├─ Creates OBO credential (proxy app + user token)
                               └─ OBO user → https://ai.azure.com/.default
                                     ▼
                               Foundry Hosted Agent (SharePointProfileAgent)
                               Called AS THE USER (OBO token)
                                     │
                                     ├─ FIRST turn: Toolbox connection fetches user's profile
                                     │  └─ (May show sign-in prompt for Toolbox consent)
                                     │
                                     └─ Calls search_faq tool ──► Azure AI Search faq-index
                                        (filter Location = profile country + Global; agent identity)
```

**Why OBO pass-through?** The proxy no longer resolves the profile itself. Instead, it passes the
user's identity to the agent via the OBO flow. The agent's **Toolbox connection** (Foundry-managed)
fetches the user's profile on demand using standard OAuth flows. This design:
- Simplifies the proxy (pure identity pass-through, no profile extraction)
- Follows Foundry best practices (Toolbox for data, not embedded profile tools)
- Supports user consent flows (Toolbox connection may prompt on first use)
- Keeps `search_faq` under the agent's identity (least-privilege for FAQ access)

See [`ARCHITECTURE.md`](./ARCHITECTURE.md) for full rationale and [`REFACTOR_SUMMARY.md`](./REFACTOR_SUMMARY.md) for before/after comparison.

**Why a proxy?** SPFx is a public, third-party SharePoint client and can't be a confidential token
broker; the Foundry data plane sends no CORS headers to a `*.sharepoint.com` origin; and an agent key
in the browser would leak. The proxy is the one confidential client that holds the OBO secret and keeps
all credentials off the client. See [`spfx-sample/README.md`](./spfx-sample/README.md).

### Endpoints

| Endpoint | Auth | Purpose |
|---|---|---|
| `POST /api/agent/chat` | `Authorization: Bearer <user JWT>` (aud `api://<proxy-app>`) | Validates the user's JWT, extracts the Bearer token, creates an OBO credential, and calls the Foundry hosted agent as the user. The agent then uses its Toolbox connection to fetch the user's profile on demand. Body `{ "message": "...", "previousResponseId": null, "greeting": false }` → `{ "reply", "responseId", "status" }`. |
| `GET /liveness`, `GET /readiness` | anonymous | Health probes. |

## App registration (exactly one)

Single-tenant. The client secret lives in Key Vault — never in app settings or `local.settings.json`.

### Proxy API — `api://<proxy-app>`
- **Expose an API** → scope `access_as_user` (admin + user consent).
- **Delegated permissions** (proxy needs only to mint the inbound token and OBO to Foundry):
  - Microsoft Graph: `User.Read`, `openid`, `profile`, `offline_access` (for user consent + token cache).
  - Office 365 SharePoint Online: `AllSites.Read` (CORS/SPFx context only; profile now fetched by agent's Toolbox).
  - Azure Machine Learning Services: `user_impersonation` (for OBO to Foundry's `https://ai.azure.com`).
  - Grant **admin consent**.
- **Pre-authorized client applications** (for `access_as_user`): SharePoint Online Client Extensibility
  Web Application Principal `08e18876-6177-487e-b8b5-cf950c1e598c` (so SPFx can mint the inbound token);
  optionally Azure CLI `04b07795-8ddb-461a-bbee-02f9e1bf7b46` for local testing via `az`.
- **Client secret** → Key Vault secret `AzureAd--ClientSecret`.

> **Note on permissions:** The proxy no longer needs `User.Read.All` because it no longer fetches the
> user's profile. The agent's **Toolbox connection** is a separate OAuth app that handles all profile
> fetching on the agent's behalf. See [`REFACTOR_SUMMARY.md`](./REFACTOR_SUMMARY.md) for the refactor rationale.

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

## Observability

Tracing uses **OpenTelemetry → Azure Monitor (Application Insights)** with the
[GenAI semantic conventions](https://opentelemetry.io/docs/specs/semconv/gen-ai/) — the Foundry-native
pattern. Details and KQL in [`ARCHITECTURE.md` §6](./ARCHITECTURE.md#6-observability-opentelemetry--application-insights-genai-semantic-conventions).

- **Proxy** emits one GenAI span per `POST /api/agent/chat` turn (`OBOFunction.AgentChat` source), tagged
  `gen_ai.response.id` / `gen_ai.previous_response.id` (the conversation cursor), `gen_ai.response.status`,
  `gen_ai.agent.name`, and `enduser.id` (the caller's Entra `oid` — **never** name/email/token). It
  propagates W3C `traceparent` to the agent.
- **Hosted agent** gets server-side GenAI traces (model + `search_faq`) **with no code change**, once an
  App Insights resource is connected to the Foundry project:
  **portal → project → Observability / Tracing → Connect** (already connected on `prj-fdr-swc` as the
  `AppInsights` project connection).
- Both halves share a trace id, so the full **SPFx → proxy → Graph/SharePoint + agent → `search_faq`**
  path is one correlated trace.
