# Project: OBOFunction — SharePoint Profile API for SPFx + Foundry Prompt Agent

## Goal
Build a **.NET 8 isolated-worker Azure Function** that exposes a small HTTP API for reading the calling user's **SharePoint / Microsoft Graph profile** via the **OAuth 2.0 On-Behalf-Of (OBO) flow**.

The Function is consumed by an **SPFx (SharePoint Framework) web part** running in the user's SharePoint context. The SPFx web part then invokes a **Microsoft Foundry prompt agent**, passing the retrieved profile as structured input.

This is **Shape 1** — SPFx is the auth boundary; the Foundry agent never needs the user's token.

## Architecture
```
   ┌──────────────────────────────┐
   │  SPFx web part (SharePoint)  │
   │  - AadHttpClient (user JWT)  │
   └──────────────┬───────────────┘
                  │ 1. GET /api/profile  (Authorization: Bearer <user-token>)
                  ▼
   ┌──────────────────────────────┐
   │ Azure Function (.NET 8 iso)  │
   │  - Validates user JWT        │
   │  - MSAL AcquireTokenOBO      │
   │  - Calls Microsoft Graph     │
   └──────────────┬───────────────┘
                  │ 2. UserProfile JSON
                  ▼
   ┌──────────────────────────────┐
   │  SPFx web part               │
   │  - Calls Foundry prompt      │
   │    agent with profile in     │
   │    the input payload         │
   └──────────────┬───────────────┘
                  │ 3. POST /threads/{id}/runs  (agent auth: MSI or API key)
                  ▼
   ┌──────────────────────────────┐
   │  Foundry Prompt Agent         │
   │  - Reasons over profile data  │
   │  - No Graph tool needed       │
   └──────────────────────────────┘
```

**Why this shape:** the Foundry OpenAPI tool only supports `anonymous`, `apiKey`, or the agent's own `managedIdentity` for auth — it cannot forward the calling user's Entra token. Therefore OBO must happen **before** the agent is invoked, with SPFx as the auth-bearing client.

## Required Stack

### Azure Function (this repo)
- **Runtime:** .NET 8, Azure Functions v4, **isolated worker**
- **Trigger:** HTTP, `AuthorizationLevel.Anonymous` (auth handled in code via JWT validation; Easy Auth optional as defense in depth)
- **Auth libs:** `Microsoft.Identity.Web` + `Microsoft.Identity.Client` for OBO
- **Graph SDK:** `Microsoft.Graph` v5+
- **CORS:** allow the SharePoint Online origin (`https://<tenant>.sharepoint.com`)
- **Observability:** Application Insights + OpenTelemetry
- **Config:** `IConfiguration` + Key Vault references via Managed Identity (no secrets in app settings)

### SPFx Web Part (separate component — out of scope for this repo but documented in README)
- SPFx **v1.18+**, TypeScript, React
- Uses `AadHttpClient.get(<functionUrl>, AadHttpClient.configurations.v1)`
- `package-solution.json` declares `webApiPermissionRequests` for the Function's API scope (`api://<function-app-id>/access_as_user`)
- Tenant admin approves the permission request in the SharePoint App Catalog → API access page

## OBO Flow — Implementation Requirements
1. Function receives request with `Authorization: Bearer <user-token>` from SPFx.
2. Validate the incoming token (issuer, audience = `api://<function-app-id>`, signature) using `Microsoft.Identity.Web`.
3. Use `IConfidentialClientApplication.AcquireTokenOnBehalfOf` twice with the same user assertion:
   - **Graph token** (`https://graph.microsoft.com/User.Read`) → `/me` + `/me/photo`
   - **SharePoint token** (`https://<tenant>.sharepoint.com/AllSites.Read User.Read.All`) → SharePoint REST `_api/SP.UserProfiles.PeopleManager/GetMyProperties` for the full SharePoint UPS (AboutMe, Skills, custom HR fields, etc.)
4. Merge both into a single `UserProfile` DTO.
5. Return a clean `UserProfile` DTO. **No token passthrough.**

## App Registrations Required (document in README)
Two app registrations:

**1. Function API app registration** (`api://obo-function-<env>`)
- **Expose an API** → scope `access_as_user` (admin + user consent)
- **API permissions (delegated):**
  - Microsoft Graph: `User.Read`, `offline_access`, `openid`, `profile`
  - Office 365 SharePoint Online: `AllSites.Read`, `User.Read.All` (for SharePoint UPS via `PeopleManager`)
- **Pre-authorized client applications:** the SharePoint Online Client Extensibility service principal (`08e18876-6177-487e-b8b5-cf950c1e598c` in most tenants) with scope `access_as_user`
- **Client secret** in Key Vault

**2. SPFx side**
- No dedicated app registration. SPFx uses the **SharePoint Online Client Extensibility** service principal automatically.
- `webApiPermissionRequests` in `package-solution.json`:
  ```json
  "webApiPermissionRequests": [
    { "resource": "obo-function-<env>", "scope": "access_as_user" }
  ]
  ```

## Foundry Prompt Agent (configured in Foundry portal)
- Define the agent's system prompt to expect a `userProfile` JSON object in its input
- No OpenAPI tool is required for the profile data (it arrives as input)
- SPFx calls the agent via the **Foundry data-plane REST API** or the JS SDK using either:
  - The agent's static API key (stored as an SPFx tenant property), or
  - A backend proxy (the Function itself can proxy if you want to avoid exposing keys to the browser — recommended for production)

> **Production-grade recommendation:** add a second Function endpoint `POST /api/ask` that takes a user prompt + already-validated user JWT, fetches the profile via OBO, then calls the Foundry agent server-side with the profile injected. Keeps agent credentials off the client.

## Project Layout
```
/src
  /OBOFunction
    Program.cs
    Functions/ProfileFunction.cs          # GET  /api/profile
    Functions/AskAgentFunction.cs         # POST /api/ask  (optional, recommended)
    Services/IGraphProfileService.cs
    Services/GraphProfileService.cs
    Services/IFoundryAgentClient.cs
    Services/FoundryAgentClient.cs
    Models/UserProfile.cs
    Models/AskRequest.cs / AskResponse.cs
    host.json
    local.settings.json.sample
/infra
  main.bicep
  main.parameters.json
  modules/
  abbreviations.json
/spfx-sample                              # reference SPFx web part (docs + sample code)
  README.md
azure.yaml
.azure/                                   # azd state (gitignored)
README.md
.github/copilot-instructions.md
```

## Azure Developer CLI (azd) Requirements
The project must be `azd up`-ready end-to-end.

**`azure.yaml`**
- `name: obo-function`
- `services.api`: `project: ./src/OBOFunction`, `language: dotnet`, `host: function`
- **Hooks (PowerShell + bash):**
  - `preprovision`: verify Az login, tenant; prompt for `AAD_CLIENT_ID` if missing
  - `postprovision`: write app settings, seed Key Vault secret, grant Function MSI `Key Vault Secrets User`, configure CORS for `https://<tenant>.sharepoint.com`
  - `postdeploy`: print Function base URL, the `api://...` resource ID, and the exact JSON snippet for SPFx `webApiPermissionRequests`

**`infra/` (Bicep)**
- `targetScope = 'subscription'`, creates RG then composes AVM modules
- Use **Azure Verified Modules** (`br/public:avm/res/...`) for: Function App, Storage, Key Vault, App Insights, Log Analytics, User-Assigned Managed Identity
- Required outputs: `AZURE_FUNCTION_APP_NAME`, `AZURE_FUNCTION_APP_HOSTNAME`, `AZURE_KEY_VAULT_ENDPOINT`, `AZURE_TENANT_ID`, `APPLICATIONINSIGHTS_CONNECTION_STRING`, `API_APP_RESOURCE_ID`
- Tags on every resource: `azd-env-name`, `purpose=demo`, `owner=gbelenky`, `project=obo-function`

**azd env vars (documented in README)**
| Variable | Purpose |
|---|---|
| `AZURE_ENV_NAME` / `AZURE_LOCATION` / `AZURE_SUBSCRIPTION_ID` | azd-managed |
| `AAD_TENANT_ID` | Entra tenant |
| `AAD_CLIENT_ID` | Function's App Registration |
| `AAD_CLIENT_SECRET` | `azd env set --secret AAD_CLIENT_SECRET` |
| `SHAREPOINT_TENANT_HOSTNAME` | e.g., `contoso.sharepoint.com` — used for CORS |
| `FOUNDRY_PROJECT_ENDPOINT` | Foundry project URL (for AskAgentFunction) |
| `FOUNDRY_AGENT_ID` | Target prompt agent ID |

**Local dev:** `azd env get-values > .env` → mirror into `local.settings.json`.
**CI/CD:** `azd pipeline config` → federated credentials, no long-lived secrets.
**Teardown:** `azd down --purge --force` (must clean up soft-deleted Key Vault).

## Code Quality Rules
- Constructor DI; no service locators
- Graph errors → typed exception → RFC 7807 `ProblemDetails` response
- No `TODO` placeholders
- Propagate `traceparent` from incoming request through Graph + Foundry calls
- Unit tests for `GraphProfileService` (Graph SDK request adapter mocks)
- Integration test gated by env var hitting a dev tenant

## Security Non-Negotiables
- Managed Identity for Function → Key Vault; no secrets in `local.settings.json` (use `.sample`)
- HTTPS only, TLS 1.2+, FTPS disabled
- CORS allow-list limited to the customer's SharePoint tenant origin
- Never log the user JWT or the Graph access token
- Validate `aud` and `iss` strictly on every request
- Resource tags: `purpose=demo`, `owner=gbelenky`, `project=obo-function`

## Avoid
- Foundry OpenAPI tool calling back into the Function for profile data (user identity is lost — see architecture rationale)
- Passing the user's access token through the agent (security anti-pattern; tokens leak into agent traces)
- v1 in-process Functions model
- ADAL — MSAL only
- Storing user tokens
- Calling Graph as the app (defeats OBO)

## Definition of Done
- [ ] `azd up` deploys cleanly to an empty subscription; `azd down --purge --force` cleans up fully
- [ ] `azd pipeline config` produces a working GitHub Actions workflow with federated creds
- [ ] SPFx sample web part calls `GET /api/profile` and renders the user's profile
- [ ] Tenant admin approval flow for the SPFx web API permission request is documented step-by-step
- [ ] Foundry prompt agent receives the profile as input and reasons over it (sample prompt + sample response in README)
- [ ] App Insights shows end-to-end trace: SPFx → Function → Graph → (optional) Foundry agent
- [ ] README documents prereqs, both app registration steps, `azd up`, SPFx packaging + tenant approval, Foundry agent wiring, teardown
