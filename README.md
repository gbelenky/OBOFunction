# OBOFunction — SharePoint Profile API + SharePoint Foundry Hosted Agent

This repo now hosts **two independent components**:

| Component | Path | What it is |
|---|---|---|
| **Profile API** | `src/OBOFunction` | .NET 8 isolated Azure Function. Reads the calling user's Microsoft Graph / SharePoint profile via the **OAuth 2.0 On-Behalf-Of (OBO) flow**, called from an **SPFx web part**. Endpoint: `GET /api/profile`. |
| **SharePoint Agent** | `src/SharePointAgent` | .NET 10 **Microsoft Foundry Hosted Agent** (Microsoft Agent Framework). A chat agent with an **embedded `get_sharepoint_profile` tool** invoked at the start of every session to load the signed-in user's profile. See [`src/SharePointAgent/README.md`](./src/SharePointAgent/README.md). |

> 📐 **See [ARCHITECTURE.md](./ARCHITECTURE.md)** for the broader architecture this Function fits into (ingestion · AI Search · request flow).
> 🧭 **See [`.github/copilot-instructions.md`](./.github/copilot-instructions.md)** for the design contract used by AI coding agents.

## Architecture

```
# Profile API (OBO)
SPFx web part (SharePoint) ──[user JWT, AadHttpClient]──► Azure Function (GET /api/profile) ──[OBO]──► Microsoft Graph + SharePoint UPS

# SharePoint Hosted Agent (replaces the old prompt-agent path)
Foundry / Agent Inspector ──[chat]──► SharePointAgent (hosted, .NET 10)
                                          └─ embedded get_sharepoint_profile tool ──► Graph /me + SharePoint UPS
```

The Function exposes two endpoints:

| Endpoint | Purpose |
|---|---|
| `GET  /api/profile` | Returns the signed-in user's profile (Graph `/me` + photo + SharePoint UPS), via OBO. |
| `POST /api/agent/chat` | **Server-side proxy** to the SharePoint hosted agent. Validates the SPFx user token, then invokes the agent with the Function's managed identity (no agent secret in the browser). Body: `{ "message": "...", "previousResponseId": null, "userProfile": null }` → `{ "reply": "...", "responseId": "..." }`. |

> The previous `POST /api/ask` endpoint (embedded **prompt**-agent call) was removed. Agent reasoning now lives in the standalone **hosted** agent under `src/SharePointAgent`. The `POST /api/agent/chat` proxy is how a browser (SPFx) safely reaches it — see [`spfx-sample/README.md`](./spfx-sample/README.md).

### Agent proxy settings (`Foundry:*`)

The proxy needs these app settings (configure for `azd` via `azd env set`, or in `local.settings.json` for local dev):

| Setting | Purpose | Default |
|---|---|---|
| `Foundry__ProjectEndpoint` | Foundry project endpoint. | — (required) |
| `Foundry__AgentName` | Deployed agent name (from `agent.yaml`). | `SharePointProfileAgent` |
| `Foundry__ApiVersion` | Responses API version. | `2025-05-01` |
| `Foundry__TokenScope` | Entra scope for the managed-identity token. | `https://ai.azure.com/.default` |
| `Foundry__ResponsesUrl` | Optional full Responses URL override (else derived as `{endpoint}/responses?api-version=...`). | — |

Grant the Function's managed identity the **Azure AI User** role on the Foundry project so it can invoke the agent. The data-plane Responses surface is still in preview — if a call fails with 404/400, adjust `Foundry__ApiVersion` / `Foundry__ResponsesUrl` to match your project's current endpoint.

#### Shape 2 (preview) — toolbox OAuth identity passthrough

By default the proxy runs **Shape 1**: it calls the agent with the Function's managed identity and injects the profile (already fetched via `GET /api/profile`) as context. To pilot **Shape 2** — where the agent's tool calls Microsoft 365 **as the signed-in user** through a Foundry **toolbox** — set `Foundry__ToolboxMcpUrl`. The proxy then OBO-exchanges the SPFx user token for an AI-audience user token and calls Responses *as the user*, attaching the toolbox as an `mcp` tool.

| Setting | Purpose | Default |
|---|---|---|
| `Foundry__ToolboxMcpUrl` | **Enables Shape 2 when set.** Toolbox consumer endpoint or MCP server URL. Blank = Shape 1. | — (blank) |
| `Foundry__ToolboxServerLabel` | `server_label` for the attached MCP tool. | `SharePointProfile` |
| `Foundry__ToolboxConnectionId` | Project connection backing the MCP server (OAuth passthrough). | `sp-profile-obo` |
| `Foundry__FeaturesHeader` | Preview feature header sent on Shape 2 calls. | `Toolboxes=V1Preview` |
| `Foundry__UserTokenScope` | OBO target scope (AI data plane). | `https://ai.azure.com/.default` |

The toolbox `SharePointProfileTools` (v1) already exists in project `prj-fdr-swc`. Its consumer endpoint:

```
https://rsc-fdr-swc.services.ai.azure.com/api/projects/prj-fdr-swc/toolboxes/SharePointProfileTools/mcp?api-version=v1
```

**One-time manual steps to make Shape 2 functional** (interactive — cannot be scripted headless):

1. **Create the `sp-profile-obo` connection** in the project, pointing at the toolbox's MCP server (a managed Microsoft 365 server for true passthrough, or this Function's own `/api/mcp` server to keep the custom SharePoint UPS fields):
   ```pwsh
   azd ai connection create --name sp-profile-obo --auth-type oauth2 `
     --project https://rsc-fdr-swc.services.ai.azure.com/api/projects/prj-fdr-swc
   ```
2. Give the signed-in user the **Foundry User** role on the project; user tenant must match the project tenant.
3. Pre-authorize the **SharePoint Online Client Extensibility** SP (`08e18876-6177-487e-b8b5-cf950c1e598c`) on the OBO app registration so SPFx can mint the inbound token.
4. The first Shape 2 call returns a **consent link** — see the SPFx consent handling below.

> All Shape 2 pieces are **preview** (Toolbox `V1Preview`, hosted-agents platform, Agent 365 managed MCP servers may be Frontier-gated). Keep `Foundry__ToolboxMcpUrl` blank to stay on GA Shape 1. Details: [ARCHITECTURE.md §8.5](./ARCHITECTURE.md).

### Does Foundry Agent Service support OBO?

**Yes — and hosted agents support it natively.** A hosted agent operates in two identity modes: **interactive** (if a user token is present, the platform runs OAuth 2.0 OBO and the agent calls downstream services as the user) and **autonomous** (no user token → the agent uses its own Entra agent identity). Agent Service also supports an "attended" OBO via the **Agent ID** framework, plus MCP-OAuth and Logic Apps connector OBO. Built-in / OpenAPI tools do **not** pass the user token through.

> **Source — Hosted agents in Foundry Agent Service (preview)**
> https://learn.microsoft.com/azure/foundry/agents/concepts/hosted-agents#agent-identity-and-endpoint

Verbatim:

> When integrated via Microsoft 365 channels (for example, Teams), Hosted agents can operate in two identity modes depending on how they are invoked:
>
> - **User-invoked scenarios (interactive)**: If a user token is present, the platform supports OAuth 2.0 On-Behalf-Of (OBO) flows. In this case, the agent can call downstream services on behalf of the user using the user's delegated permissions, subject to Microsoft Entra ID tenant policies.
> - **Autonomous or background scenarios**: If no user token is available, the agent authenticates using its own Microsoft Entra ID (agent identity), typically via managed identity, to access downstream services.

Tool-level user passthrough (recommended): *"using toolbox in Foundry for connecting tools in Hosted agent with consolidated auth support across **OAuth Identity passthrough**, agent identity, key based and more."*

> **Source — Agent identity concepts in Microsoft Foundry**
> https://learn.microsoft.com/azure/foundry/agents/concepts/agent-identity#key-concepts

Verbatim:

> Agent identities support two key authentication scenarios:
>
> 1. **Attended (delegated access or on-behalf-of flow)**: The agent operates on behalf of a human user, using the OAuth 2.0 on-behalf-of (OBO) flow. The user first authenticates to the application, and the application passes the user's token to Agent Service. Agent Service then exchanges that token for one that carries both the agent identity and the user's delegated permissions. This approach means the agent can only access resources that the user has consented to and is authorized for.
> 2. **Unattended (application-only flow)**: The agent acts under its own authority, using the OAuth 2.0 client credentials flow. Agent Service authenticates the blueprint to Microsoft Entra ID, obtains a token for the agent identity, and requests a scoped access token for the downstream resource.

> [Agents] can't initiate interactive authorization (`/authorize`) flows directly. They must receive a user token from a client application and then perform an OBO token exchange.

> **Status (evidence-based, checked 2026-06-18):** The OBO **behaviour is GA-documented as a concept** (the Agent identity concepts page is `permissioned-type: public` with no preview banner). The **hosted-agents hosting platform that delivers it is preview** — the doc title is literally *"Hosted agents in Foundry Agent Service (preview)"* and the publishing/agent-application packages are *"currently in preview."* So the capability is real; the platform maturity is preview — verify GA in your region before a hard production dependency.

**Two valid shapes for this repo:**

- **Shape 1 (GA, current default):** the **Function** does OBO → Graph/SharePoint and the agent receives the profile as input. Pure GA libraries; single security boundary; works today regardless of hosted-agent preview status.
- **Shape 2 (preview platform, forward-looking):** present the user token at agent invocation (Teams/M365 channel SSO, or the Function proxy forwarding the validated user token) so the platform performs native OBO, and switch the agent's profile tool from `DefaultAzureCredential` to **Toolbox OAuth identity passthrough** so it calls Graph/SharePoint as the user.

**SPFx** (the public *SharePoint Online Client Extensibility* principal) cannot itself be the confidential token-broker; either a backend (the **Function**) or an **M365 channel** must supply the user token. This repo ships **Shape 1** as the GA baseline and treats **Shape 2** as the path to adopt once hosted agents reach GA. Full matrix and citations: [ARCHITECTURE.md §8](./ARCHITECTURE.md#8-foundry-agent-auth-matrix--does-agent-service-support-obo).

## Prerequisites

- Azure subscription
- An Entra ID tenant where you can create app registrations and grant admin consent
- A Microsoft Foundry project with a deployed chat model (for the hosted agent)
- Tools: .NET 8 SDK (Function), **.NET 10 SDK (agent)**, Azure Functions Core Tools v4, Azure Developer CLI (`azd`), PowerShell 7

## App Registration (one-time, manual)

Create one app registration for the Function API:

1. **Entra admin center → App registrations → New registration**
   - Name: `obo-function-<env>`
   - Supported account types: single tenant
2. **Expose an API**
   - Application ID URI: `api://<client-id>` (accept default)
   - Add scope `access_as_user` (admin + user consent), display name "Access OBO Function as the signed-in user"
3. **API permissions** (delegated)
   - `Microsoft Graph` (delegated) → `User.Read`, `offline_access`, `openid`, `profile`
   - **`Office 365 SharePoint Online`** (delegated) → `AllSites.Read`, `User.Read.All`
   - Grant admin consent (required for `AllSites.Read` and `User.Read.All`)
4. **Pre-authorized client applications**
   - Add the **SharePoint Online Client Extensibility Web Application Principal** (well-known ID `08e18876-6177-487e-b8b5-cf950c1e598c` in most tenants) with the `access_as_user` scope
5. **Certificates & secrets** → New client secret. Copy the value — you'll seed it into `azd` next.

## Deploy

```pwsh
cd c:\src\OBOFunction

azd auth login
azd env new obo-dev

azd env set AAD_TENANT_ID              <your-tenant-guid>
azd env set AAD_CLIENT_ID              <function-app-registration-client-id>
azd env set --secret AAD_CLIENT_SECRET                       # paste when prompted
azd env set SHAREPOINT_TENANT_HOSTNAME contoso.sharepoint.com

# Agent proxy (POST /api/agent/chat)
azd env set FOUNDRY_PROJECT_ENDPOINT   https://<proj>.services.ai.azure.com/api/projects/<proj>
azd env set FOUNDRY_AGENT_NAME         SharePointProfileAgent

azd up
```

`azd up` provisions the resource group, Function App (Linux, .NET 8 isolated), Storage, Key Vault, Log Analytics, App Insights, and a User-Assigned Managed Identity — all via [Azure Verified Modules](https://aka.ms/avm). The client secret lands in Key Vault; the Function App settings use a Key Vault reference.

When complete, the `postdeploy` hook prints the Function URL and the JSON snippet you need for the SPFx solution.

## SPFx web part (consumer)

A starter consumer lives in `./spfx-sample`. In your `package-solution.json`:

```json
"webApiPermissionRequests": [
  {
    "resource": "api://<function-app-registration-client-id>",
    "scope": "access_as_user"
  }
]
```

Then in your web part code:

```ts
import { AadHttpClient } from "@microsoft/sp-http";

const client = await this.context.aadHttpClientFactory
  .getClient("api://<function-app-registration-client-id>");

const res = await client.get(
  "https://<function-hostname>/api/profile",
  AadHttpClient.configurations.v1
);
const profile = await res.json();
```

After uploading the `.sppkg` to the tenant App Catalog, a SharePoint admin must approve the permission request at **SharePoint Admin Center → Advanced → API access**.

### SPFx: chat + Shape 2 consent handling

Call the proxy and handle the optional Shape 2 consent ceremony. When the reply has `status === "consent_required"`, open `consentUrl`; after the user approves, resend the **same** message with `previousResponseId` set to the returned `responseId`:

```ts
async function ask(client: AadHttpClient, message: string, previousResponseId?: string) {
  const res = await client.post(
    "https://<function-hostname>/api/agent/chat",
    AadHttpClient.configurations.v1,
    { headers: { "content-type": "application/json" },
      body: JSON.stringify({ message, previousResponseId }) }
  );
  const reply = await res.json(); // { reply, responseId, status, consentUrl? }

  if (reply.status === "consent_required" && reply.consentUrl) {
    // Open the one-time consent link in a popup; after the user consents, resume the run.
    window.open(reply.consentUrl, "_blank", "width=600,height=700");
    // e.g. show a "I've signed in — continue" button that calls:
    return ask(client, message, reply.responseId);
  }
  return reply.reply as string;
}
```

`status` is `"completed"` for a normal answer (Shape 1 or a fully-consented Shape 2 turn); `"consent_required"` only occurs on the Shape 2 path before the user has granted delegated consent.

## Local development

```pwsh
cd src\OBOFunction
Copy-Item local.settings.json.sample local.settings.json
# fill in dev tenant + dev secret (or leave KeyVault__Uri set to use Key Vault via DefaultAzureCredential)
dotnet build
func start
```

Use **`http/profile.http`** with the [REST Client](https://marketplace.visualstudio.com/items?itemName=humao.rest-client) extension to invoke the endpoints. It handles Entra login + token refresh automatically — no `az` commands or token copy-pasting. First call opens a browser to sign in; subsequent calls use the cached token.

> Requires a one-time tweak: in the Function App registration's **Expose an API** blade, pre-authorize the Azure CLI public client (`04b07795-8ddb-461a-bbee-02f9e1bf7b46`) for the `access_as_user` scope.

## CI/CD

```pwsh
azd pipeline config
```

Generates `.github/workflows/azure-dev.yml` using **federated credentials** — no long-lived secrets in GitHub.

## Teardown

```pwsh
azd down --purge --force
```

Purges the Key Vault soft-delete as well.

## Project layout

```
/src/OBOFunction         .NET 8 isolated Function (ProfileFunction)
/src/SharePointAgent     .NET 10 Foundry Hosted Agent (embedded get_sharepoint_profile tool)
/infra                   Bicep (main.bicep + modules/, AVM-based)
/infra/hooks             azd lifecycle hooks (pwsh + sh)
/spfx-sample             Reference SPFx consumer (docs + sample)
azure.yaml               azd service mapping
.github/copilot-instructions.md   Project guardrails for AI agents
```

## Security

- Function MSI → Key Vault (Secrets User)
- Client secret never in app settings directly (only Key Vault reference)
- CORS limited to the configured SharePoint tenant
- Strict JWT audience + issuer validation via `Microsoft.Identity.Web`
- HTTPS only, TLS 1.2+, FTPS disabled
