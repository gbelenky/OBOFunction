# OBOFunction

This repository contains the **BFF proxy** for a SharePoint-hosted SPFx web part that talks to a
**Microsoft Foundry hosted agent** as the signed-in user.

## Current architecture

```text
SPFx web part (SharePoint page)
  -> POST /api/agent/chat
  -> OBOFunction proxy (App Service, .NET 8)
  -> OBO to https://ai.azure.com/.default
  -> HostedSecureMcpAgent (existing Foundry hosted agent, outside this repo)
  -> agent-owned Toolbox / MCP / downstream tools
```

## What is in this repo

| Component | Path | Purpose |
| --- | --- | --- |
| BFF proxy | `src/OBOFunction` | Validates the SPFx JWT, extracts the user bearer token, performs OBO to Foundry, and forwards chat turns to the configured hosted agent. |
| SPFx sample | `spfx-sample/spfx-profile-agent` | Reference client that calls the proxy with `AadHttpClient`. |
| Infra | `infra/` | Bicep for the proxy App Service, App Insights, Log Analytics, plan, and user-assigned managed identity. |
| Tests | `test/OBOFunction.Tests` | Unit tests for the proxy’s Foundry response parsing and error handling. |

## What is **not** in this repo anymore

- No local hosted-agent project
- No local `search_faq` tool implementation
- No local profile-fetching service
- No Key Vault dependency in the runtime path
- No Azure Functions host / Functions-specific runtime model

The **actual hosted agent is `HostedSecureMcpAgent` in Foundry** and is configured through proxy app
settings. The hosted-agent source is published at
[`gbelenky/HostedSecureMcpAgent`](https://github.com/gbelenky/HostedSecureMcpAgent):

- `Foundry__AgentName=HostedSecureMcpAgent`
- `Foundry__AgentResponsesUrl=https://rsc-fdr-swc.services.ai.azure.com/api/projects/prj-fdr-swc/agents/HostedSecureMcpAgent/endpoint/protocols/openai/responses?api-version=v1`

## Why the BFF proxy is required

The proxy is not incidental; it is the design boundary that makes this solution viable.

### 1. SPFx is a public client, but OBO needs a confidential client

The browser can obtain a user token for `api://<proxy-app>`, but it **cannot safely hold the
client secret** required to perform the OBO exchange to the Foundry data plane. The proxy is the
one confidential component allowed to hold that secret.

### 2. The browser must not call Foundry directly

The Foundry Responses endpoint is a server-side API. In this solution, the browser talks only to
SharePoint and the proxy. Keeping Foundry behind the proxy:

- keeps the Foundry endpoint, auth scope, and agent routing out of the client bundle;
- avoids exposing server-side credentials or bearer-token minting logic in the browser; and
- gives the app one stable API contract: `POST /api/agent/chat`.

### 3. The proxy is the policy and trust boundary

The proxy enforces:

- JWT validation (`aud`, issuer, signature)
- SharePoint-origin CORS
- RFC 7807 error shaping
- conversation continuity handling (`previousResponseId`)
- server-side telemetry and trace correlation

Without the proxy, those controls would have to move into an environment that cannot safely own
them: the browser.

### 4. The live system proves the proxy’s role

The live SharePoint page was validated end-to-end through this path:

- SharePoint page -> SPFx web part
- SPFx web part -> `https://app-proxy-z6vb2tjg2j4ye.azurewebsites.net/api/agent/chat`
- Proxy -> `HostedSecureMcpAgent`

That browser verification returned:

- a greeting for the signed-in user
- the signed-in user’s profile from the hosted agent path

So the proxy is not theoretical; it is the **working production integration point**.

For deeper reasoning, see [`ARCHITECTURE.md`](./ARCHITECTURE.md).

## Runtime behavior

### `POST /api/agent/chat`

Request:

```json
{ "message": "What is my profile?" }
```

or for first-load greeting:

```json
{ "message": "", "greeting": true }
```

Follow-up turn:

```json
{ "message": "And what about vacation policy?", "previousResponseId": "caresp_..." }
```

Response:

```json
{ "reply": "...", "responseId": "caresp_...", "status": "success" }
```

### Health

- `GET /liveness`
- `GET /readiness`

## Configuration

### Proxy app settings

| Setting | Purpose |
| --- | --- |
| `AzureAd__TenantId` | Entra tenant for inbound JWT validation and OBO |
| `AzureAd__ClientId` | Proxy app registration client id |
| `AzureAd__Audience` | Expected audience (`api://<client-id>`) |
| `AzureAd__ClientSecret` | Confidential secret used for OBO |
| `SharePoint__TenantHostname` | Allowed SharePoint origin for CORS |
| `Foundry__ProjectEndpoint` | Foundry project base endpoint |
| `Foundry__AgentName` | Hosted agent name (`HostedSecureMcpAgent`) |
| `Foundry__AgentResponsesUrl` | Explicit Responses endpoint override |
| `Foundry__ApiVersion` | Foundry Responses API version (`v1`) |
| `Foundry__TokenScope` | OBO target scope (`https://ai.azure.com/.default`) |

### App registration

One app registration is used for the proxy API:

- exposes `access_as_user`
- is requested by SPFx through `AadHttpClient`
- is the confidential client that performs OBO to Foundry

## Deploy

This repo now deploys **only the proxy**.

```powershell
cd C:\src\OBOFunction
azd up
```

Useful follow-up commands:

```powershell
azd deploy api
azd env get-values
```

## Local development

```powershell
cd C:\src\OBOFunction\src\OBOFunction
dotnet run
```

Default local URL from `launchSettings.json`:

- `http://localhost:7071`

REST client sample:

- `http/agent-chat.http`

## Validation

### Build

```powershell
dotnet build .\src\OBOFunction\OBOFunction.csproj -c Debug
```

### Test

```powershell
dotnet test .\test\OBOFunction.Tests\OBOFunction.Tests.csproj -c Debug
```

### Live browser verification

The SharePoint page was verified with local Edge + Playwright against:

- `https://mngenv168112.sharepoint.com/sites/gbelelenky-dev-sp/SitePages/CollabHome.aspx`

Artifacts from the verification live under the session-state browser helper, not the repo.

## Project layout

```text
src/OBOFunction                  ASP.NET Core BFF proxy
spfx-sample/spfx-profile-agent   SPFx sample client
infra/                           Bicep for the proxy and observability resources
http/                            REST Client examples
test/OBOFunction.Tests           Proxy unit tests
```
