# SPFx sample — consuming the profile API and chatting with the agent

An SPFx web part calls **one backend** (the proxy, `src/OBOFunction`) for two things, both with the
signed-in user's identity via `AadHttpClient`:

1. **`GET /api/profile`** — the user's Microsoft Graph + SharePoint profile (OBO).
2. **`POST /api/agent/chat`** — chat. The proxy calls the Foundry model server-side and attaches the
   **SharePointMcp** server as a per-user `mcp` tool, so the agent reasons over the *real signed-in
   user's* profile.

The browser never talks to Foundry or the MCP server directly.

## Built solution in this repo: `spfx-profile-agent/`

A ready-to-build SPFx **1.23** React web part is scaffolded under
[`spfx-profile-agent/`](./spfx-profile-agent). It is already wired to the live proxy:

| Wiring | Value | Where |
|---|---|---|
| `PROXY_RESOURCE` | `api://7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a` | `src/webparts/profileAgent/services/ProxyClient.ts` |
| `PROXY_BASE` | `https://app-proxy-z6vb2tjg2j4ye.azurewebsites.net` | same |
| `webApiPermissionRequests` | `api://7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a` / `access_as_user` | `config/package-solution.json` |

The web part loads the profile on mount (`GET /api/profile`) and renders a chat box wired to
`POST /api/agent/chat`, with multi-turn `previousResponseId` and OAuth-consent handling.

### Build (toolchain)

SPFx 1.23 tooling requires **Node 22** (not Node 24). Use nvm-windows to keep an isolated Node 22:

```powershell
nvm install 22
nvm use 22
cd spfx-sample\spfx-profile-agent
npm install
npm run build          # heft: lint + test + package -> sharepoint/solution/spfx-profile-agent.sppkg
```

`npm run build` is green with **0 warnings / 0 errors** and emits
`sharepoint/solution/spfx-profile-agent.sppkg` (gitignored — it is a build artifact).

### Local workbench

```powershell
npm run start          # heft start -> https://<tenant>.sharepoint.com/_layouts/workbench.aspx
```

### Deploy + tenant approval

1. Upload `sharepoint/solution/spfx-profile-agent.sppkg` to the tenant **App Catalog**, choose
   *Make this solution available to all sites*.
2. Approve the API permission request in **SharePoint Admin Center → Advanced → API access**:
   `access_as_user` on `api://7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a` (the SharePoint Online Client
   Extensibility principal is the caller).
3. Add the **ProfileAgent** web part to a page.

The blueprint below documents the contract and code that this solution implements.


## Required SPFx version

SPFx **1.18+**, TypeScript, React (or no-framework — the code is identical).

## Why the browser can't call the agent directly

Foundry's Responses data-plane API needs an **Entra token for the AI resource**
(`https://ai.azure.com/.default`). From a `*.sharepoint.com` page that's blocked three ways:

- the Foundry endpoint sends **no CORS headers** to a SharePoint origin → `fetch` fails;
- an **agent/API key** in the client bundle would leak;
- the user's identity must reach Graph/SharePoint via **OBO**, which only a confidential server can do.

So the proxy fronts the call. It is the one confidential client holding the OBO secret and the AI app
identity.

```
SPFx ──[AadHttpClient, user JWT]──► proxy POST /api/agent/chat
        leg ① proxy app token authorizes the Foundry call
        leg ② OBO user→MCP-audience token on the mcp tool
             ▼
        Foundry model Responses ──forwards per-user token──► SharePointMcp ──OBO──► Graph + SharePoint
```

## Step 1 — the proxy endpoint (already implemented)

`POST /api/agent/chat` (in `src/OBOFunction/Program.cs`, logic in `Services/AgentChatClient.cs`)
validates the SPFx user token (audience `api://<proxy-app>`, same as `/api/profile`), then:

- **leg ①** authorizes the Foundry **model** Responses call with the proxy's own identity (`DefaultAzureCredential` — managed identity in Azure);
- **leg ②** OBO-exchanges the user token for the **MCP server's audience** and places it in the `mcp` tool's `authorization` field; Foundry forwards it to the MCP server, which does its own OBO to Graph + SharePoint.

> It targets the **raw model deployment** (e.g. `gpt-4.1-mini`), not the named hosted agent, because
> hosted agents reject request-supplied tools and the per-user `authorization` must be attached per call.

Request / response contract:

```jsonc
// POST /api/agent/chat
{ "message": "Who am I and what are my skills?", "previousResponseId": null }

// 200 OK
{ "reply": "...", "responseId": "resp_...", "status": "completed", "consentUrl": null }
```

Pass `responseId` back as `previousResponseId` on the next turn (server-side multi-turn state). If a
Foundry OAuth identity-passthrough tool needs consent, the reply is `{ "status": "consent_required",
"consentUrl": "https://..." }` — open the link, then re-send the same message with `previousResponseId`
set to that `responseId`.

App settings (and grant the proxy's managed identity the **Azure AI User** role on the Foundry project):

| Setting | Value |
|---|---|
| `Foundry__ProjectEndpoint` | `https://<acct>.services.ai.azure.com/api/projects/<project>` |
| `Mcp__ServerUrl` | `https://app-mcp-<token>.azurewebsites.net/mcp` |
| `Mcp__ServerLabel` | `SharePointMcp` |
| `Mcp__UserTokenScope` | `api://<mcp-app>/.default` (OBO target audience) |
| `AzureAd__*` | proxy app registration + secret (the OBO confidential client) |

## Step 2 — `package-solution.json`

The web part only needs permission to the **proxy API** (it fronts both endpoints):

```json
{
  "solution": {
    "name": "obo-function-consumer-client-side-solution",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "isDomainIsolated": false,
    "webApiPermissionRequests": [
      { "resource": "api://<proxy-app-client-id>", "scope": "access_as_user" }
    ]
  }
}
```

## Step 3 — web part: profile + chat

```ts
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";

const RESOURCE = "api://<proxy-app-client-id>";
const PROXY_BASE = "https://app-proxy-<token>.azurewebsites.net";

interface UserProfile { id: string; displayName: string; jobTitle?: string; department?: string; }
interface ChatReply { reply: string; responseId: string; status: string; consentUrl?: string | null; }

export default class ProfileAgentWebPart /* ... */ {
  private client!: AadHttpClient;
  private previousResponseId: string | null = null;

  public async onInit(): Promise<void> {
    this.client = await this.context.aadHttpClientFactory.getClient(RESOURCE);
  }

  public async loadProfile(): Promise<UserProfile> {
    const res: HttpClientResponse = await this.client.get(
      `${PROXY_BASE}/api/profile`, AadHttpClient.configurations.v1);
    if (!res.ok) throw new Error(`profile ${res.status}: ${await res.text()}`);
    return res.json();
  }

  public async ask(message: string): Promise<string> {
    const res: HttpClientResponse = await this.client.post(
      `${PROXY_BASE}/api/agent/chat`, AadHttpClient.configurations.v1,
      { headers: { "content-type": "application/json" },
        body: JSON.stringify({ message, previousResponseId: this.previousResponseId }) });
    if (!res.ok) throw new Error(`agent ${res.status}: ${await res.text()}`);
    const data: ChatReply = await res.json();
    if (data.status === "consent_required" && data.consentUrl) {
      window.open(data.consentUrl, "_blank");          // user consents once, then call ask() again
    }
    this.previousResponseId = data.responseId;         // keep the thread for multi-turn
    return data.reply;
  }
}
```

Render your own chat box and wire it to `ask()`. The agent calls `get_sharepoint_profile` on the first
turn (as the signed-in user), so it already knows who it's talking to — no need to pass the profile in.

## Admin approval

After uploading `.sppkg` to the tenant App Catalog, an admin approves the request in
**SharePoint Admin Center → Advanced → API access** (`access_as_user` on `api://<proxy-app-client-id>`
for the SharePoint Online Client Extensibility principal).
