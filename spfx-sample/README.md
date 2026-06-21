# SPFx sample — chatting with the profile agent

An SPFx web part calls **one backend** — the proxy (`src/OBOFunction`) — with the signed-in user's
identity via `AadHttpClient`. It is **agnostic**: it holds no profile or greeting logic and never talks
to Foundry directly.

- **`POST /api/agent/chat`** — chat. The proxy validates the user JWT, resolves the user's profile via
  OBO (SharePoint UPS + Graph), injects it as context, and calls the Foundry **hosted agent** as the
  user. The agent greets by name and answers FAQ questions filtered by the user's country.

The browser never talks to Foundry directly.

## Built solution in this repo: `spfx-profile-agent/`

A ready-to-build SPFx **1.23** React web part is scaffolded under
[`spfx-profile-agent/`](./spfx-profile-agent). It is already wired to the live proxy:

| Wiring | Value | Where |
|---|---|---|
| `PROXY_RESOURCE` | `api://7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a` | `src/webparts/profileAgent/services/ProxyClient.ts` |
| `PROXY_BASE` | `https://app-proxy-z6vb2tjg2j4ye.azurewebsites.net` | same |
| `webApiPermissionRequests` | `api://7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a` / `access_as_user` | `config/package-solution.json` |

On mount the web part posts `{ greeting: true }` (no wording) to `POST /api/agent/chat`; the proxy +
agent own the greeting and reply with a one-sentence name greeting. The chat box supports multi-turn
`previousResponseId`.

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
2. Grant the web part access to the proxy API (`access_as_user` on
   `api://7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a`). Use **either**:
   - **Admin UI** — approve the request in **SharePoint Admin Center → Advanced → API access**
     (the SharePoint Online Client Extensibility principal is the caller); **or**
   - **Pre-authorization** — add the SharePoint Online Client Extensibility principal
     (`08e18876-6177-487e-b8b5-cf950c1e598c`) to the proxy app registration's
     `preAuthorizedApplications` for the `access_as_user` scope. This skips the Admin "API access"
     approval step entirely.
3. Add the **ProfileAgent** web part to a page.

> **PnP.PowerShell deploy (scripted, used for this environment).** The tenant App Catalog can be
> provisioned and the package added/published/installed without the Admin UI:
> ```powershell
> Connect-PnPOnline -Url https://<tenant>.sharepoint.com/sites/appcatalog -ClientId <pnp-login-app> -Tenant <tenant>.onmicrosoft.com -Interactive
> Add-PnPApp     -Path spfx-profile-agent.sppkg -Scope Tenant -Publish -Overwrite
> ```
> The proxy API was pre-authorized for the SharePoint Online Client Extensibility principal (option 2
> above), so no SharePoint Admin "API access" approval was required. Interactive login is required —
> device-code login does not complete the SharePoint admin sign-in.

The blueprint below documents the contract and code that this solution implements.

## Required SPFx version

SPFx **1.18+**, TypeScript, React (or no-framework — the code is identical).

## Why the browser can't call the agent directly

Foundry's Responses data-plane API needs an **Entra token for the AI resource**
(`https://ai.azure.com/.default`). From a `*.sharepoint.com` page that's blocked three ways:

- the Foundry endpoint sends **no CORS headers** to a SharePoint origin → `fetch` fails;
- an **agent/API key** in the client bundle would leak;
- the user's identity must reach Graph/SharePoint via **OBO**, which only a confidential server can do.

So the proxy fronts the call. It is the one confidential client holding the OBO secret and performing
every OBO exchange.

```
SPFx ──[AadHttpClient, user JWT]──► proxy POST /api/agent/chat
        FIRST turn: proxy OBO user→SharePoint + user→Graph ⇒ build profile, inject as context
                    proxy OBO user→https://ai.azure.com/.default ⇒ call hosted agent as the user
             ▼
        Foundry Hosted Agent (SharePointProfileAgent) ──local search_faq──► Azure AI Search
```

## Step 1 — the proxy endpoint (already implemented)

`POST /api/agent/chat` (in `src/OBOFunction/Program.cs`, logic in `Services/AgentChatClient.cs`)
validates the SPFx user token (audience `api://<proxy-app>`), then on the **first turn**:

- resolves the user's profile via OBO (`Services/ProfileContextService.cs`) — SharePoint UPS
  `GetMyProperties` (custom `IntranetCountry`) + Graph `/me` — and injects it as a developer-role
  `USER_PROFILE_JSON` context item;
- OBO-exchanges the user token for `https://ai.azure.com/.default` and calls the hosted agent via
  `agent_reference` **as the user**.

Request / response contract:

```jsonc
// POST /api/agent/chat
{ "message": "How can I request a vacation?", "previousResponseId": null }
// or, for the auto-greeting on mount:
{ "greeting": true }

// 200 OK
{ "reply": "...", "responseId": "resp_...", "status": "completed" }
```

Pass `responseId` back as `previousResponseId` on the next turn (server-side multi-turn state). The
profile is resolved and injected only on the first turn; follow-ups inherit it via `previousResponseId`.

App settings (and grant the proxy's managed identity the **Azure AI User** role on the Foundry project):

| Setting | Value |
|---|---|
| `Foundry__ProjectEndpoint` | `https://<acct>.services.ai.azure.com/api/projects/<project>` |
| `Foundry__AgentName` | `SharePointProfileAgent` |
| `AzureAd__*` | proxy app registration + secret (the OBO confidential client) |

## Step 2 — `package-solution.json`

The web part only needs permission to the **proxy API** (its single endpoint):

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

## Step 3 — web part: agnostic chat with auto-greeting

```ts
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";

const RESOURCE = "api://<proxy-app-client-id>";
const PROXY_BASE = "https://app-proxy-<token>.azurewebsites.net";

interface ChatReply { reply: string; responseId: string; status: string; }

export default class ProfileAgentWebPart /* ... */ {
  private client!: AadHttpClient;
  private previousResponseId: string | null = null;

  public async onInit(): Promise<void> {
    this.client = await this.context.aadHttpClientFactory.getClient(RESOURCE);
  }

  private async post(body: object): Promise<ChatReply> {
    const res: HttpClientResponse = await this.client.post(
      `${PROXY_BASE}/api/agent/chat`, AadHttpClient.configurations.v1,
      { headers: { "content-type": "application/json" }, body: JSON.stringify(body) });
    if (!res.ok) throw new Error(`agent ${res.status}: ${await res.text()}`);
    const data: ChatReply = await res.json();
    this.previousResponseId = data.responseId;          // keep the thread for multi-turn
    return data;
  }

  public async ask(message: string): Promise<string> {
    return (await this.post({ message, previousResponseId: this.previousResponseId })).reply;
  }

  // Fire once on mount: the proxy + agent own the greeting. The client sends no wording.
  public async greet(): Promise<string> {
    return (await this.post({ greeting: true })).reply;
  }
}
```

Render your own chat box and wire it to `ask()`, and call `greet()` on mount. The proxy resolves the
profile and the agent composes the greeting — the client passes no profile and no greeting wording.
Render the reply with `white-space: pre-wrap` so line breaks display.

## Admin approval

After uploading `.sppkg` to the tenant App Catalog, an admin approves the request in
**SharePoint Admin Center → Advanced → API access** (`access_as_user` on `api://<proxy-app-client-id>`
for the SharePoint Online Client Extensibility principal).
