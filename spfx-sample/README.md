# SPFx Sample — Consuming the Profile API & the SharePoint Hosted Agent

Two things an SPFx web part can call:

1. **`GET /api/profile`** on the Function — the signed-in user's Graph + SharePoint profile (real per-user identity, via OBO).
2. **The SharePoint Hosted Agent** (`src/SharePointAgent`) — a chat agent. **The browser must not talk to the agent directly** (see below), so SPFx calls a small **server-side proxy** that forwards chat to the agent.

## Required SPFx version
- SPFx **1.18+**, TypeScript, React (or no-framework — code is identical)

---

## Why you can't call the hosted agent directly from the browser

A Foundry hosted agent is invoked over the **Responses** data-plane API and requires an **Entra token for the Foundry / Cognitive Services resource** (`https://ai.azure.com/.default` / `https://cognitiveservices.azure.com/.default`). In the browser that is a problem:

- The Foundry data-plane endpoint does **not** send CORS headers for a `*.sharepoint.com` origin, so `fetch` from the web part is blocked.
- Putting an agent **API key** in the client-side bundle = leaked credential.
- The agent's profile tool runs as the **agent's identity** (managed identity), not the end user — so a browser-acquired user token wouldn't help anyway.

➡️ **Embed the agent via a thin server-side proxy.** The web part authenticates to the proxy with `AadHttpClient` (user token); the proxy calls the agent with its **managed identity**. Same auth/permission model you already use for `/api/profile`.

```
SPFx web part ──[AadHttpClient, user JWT]──► Function proxy /api/agent/chat ──[managed identity]──► Hosted Agent (Responses)
                                                                                                        └─ get_sharepoint_profile tool
```

---

## Step 1 — The proxy endpoint (already implemented)

The Function ships a `POST /api/agent/chat` endpoint that validates the SPFx user token (same `api://<client-id>` audience as `/api/profile`), then invokes the deployed hosted agent over the Foundry **Responses** API using `DefaultAzureCredential` (the Function's managed identity). **No agent secret reaches the browser.**

- Implementation: `src/OBOFunction/Functions/AgentChatFunction.cs` + `Services/AgentChatClient.cs`
- Request: `{ "message": "...", "previousResponseId": "<id|null>", "userProfile": <json|null> }`
- Response: `{ "reply": "...", "responseId": "..." }`
- Pass `responseId` back as `previousResponseId` on the next turn for multi-turn chat.

Configure these app settings (see root README → *Agent proxy settings*) and grant the Function's managed identity the **Azure AI User** role on the Foundry project:

| Setting | Value |
|---|---|
| `Foundry__ProjectEndpoint` | `https://<project>.services.ai.azure.com/api/projects/<project>` |
| `Foundry__AgentName` | `SharePointProfileAgent` (from `agent.yaml`) |
| `Foundry__ApiVersion` | `2025-05-01` (default; tune to your project) |
| `Foundry__TokenScope` | `https://ai.azure.com/.default` (default) |

---

## Step 2 — `package-solution.json` (unchanged)

The web part still only needs permission to the **Function API** — not to Foundry, because the proxy fronts the agent:

```json
{
  "solution": {
    "name": "obo-function-consumer-client-side-solution",
    "id": "00000000-0000-0000-0000-000000000000",
    "version": "1.0.0.0",
    "includeClientSideAssets": true,
    "isDomainIsolated": false,
    "webApiPermissionRequests": [
      {
        "resource": "api://<function-app-registration-client-id>",
        "scope": "access_as_user"
      }
    ]
  }
}
```

---

## Step 3 — Web part: profile + agent chat

```ts
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";

const RESOURCE = "api://<function-app-registration-client-id>";
const FUNCTION_BASE = "https://<func-host>.azurewebsites.net";

interface UserProfile {
  id: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
}

interface ChatReply { reply: string; responseId: string; }

export default class ProfileAgentWebPart /* ... */ {
  private client!: AadHttpClient;
  private previousResponseId: string | null = null;

  public async onInit(): Promise<void> {
    this.client = await this.context.aadHttpClientFactory.getClient(RESOURCE);
  }

  // (optional) show the real per-user profile in the UI
  public async loadProfile(): Promise<UserProfile> {
    const res: HttpClientResponse = await this.client.get(
      `${FUNCTION_BASE}/api/profile`,
      AadHttpClient.configurations.v1
    );
    if (!res.ok) throw new Error(`profile ${res.status}: ${await res.text()}`);
    return res.json();
  }

  // chat with the hosted agent through the proxy
  public async ask(message: string): Promise<string> {
    const res: HttpClientResponse = await this.client.post(
      `${FUNCTION_BASE}/api/agent/chat`,
      AadHttpClient.configurations.v1,
      {
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ message, previousResponseId: this.previousResponseId })
      }
    );
    if (!res.ok) throw new Error(`agent ${res.status}: ${await res.text()}`);
    const data: ChatReply = await res.json();
    this.previousResponseId = data.responseId;   // keep the thread for multi-turn
    return data.reply;
  }
}
```

That's the embed: render your own chat box (input + message list) in the web part and wire it to `ask()`. The agent calls its `get_sharepoint_profile` tool on the first turn, so it already knows who it's talking to.

---

## Optional — inject the *real* user's profile into the agent

The hosted agent's embedded tool fetches the profile as the **agent's** managed identity, not the SharePoint user. If you need the agent to reason over the **calling user's** real profile (true "Shape 1"):

1. Web part calls `GET /api/profile` (OBO → real user).
2. Web part passes that JSON in the chat call as `userProfile`:
   `{ message, userProfile: profile, previousResponseId }`.
3. The proxy (`AgentChatClient`) prepends it as a system context message, so the agent uses the real user's data instead of its own tool result.

This keeps user identity flowing without ever handing the browser any agent credentials.

---

## Admin approval

After uploading `.sppkg` to the tenant App Catalog, an admin approves the API permission request in **SharePoint Admin Center → Advanced → API access** (approves `access_as_user` on `api://<function-app-registration-client-id>` for the SharePoint Online Client Extensibility principal).
