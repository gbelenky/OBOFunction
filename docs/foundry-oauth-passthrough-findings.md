# Foundry OAuth identity passthrough — where it works, where it does not

> **TL;DR.** Foundry OAuth identity passthrough for a **custom** remote MCP server works only in an
> **interactive, first‑party channel** (Foundry Playground today; by extension Microsoft Teams and
> Microsoft 365 Copilot). It **cannot** auto‑deliver the signed‑in user's profile through a **custom
> web app** (SPFx → proxy → agent), because a hosted agent discovers and runs toolbox/MCP tools under
> its **own managed identity** — there is no user context at discovery time, so the passthrough tool is
> silently dropped and **no consent event is ever emitted** in the agent's response. The working
> alternative for the custom path is **Option A**: the proxy resolves the user's country itself via OBO
> and injects it as context for the agent's local `search_faq` tool. Option A is implemented in this repo.

This is a **platform/channel reality, not a code defect**. Everything below is evidence‑based — it was
verified with live telemetry, CORS probes, and direct (and streamed) agent‑API calls.

---

## 1. What "OAuth identity passthrough" promises

With OAuth identity passthrough, the signed‑in user's identity is supposed to flow all the way through the
agent to the downstream Azure service via the OBO exchange
([Microsoft Learn — MCP authentication](https://learn.microsoft.com/azure/foundry/agents/how-to/mcp-authentication#oauth-identity-passthrough)).
For a **custom** remote MCP server the documented mechanism is an **OAuth2 connection** (two app
registrations: a client app = the Foundry connection's ID/secret, and a server app = the OBO resource),
demonstrated working **in the Foundry Playground**, where the user signs in and consents in‑browser
([Microsoft Learn — deploy a remote MCP server (OBO)](https://learn.microsoft.com/azure/developer/azure-mcp-server/how-to/deploy-remote-mcp-server-on-behalf-of)).

> `UserEntraToken` is **not** documented for custom MCP servers — it is the first‑party / catalog
> mechanism only. (Confirmed empirically: see §4.)

## 2. Where it WORKS — interactive first‑party channels

The passthrough completes only when the channel can satisfy **all four** requirements of an interactive
channel:

| # | Requirement | Why it matters |
|---|---|---|
| 1 | Per‑user identity reaches the agent at **run time** | The tool must execute as the user, not the agent MI. |
| 2 | A **browser** can render the apihub consent redirect to a registered redirect URI | OAuth passthrough consent is per‑user and interactive. |
| 3 | Consent **persists** per user | So subsequent turns don't re‑prompt. |
| 4 | The tool then **runs under the consented user** | So OBO downstream (Graph/SharePoint) is the user. |

Channels that meet all four:

- **Foundry Playground** — verified live: a direct REST `tools/list` against the toolbox MCP endpoint with
  a **user** `ai.azure.com` token returns `-32006 CONSENT_REQUIRED` **with a real per‑user consent URL**
  (`logic-swedencentral-001.consent.azure-apihub.net/login?…`). The Playground renders that URL, the user
  consents, and the tool then runs as the user.
- **Microsoft Teams** and **Microsoft 365 Copilot** — the same first‑party surfaces. They carry per‑user
  identity into the agent and can render the in‑browser consent loop, so the passthrough is expected to
  work there.

**Answer to "would the desired approach work via Teams / M365 Copilot and not a custom web app?" → YES.**
The auto‑profile‑via‑passthrough design is expected to work through Microsoft first‑party channels because
those surfaces complete the per‑user OAuth consent loop in‑browser **and** run the tool under the user. It
does **not** work for a custom web app / SPFx calling the agent API (see §3).

## 3. Where it FAILS — custom web app (SPFx → proxy → agent)

When a **custom caller** hits the agent API — whether SPFx calls the agent data‑plane **directly** or
**via the proxy** — the passthrough tool is **silently dropped** and **no consent is surfaced**.

**Root cause:** a hosted agent **discovers and executes** its toolbox/MCP tools under the **agent's
managed identity**, which has **no user context**. OAuth identity passthrough consent is strictly
per‑end‑user, and there is no user (and no browser) at the agent's tool‑discovery/run step in this path. So
`tools/list` for the passthrough connection returns `-32006 CONSENT_REQUIRED`, the tool is dropped
(`StrictMode=false`), and the agent simply answers without it.

### Evidence (this repo, live)

| Probe | Result |
|---|---|
| App Insights `invoke_agent` (agent v14, real user token via proxy) | **code=0, success**, ~7.8 s. Local `search_faq` ran; `consentUrl: null`; profile **not** loaded. |
| App Insights `tools/list SharePointProfileTools:2` (OAuth2 passthrough conn) | **500 / `-32006`** — discovery ran under the agent **managed identity**, no user context → tool dropped, no consent relayed. |
| Direct hosted‑agent endpoint with a **user** token (`/agents/SharePointProfileAgent/endpoint/protocols/openai/responses`) | **200 completed**, but `get_sharepoint_profile` **silently dropped**; only `search_faq` ran; **no consent**. |
| Same call **streamed** (`stream=true`), SSE scanned | **Zero** `consent` / `required_action` / `mcp_approval` events. Only `response.output_text.delta`, `function_call`, `message`, `response.completed`. |
| `UserEntraToken` connection (toolbox v3) — obotoken `getuseraccesstoken` | **400 "Audience for incoming token is incorrect"** for every audience form (bare GUID, `api://`, `/.default`, `/access_as_user`), even after adding the AML preauth. A first‑party app (WorkIQ) passes obotoken with the **identical** token → proves the rejection is **custom‑app‑specific**, not test‑method‑specific. |

### CORS is **not** the blocker

The Foundry data plane is CORS‑open. `OPTIONS` on the project Responses endpoint, on the hosted‑agent
endpoint, and on the toolbox MCP endpoint all return `Access-Control-Allow-Origin: *` with
`Access-Control-Allow-Methods: POST`. A browser **can** call these; the blocker is the **MI‑context tool
discovery**, not CORS.

## 4. Why `UserEntraToken` is not an escape hatch

`UserEntraToken` connections are the first‑party/catalog mechanism. For a custom app registration, the
obotoken endpoint rejects the audience regardless of form (see the table in §3). The same token works for a
first‑party caller, so the rejection is bound to the custom app — `UserEntraToken` cannot be coerced into
serving the custom SPFx path.

## 5. The working alternative for the custom path — Option A (implemented)

Because passthrough can't deliver the profile through SPFx → proxy → agent, the **proxy resolves the
user's country itself** and injects it as context. The proxy already holds the user's OBO assertion, so:

```
SPFx ──user JWT──► Proxy
        │  OBO #1: user → SharePoint  → PeopleManager/GetMyProperties → IntranetCountry   (primary)
        │  OBO #2: user → Microsoft Graph → /me?$select=country,usageLocation              (fallback)
        ▼
   Proxy prepends a one‑line country context to the FIRST turn's message, then calls the agent
        ▼
   Agent → local search_faq(country, question) → Azure AI Search faq-index (Location filter + Global)
```

**Implementation (this repo):**

- `src/OBOFunction/Services/ProfileCountryService.cs` (+ `IProfileCountryService.cs`) — OBO #1/#2 above.
  Best‑effort: any failure (missing consent, app‑only identity, 401) logs and returns `null`, so the chat
  still completes with global/default results.
- `src/OBOFunction/Program.cs` — on the **first** turn only (no `previous_response_id`), the endpoint
  resolves the country and prepends:
  `"[Profile context provided by the host: the signed-in user's country is \"<country>\". Treat this as the authoritative country for any country-filtered features such as search_faq; do not ask the user for their country.]"`
  Follow‑up turns inherit the context via `previous_response_id`.

**What stays true to the original constraints:**

- The **proxy never declares, enumerates, or whitelists agent tools.** It injects only a plain country
  **string** as conversation context. It is unaware of `search_faq`, MCP, or any tool wiring.
- `search_faq` **remains a LOCAL in‑process tool** on the hosted agent — unchanged.

**Accepted trade‑off:** Option A makes the proxy *profile‑aware* (it reads one field — the country — via
OBO). This is a deliberate, minimal deviation from "the proxy only calls the agent," chosen because the
standard passthrough path is blocked for custom web apps by the platform reality in §3. The proxy is **not**
tool‑aware.

## 6. Decision matrix — how to get auto‑country per channel

| Channel | Mechanism | Works today? |
|---|---|---|
| Foundry Playground | OAuth identity passthrough (toolbox v2 + OAuth2 connection) | ✅ (interactive consent in‑browser) |
| Microsoft Teams / M365 Copilot | OAuth identity passthrough (first‑party surface) | ✅ expected (same interactive consent model) |
| Custom web app / SPFx (direct or via proxy) | OAuth identity passthrough | ❌ tool dropped under agent MI; no consent emitted |
| Custom web app / SPFx (via proxy) | **Option A** — proxy OBO resolves country, injects context | ✅ **implemented in this repo** |
| Any | User states their country in chat | ✅ interim, always works |

---

### Appendix — environment facts used for verification

- Project `rsc-fdr-swc` / `prj-fdr-swc` (RG `rg-fdr`), App Insights `prj-fdr-swc-appinsights-HRAgent`.
- Agent **`SharePointProfileAgent`** (azd service name `sharepoint-agent`), deployed **v14**.
- Hosted‑agent endpoint form:
  `…/api/projects/prj-fdr-swc/agents/{AgentName}/endpoint/protocols/openai/responses?api-version=2025-05-15-preview`.
- Toolbox `SharePointProfileTools` v2 = OAuth2 `SharePointMCPOAuth` connection (passthrough), default.
- Proxy app `7ce28b8f-…` has delegated SharePoint `AllSites.Read` + `User.Read.All` and Graph `User.Read`,
  which is what makes Option A's OBO #1/#2 possible.
