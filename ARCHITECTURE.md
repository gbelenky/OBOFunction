# Personalized & Secure Intranet AI Assistant
## Architecture (focused: ingestion · search · request flow)

---

## 1. What the system does
A SharePoint-embedded assistant answers user questions over intranet content. **The orchestration tier — not the LLM — enforces SharePoint ACLs by attaching a security filter to every search.**

---

## 2. The three protagonists

```
┌──────────────┐    user JWT + question     ┌──────────────────┐
│   SPFx       │ ─────────────────────────► │  Orchestration   │
│  web part    │                            │   Function       │
│              │ ◄─────── answer ─────────  │  (security       │
└──────────────┘                            │   boundary)      │
                                            └────────┬─────────┘
                                                     │ search query
                                                     │ + ACL filter
                                                     │ + profile scope
                                                     ▼
                                            ┌──────────────────┐
                                            │  Azure AI Search │
                                            │   index with     │
                                            │   per-doc ACL    │
                                            └──────────────────┘
                                                     ▲
                                                     │ writes
                                            ┌────────┴─────────┐
                                            │   Ingestion      │
                                            │   pipeline       │
                                            │ (SharePoint →    │
                                            │  AI Search)      │
                                            └──────────────────┘
```

---

## 3. Ingestion plane

A **separate pipeline** — runs independently of user traffic — projects SharePoint content into Azure AI Search.

### What it does per document
1. **Crawl** SharePoint sites/lists (incremental delta every 15 min; full reconcile nightly)
2. **Extract content** — text, basic metadata (title, URL, author, last-modified, site, language, document type)
3. **Extract permissions** — the actual security trustees on the item:
   - Users with explicit access → `aclAllowUsers` (Entra Object IDs)
   - Groups with access → `aclAllowGroups` (group Object IDs, including transitive expansion)
   - Deny entries → `aclDenyGroups`
   - "Everyone except external users" → flagged
   - Sharing-link state and sensitivity-label markers
4. **Chunk + embed** — split content, generate vector embeddings
5. **Push** the document + ACL fields to AI Search

### Index schema (key fields)

| Field | Purpose |
|---|---|
| `id`, `title`, `content`, `contentVector`, `url`, `lastModified` | Retrieval + display |
| `siteId`, `siteUrl`, `department`, `documentType`, `language` | Profile-driven scoping |
| `aclAllowUsers`, `aclAllowGroups`, `aclDenyGroups`, `isPublic` | **Security trim — the architectural floor of access control** |

### Why ingestion is the riskiest plane
The system's security guarantee is **only as good as the ACL extraction**. Edge cases that must be handled correctly:
- Unique permissions at item/folder level (most leaks happen here)
- Sharing links (anonymous, organization-wide, specific people)
- Broken inheritance
- Sensitivity-label-driven access
- External users / B2B guests

Building this in-house is months of work; commercial connectors (Glean, BA Insight, Coveo) exist specifically because this is hard.

### Operating cadence trade-off
- **Faster crawl** → fresher ACLs → higher cost
- **Slower crawl** → cheaper → larger lag between SharePoint change and effective enforcement
- Recommended baseline: **15-min delta + nightly full** = ≤15-minute permission lag

---

## 4. Search plane

**Azure AI Search**, queried on every user turn by the orchestration tier — never by the LLM, never by the browser.

### What goes into the query — and where it comes from

Three distinct inputs combine into every search request. Keeping them separate is the architecture's core security property.

| Input | Source | Fetched when | Role |
|---|---|---|---|
| **Identity** (`userOid` + transitive group OIDs) | Validated JWT claims (+ Graph `transitiveMemberOf` on overage) | Every request, server-side | **Security trim** — matched against `aclAllow*` fields on docs |
| **Profile** (department, country, skills, AboutMe…) | OBO → Graph `/me` + SharePoint UPS | Every request, server-side | **Relevance & business filters** (e.g., `department eq 'Finance'`) + agent reasoning input |
| **Document ACL** (`aclAllowUsers`, `aclAllowGroups`, `aclDenyGroups`, `isPublic`) | Resolved by the ingestion pipeline per document | Index time (15-min delta) | **The access contract** — baked into the index, not the request |

> **Profile ≠ ACL.** The `/api/profile` call never fetches "what the user can see." That would be unbounded and instantly stale. Authorization is computed by *intersecting the user's current identity claims with the per-document ACLs already in the index* — every request, no caching.

### Mental model
- **Profile = what the user *is*** → drives relevance + optional business filters
- **Identity + groups = who the user *is*** → drives the ACL trim
- **ACL fields on docs = who can see *this thing*** → the security floor

The Search filter AND-combines all three. Compromising the profile cannot widen access; only validated identity claims can.

### Query shape
```
search:  <user question (or rewritten standalone form)>
filter:  (
           isPublic eq true
           or aclAllowUsers/any(u: u eq '<userOid>')
           or aclAllowGroups/any(g: search.in(g, '<group1>,<group2>,...'))
         )
         and not aclDenyGroups/any(g: search.in(g, '<group1>,<group2>,...'))
         and department eq '<userDept>'      // optional profile scope
queryType: semantic + vector hybrid
top: 5–10
```

### Architectural invariants
- The filter is **always attached** — there is no code path in the orchestration tier that issues a search without one.
- Filter values come **only from validated token claims and OBO-fetched profile data** — never from user prompts or LLM output.
- A user with no group OIDs and no public matches returns **zero results**, not "fallback to broader scope."

### Filter scaling concerns
- **Token group overage** (>200 groups): fall back to Graph `/me/transitiveMemberOf`
- **Filter expression cap** (~32 KB): for users in thousands of groups, chunk into multiple search calls or maintain a compact per-user access list

---

## 5. Request flow — SPFx ↔ Function ↔ Search

```
(1)  User types question in SPFx web part
                │
(2)  SPFx acquires user token (AadHttpClient,
     SharePoint Online Client Extensibility SP)
                │
(3)  SPFx POSTs to Function:
                Authorization: Bearer <user token>
                { prompt, threadId? }
                │
                ▼
     ┌──────────────────────────────────────┐
     │ Orchestration Function — per request │
     │                                      │
(4)  │  Validate JWT (audience, issuer,     │
     │    signature, lifetime)              │
     │                                      │
(5)  │  OBO #1 → Graph → /me + transitive   │
     │    group OIDs                        │
     │                                      │
(6)  │  OBO #2 → SharePoint UPS → extended  │
     │    profile (skills, dept, etc.)      │
     │                                      │
(7)  │  If threadId present, rewrite        │
     │    question to standalone form       │
     │                                      │
(8)  │  Build filter:                       │
     │    ACL trim (userOid + groupOids)    │
     │    AND profile scope (dept, etc.)    │
     │                                      │
(9)  │  Search AI Search:                   │
     │    question + filter → top N chunks  │
     │                                      │
(10) │  Append question + chunks to         │
     │    Foundry thread; run agent         │
     │                                      │
(11) │  Return { answer, threadId,          │
     │           citations }                │
     └──────────────────┬───────────────────┘
                        │
(12) SPFx renders answer + citations,
     stores threadId for next turn
```

**Per-turn invariants** — steps 4–9 run on *every* request, including every follow-up. No identity, profile, or filter is cached across turns. A user who loses group membership sees the consequence on the next question, against the latest indexed permissions.

---

## 6. Why the LLM cannot subvert the filter
- The orchestration tier owns the only code path that calls AI Search.
- The agent has no AI Search credentials, no Graph credentials, no tools that touch user data.
- A prompt-injected document at worst corrupts the answer text — it cannot expand the result set, cannot re-issue a search, cannot leak content the user wasn't allowed to retrieve.

The agent is a **reasoning component**, not a **data-access component**. That separation is the architecture's core security property.

---

## 7. Honest limits
- **Permission freshness** = max(token TTL, index crawl interval). ≤15 min with recommended settings.
- **ACL fidelity** = whatever the ingestion pipeline captures. This is the architectural floor — invest accordingly.
- **Scale of group membership** matters for filter construction; thousands-of-groups users need a different strategy than dozens.

---

## 8. Foundry agent auth matrix — does Agent Service support OBO?

A recurring design question is *"can the Foundry agent itself act on behalf of the calling user, so I don't need OBO in the Function?"* The answer drives the **Shape 1** decision. The text below is taken **verbatim** from current Microsoft Learn documentation.

### 8.0 Hosted agents DO support attended OBO natively (headline)

> **Source — Hosted agents in Foundry Agent Service (preview)**
> https://learn.microsoft.com/azure/foundry/agents/concepts/hosted-agents#agent-identity-and-endpoint

Every hosted agent gets a **dedicated Microsoft Entra ID (agent identity)** auto-created at deploy time. Verbatim:

> When integrated via Microsoft 365 channels (for example, Teams), Hosted agents can operate in two identity modes depending on how they are invoked:
>
> - **User-invoked scenarios (interactive)**: If a user token is present, the platform supports OAuth 2.0 On-Behalf-Of (OBO) flows. In this case, the agent can call downstream services on behalf of the user using the user's delegated permissions, subject to Microsoft Entra ID tenant policies.
> - **Autonomous or background scenarios**: If no user token is available, the agent authenticates using its own Microsoft Entra ID (agent identity), typically via managed identity, to access downstream services.
>
> In both cases, the agent retains its dedicated Microsoft Entra ID for authentication, authorization, and auditability.

Tool-level user-identity passthrough is delivered through **Toolbox in Foundry**. Verbatim (same doc, *Toolbox in Foundry*):

> We recommend customers using toolbox in Foundry for connecting tools in Hosted agent with consolidated auth support across **OAuth Identity passthrough**, agent identity, key based and more.

**What this means for us:** the earlier claim that *"the bare hosted endpoint never forwards a user token, so the agent can't OBO for a raw caller"* was **wrong**. A hosted agent **does** perform attended OBO **when a user token is present at invocation**. The two documented ways to make a user token "present" are: (a) **M365 channel integration** (Teams/Copilot SSO supplies it), or (b) a **caller that presents the user token** to the agent's Responses endpoint. The recommended way for the agent's *tools* to then call Graph/SharePoint as the user is **Toolbox OAuth identity passthrough**, not `DefaultAzureCredential`.

> **Status:** The hosted-agents platform itself is **preview** — the doc title is literally *"Hosted agents in Foundry Agent Service (preview)"* and the publishing/agent-application packages are *"currently in preview"* (*Agent applications in Microsoft Foundry*). The **OBO behaviour is GA-documented as a concept**; the **hosting platform that delivers it is preview**.

### 8.1 Agent Service DOES support OBO (attended) — Agent ID framework

> **Source — Agent identity concepts in Microsoft Foundry**
> https://learn.microsoft.com/azure/foundry/agents/concepts/agent-identity#key-concepts

Verbatim:

> The Agent ID platform framework introduces formal *agent identities* and *agent identity blueprints* in Microsoft Entra ID to represent AI agents.

> Agent identities support two key authentication scenarios:
>
> 1. **Attended (delegated access or on-behalf-of flow)**: The agent operates on behalf of a human user, using the OAuth 2.0 on-behalf-of (OBO) flow. The user first authenticates to the application, and the application passes the user's token to Agent Service. Agent Service then exchanges that token for one that carries both the agent identity and the user's delegated permissions. This approach means the agent can only access resources that the user has consented to and is authorized for.
> 2. **Unattended (application-only flow)**: The agent acts under its own authority, using the OAuth 2.0 client credentials flow. Agent Service authenticates the blueprint to Microsoft Entra ID, obtains a token for the agent identity, and requests a scoped access token for the downstream resource.

Verbatim (initiation constraint, same doc):

> [Agents] can't initiate interactive authorization (`/authorize`) flows directly. They must receive a user token from a client application and then perform an OBO token exchange. A web redirect URI can be configured on a blueprint for consent flows only (`response_type=none`), but it has limited functionality compared to a redirect URI on an app registration.

> **Status (evidence-based, checked 2026-06-18):** The Agent identity concepts page that carries the attended/unattended OBO text is published as a **public conceptual article** (`permissioned-type: public`, `ms.date: 2026-04-13`) with **no preview banner**. So the **attended OBO concept itself is not documented as preview**. The broader **Microsoft Entra Agent ID** platform is new and evolving — verify GA status before taking a hard dependency. The only adjacent OBO paths that *are* explicitly marked **(preview)** in their doc titles are **Logic Apps connector OBO** and **incoming A2A** (see §8.2).

### 8.2 Tool-level OBO paths (verbatim)

**MCP server OAuth identity passthrough** — *Connect an MCP server on Azure Functions to a Foundry Agent Service agent*
https://learn.microsoft.com/azure/foundry/agents/how-to/mcp-authentication#oauth-identity-passthrough

> When you use OAuth identity passthrough, the agent prompts the user to sign in and then uses the returned access token when connecting to the server.

> A Foundry agent connects to the Azure MCP Server by using OAuth identity passthrough. In this mode, the signed-in user's identity flows through all the way to the Azure service calls via the OBO exchange.

**Logic Apps connector OBO** — *Authorize agent tool access to resources with on-behalf-of (OBO) flow in Azure Logic Apps (preview)*
https://learn.microsoft.com/azure/logic-apps/set-up-on-behalf-of-user-flow

> This flow passes the signed-in chat user's identity and permissions through the request chain so that resource connections use the person's identity and permissions to gain access. … OBO is also known as *user context* because agent tools apply the signed-in user's specific security context, including personalized licensing and data access.

**Built-in / OpenAPI tools (no user passthrough)** — *Agent tools overview — Manage authentication for tools*
https://learn.microsoft.com/azure/foundry/agents/concepts/tool-catalog#manage-authentication-for-tools

> OpenAPI tools support anonymous, API key, and managed identity authentication.

### 8.3 Summary matrix

| Path | OBO / user passthrough | Status | Who initiates the exchange |
|---|---|---|---|
| **Hosted agent native attended OBO** (user token present at invocation → platform does OBO) | ✅ Yes | Behaviour GA-documented; **hosting platform is preview** | The platform, once a **user token is present** (M365 channel SSO, or a caller that presents it) |
| **Hosted agent Toolbox OAuth identity passthrough** (tool calls Graph/SharePoint as the user) | ✅ Yes | Toolbox in Foundry; platform preview | Toolbox brokers the signed-in user's token to the tool |
| **Agent ID "attended" OBO** (app passes user token → Agent Service exchanges it) | ✅ Yes | Public concept doc — **not marked preview**; verify GA before hard dependency | A **client application** that already holds the user token |
| **MCP server OAuth passthrough** | ✅ Yes | Not marked preview in how-to | Agent prompts user to sign in (interactive) |
| **Logic Apps connector OBO** | ✅ Yes | ✅ **Preview** (doc title) | Per-user connection consent |
| **Incoming A2A** | n/a | ✅ **Preview** (doc title) | — |
| **Built-in / OpenAPI tools** | ❌ No | GA | n/a (anonymous / API key / managed identity) |

### 8.4 Two valid shapes — Shape 1 (Function OBO) vs Shape 2 (hosted-agent native OBO)

Both are now confirmed supported. The choice is a **maturity / blast-radius** trade-off, not a capability gap.

**Shape 1 — Function-side OBO (GA, current default).** The Function is a confidential client that does OBO → Graph/SharePoint directly; the agent receives the profile as input and never sees user credentials. Relies entirely on **GA** libraries (`Microsoft.Identity.Web` + Graph SDK). Single security boundary (Sections 5–6). Works today regardless of hosted-agent preview status.

**Shape 2 — Hosted-agent native attended OBO (preview platform).** A user token is presented at agent invocation; the platform performs OBO and the agent's profile tool calls Graph/SharePoint **as the user** via **Toolbox OAuth identity passthrough** (replacing `DefaultAzureCredential`). Cleaner "agent acts as the user" model, but depends on **preview** hosting (`hosted-agents`, agent-application publishing) and the user token must be made *present* at invocation — easiest via **M365 channel (Teams) SSO**, or by a caller (the Function proxy) forwarding the validated user token to the agent's Responses endpoint.

- For our SPFx scenario, the cleanest **GA** path remains **Shape 1**. **SPFx** is a public, third-party SharePoint client and cannot itself be the confidential token-broker; a backend (the **Function**) or an **M365 channel** must supply the user token.
- **Shape 2 is now the documented target for "agent acts as the user"** and is viable to pilot — just track the **preview** status of the hosting platform before a hard production dependency.

**Decision for this solution:** ship **Shape 1** as the GA baseline (Function performs OBO → Graph/SharePoint, agent receives profile as input), and treat **Shape 2 (hosted-agent native OBO via Toolbox passthrough)** as the forward-looking path to adopt once the hosted-agent platform reaches GA in our region. To pilot Shape 2 now: present the user token at invocation (Teams channel SSO, or Function proxy forwarding the token) and switch the agent's profile tool from `DefaultAzureCredential` to Toolbox OAuth identity passthrough.

> **Correction notes:** (1) An earlier revision stated Agent Service does *not* support OBO at all — inaccurate; it supports **attended OBO** via the Agent ID framework, **hosted-agent native OBO** (user token present → platform OBO), plus MCP-OAuth and Logic Apps connector OBO. (2) A later revision labelled the Agent ID attended OBO concept **preview** — that was a misattribution; the "(preview)" belongs to the neighbouring **Logic Apps connector OBO** and **A2A** docs, and to the **hosted-agents hosting platform** — *not* to the attended-OBO concept itself. (3) An earlier revision claimed the bare hosted-agent endpoint can never OBO for a non-channel caller — **wrong**; the platform does OBO whenever a **user token is present**, so a caller that forwards the validated user token enables it. The Agent identity concepts page is a **public conceptual article with no preview banner**.

### 8.5 Shape 2 as implemented in this repo (SPFx → Function proxy → toolbox passthrough)

Because we have **no Teams/M365 channel**, the user token is made *present at invocation* by the **Function proxy**, which forwards the signed-in user's identity to the agent's Responses endpoint. The agent's *tools* then call Microsoft 365 as the user through **Toolbox OAuth identity passthrough**. Shape 1 stays the default; Shape 2 is enabled purely by configuration.

**Components created / wired:**

| Piece | Where | Detail |
|---|---|---|
| **Foundry toolbox** `SharePointProfileTools` (v1) | project `prj-fdr-swc` | One `mcp` tool `SharePointProfile` → `server_url` (the MCP server) + `project_connection_id: sp-profile-obo`. Created via REST with header `Foundry-Features: Toolboxes=V1Preview`. |
| **Proxy passthrough** | `Services/AgentChatClient.cs` (`ChatViaToolboxAsync`) | When `Foundry:ToolboxMcpUrl` is set, OBO-exchanges the SPFx user token for an **AI-audience** user token (`https://ai.azure.com/.default`) and calls Responses **as the user**, attaching the toolbox as an `mcp` tool. |
| **Consent ceremony** | `AgentChatClient.ParseReply` / `TryFindConsentLink` | Detects an `oauth_consent_request` in `output[]`, returns `{ status:"consent_required", consentUrl }`; SPFx opens the link and re-sends with `previousResponseId`. |
| **Reply contract** | `Models/AgentChatModels.cs` | `AgentChatReply` gains `Status` (`completed`\|`consent_required`) and `ConsentUrl`. |

**Request-time tool injection vs. agent-definition tool.** We attach the toolbox at the **proxy/request layer** (`tools:[{type:"mcp",…}]` in the Responses payload) rather than baking it into the hosted-agent definition. This keeps the deployed agent unchanged (the in-process `get_sharepoint_profile` tool remains the Shape 1 / GA fallback that preserves the **custom SharePoint UPS fields** — Country, Language, Skills, Interests) and lets Shape 2 be toggled with one app setting.

**Toolbox consumer endpoint** (set as `Foundry:ToolboxMcpUrl` to route through the toolbox rather than a raw MCP server):

```
https://rsc-fdr-swc.services.ai.azure.com/api/projects/prj-fdr-swc/toolboxes/SharePointProfileTools/mcp?api-version=v1
```

**Two sub-options for the toolbox's MCP server (`sp-profile-obo` connection target):**

1. **Managed Microsoft 365 MCP server (true passthrough, recommended for standard profile).** Point the toolbox tool at an Agent 365 server (e.g. *Microsoft 365 User Profile* / *SharePoint & OneDrive*, the same family as the tenant's existing `WorkIQMail` → `agent365.svc.cloud.microsoft/agents/servers/mcp_MailTools`). The platform brokers the user's token; **no custom code, no OBO**. Trade-off: returns standard Graph-shaped profile and likely **not** the custom SharePoint UPS fields.
2. **This Function's own `/api/mcp` server (preserves custom UPS fields).** A custom MCP server that internally runs the existing OBO → Graph + SharePoint UPS merge. Architecturally this is **Shape 1 relocated behind MCP** (the passthrough token has *this API's* audience and the server must OBO downstream — per `mcp-authentication`: *"Don't design your MCP server to rely on passthrough of its authentication token to a downstream Microsoft service"*), but it surfaces the full custom profile through the toolbox.

**Remaining manual (interactive) steps — cannot be automated headless:**

- Create the **`sp-profile-obo` connection** in the project with `--auth-type oauth2` (or `user-entra-token`) pointing at the chosen MCP server.
- **Admin/user consent**: the signed-in user (and, for option 2, the OBO app's delegated Graph/SharePoint scopes) must be consented once; the first run returns a `consent_link`.
- The user must hold the **Foundry User** RBAC role on `prj-fdr-swc`, and the user tenant must match the project tenant (no cross-tenant passthrough).
- Pre-authorize the **SharePoint Online Client Extensibility** SP (`08e18876-6177-487e-b8b5-cf950c1e598c`) on the OBO app so SPFx can mint the inbound token (currently only the Azure CLI app is pre-authorized for local testing).

> **Status: all Shape 2 components are preview** (Toolbox `V1Preview`, hosted-agents platform, Agent 365 managed MCP servers may be Frontier-gated). Keep Shape 1 as the GA fallback; flip `Foundry:ToolboxMcpUrl` to pilot Shape 2.
