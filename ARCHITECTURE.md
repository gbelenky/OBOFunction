# Architecture — identity & OBO design

This document explains **how the signed-in user's identity flows end-to-end** in this repo, and the
security reasoning behind it. For setup/run steps see [`README.md`](./README.md).

> **Scope.** What this repo *implements* is **§1 – §3** below: the SharePoint-profile-for-a-Foundry-agent
> OBO chain. **§4 (Broader vision)** sketches a larger "secure intranet AI assistant" (content ingestion,
> Azure AI Search, per-document ACL trimming) that this profile service would slot into — it is **context,
> not code in this repo**. **§5** is a reference matrix of which Foundry agent auth paths support OBO.

---

## 1. The three components and the one identity chain

| Component | Role |
|---|---|
| **SPFx web part** | Auth boundary. Mints the user JWT (`AadHttpClient`, aud `api://<proxy-app>`). Holds no secrets. |
| **Proxy** (`src/OBOFunction`, App Service) | The only confidential client. Validates the JWT, performs OBO, brokers the Foundry call. |
| **MCP server** (`src/SharePointMcp`, App Service) | Standalone MCP server; `get_sharepoint_profile` does its own OBO → Graph + SharePoint UPS *as the user*. |
| **Hosted agent** (`src/SharePointAgent`, Foundry) | Playground/autonomous demo; references the MCP server as a Foundry-native `mcp` tool. |

Two request paths, both carrying the **same** signed-in user identity:

```
(A) profile read
SPFx ──user JWT──► Proxy ──OBO──► Graph /me + SharePoint UPS ──► UserProfile

(B) chat with per-user tool
SPFx ──user JWT──► Proxy
   leg ① proxy app token (DefaultAzureCredential, https://ai.azure.com/.default) authorizes the call
   leg ② OBO: user JWT → MCP-audience token (api://<mcp-app>/.default), attached to the mcp tool
        ▼
   Foundry model Responses (gpt-4.1-mini)  ──forwards the per-user token──►  SharePointMcp /mcp
                                                                              ──OBO──► Graph + SharePoint UPS
```

**Why the proxy exists (three independent reasons):**
1. **SPFx can't be a confidential client.** It's a public, third-party SharePoint app; it cannot safely hold the OBO client secret.
2. **No CORS from the Foundry data plane** to a `*.sharepoint.com` origin — the browser can't call Foundry directly.
3. **Agent credentials must stay server-side.** A key or app token in the browser would leak.

The proxy is therefore the single place that holds the OBO secret and the Foundry app identity. The user
JWT and every OBO access token are validated, used, and discarded — never logged, never stored.

## 2. Two app registrations (and why exactly two)

| App reg | Audience | OBO performed by it | Why it's needed |
|---|---|---|---|
| **Proxy API** | `api://<proxy-app>` | user → Graph/SharePoint (path A); user → MCP audience (path B leg ②) | The token SPFx mints; the proxy is its confidential client. |
| **MCP API** | `api://<mcp-app>` | user → Graph/SharePoint (as the user) | A distinct audience so the proxy's OBO token can target the MCP server, which then OBOs downstream. |

A **separate MCP audience** is mandatory, not cosmetic: per Microsoft's MCP-authentication guidance, an MCP
server must own its own audience and re-exchange (OBO) for downstream Microsoft services rather than passing
its inbound token through. Two audiences = two clean OBO hops, each consentable and auditable.

Pre-authorizations: the **SharePoint Online Client Extensibility** principal
`08e18876-6177-487e-b8b5-cf950c1e598c` is pre-authorized on the **Proxy API** (so SPFx can mint the inbound
token); the **Proxy API** app is pre-authorized on the **MCP API** (so the proxy's OBO token is accepted).
Both app secrets live in Key Vault, surfaced as config via managed identity.

## 3. Why custom SharePoint UPS fields require the user's identity

Custom User Profile Service attributes (Country, Language, Skills, Interests, Responsibilities, past
projects…) are readable only through a **SharePoint-consented token**:

- **User mode** (path A, and path B once the user token reaches the MCP server) → `GetMyProperties` with an
  OBO user token returns the full UPS including custom `ExtendedProperties`. Proven by
  `scripts/test-spfx-chain.ps1` (`"ResolvedVia": "user"`, real `country`, skills, interests).
- **App-only mode** (a hosted agent's managed identity with no user token, or `DefaultAzureCredential` =
  Azure CLI app) lacks SharePoint delegated consent → `GetMyProperties` 401s → custom fields are empty. This
  is an **identity/consent** outcome, by design — not a code bug.

> **UI note:** custom UPS properties no longer render in the modern SPO profile/Delve UI (the classic profile
> page is retired). The values live in the UPS and are retrievable by API only — which is exactly what this
> service does.

---

## 4. Broader vision (context — not implemented in this repo)

This profile service is designed to plug into a larger **personalized & secure intranet AI assistant**, where
the **orchestration tier (not the LLM) enforces SharePoint ACLs** on every search. That broader system adds an
ingestion pipeline and an Azure AI Search index; it is described here only to show where the profile fits.

### 4.1 The three protagonists (broader system)

```
SPFx web part ──user JWT + question──► Orchestration Function (security boundary)
                                              │ search query + ACL filter + profile scope
                                              ▼
                                       Azure AI Search (index with per-doc ACL)
                                              ▲ writes
                                       Ingestion pipeline (SharePoint → AI Search)
```

### 4.2 Three distinct inputs to every search (keep them separate)

| Input | Source | Role |
|---|---|---|
| **Identity** (`userOid` + transitive group OIDs) | Validated JWT claims (+ Graph `transitiveMemberOf` on overage) | **Security trim** vs `aclAllow*` fields |
| **Profile** (department, country, skills…) | OBO → Graph `/me` + SharePoint UPS — **this repo** | **Relevance / business filters** + agent reasoning input |
| **Document ACL** (`aclAllowUsers/Groups`, `aclDenyGroups`, `isPublic`) | Ingestion pipeline, per document | **The access contract**, baked into the index |

- **Profile = what the user *is*** → relevance + optional business filters.
- **Identity + groups = who the user *is*** → the ACL trim.
- **ACL on docs = who can see *this thing*** → the security floor.

The Search filter AND-combines all three. Compromising the profile cannot widen access; only validated
identity claims can. Authorization is computed by intersecting current identity claims with per-document ACLs
on **every** request — no caching, so a user who loses group membership sees it on the next turn.

### 4.3 Why the LLM cannot subvert the filter

The orchestration tier owns the only code path that calls AI Search. The agent has no Search credentials, no
Graph credentials, and no tools that touch user data — it is a **reasoning component, not a data-access
component**. A prompt-injected document can at worst corrupt answer text; it cannot expand the result set or
leak content the user wasn't allowed to retrieve.

### 4.4 Honest limits

- **Permission freshness** = max(token TTL, index crawl interval) — ≤15 min with recommended settings.
- **ACL fidelity** = whatever the ingestion pipeline captures (the architectural floor).
- **Group-membership scale** affects filter construction (token overage >200 groups → `transitiveMemberOf`;
  ~32 KB filter cap → chunk into multiple searches).

---

## 5. Reference — does Foundry Agent Service support OBO?

The implemented design (the proxy does OBO; the MCP server does OBO as the user) is the **GA** path and does
not depend on any preview platform feature. The matrix below records which Foundry agent auth paths *can*
carry the user's identity, drawn verbatim from Microsoft Learn — useful when evaluating alternatives.

> **Agent identity concepts** — <https://learn.microsoft.com/azure/foundry/agents/concepts/agent-identity>
> > Agent identities support two authentication scenarios:
> > 1. **Attended (OBO)**: the agent operates on behalf of a human user via OAuth 2.0 OBO. The application
> >    passes the user's token to Agent Service, which exchanges it for one carrying both the agent identity
> >    and the user's delegated permissions.
> > 2. **Unattended (application-only)**: the agent acts under its own authority via client-credentials.
> >
> > [Agents] can't initiate interactive `/authorize` flows directly. They must receive a user token from a
> > client application and then perform an OBO token exchange.

> **MCP server OAuth identity passthrough** — <https://learn.microsoft.com/azure/foundry/agents/how-to/mcp-authentication#oauth-identity-passthrough>
> > When you use OAuth identity passthrough … the signed-in user's identity flows through all the way to the
> > Azure service calls via the OBO exchange.

| Path | User passthrough | Notes |
|---|---|---|
| **Proxy-side OBO** (this repo, path A + path B leg ②) | ✅ | GA libraries (`Microsoft.Identity.Web` + Graph SDK). Single confidential client. |
| **MCP server OAuth identity passthrough** (Foundry connection forwards the user token) | ✅ | The documented way for the *Playground/hosted-agent* path to reach the MCP server as the user. Requires a Foundry **OAuth** connection + per-user consent + **Foundry User** role, same tenant. |
| Hosted-agent native attended OBO (user token present at invocation) | ✅ | Behaviour GA-documented; hosting platform is preview. |
| Built-in / OpenAPI tools | ❌ | Anonymous, API key, or managed identity only — no user passthrough. |

### 5.1 Playground path — the one remaining manual step

Path B (proxy-brokered chat) works today end-to-end. The **Foundry Playground** path (where the hosted agent
calls the MCP server directly) needs the project's MCP connection configured as **OAuth identity
passthrough** (Custom → MCP, client = MCP app `api://<mcp-app>`, tenant authorize/token URLs, scopes
`api://<mcp-app>/access_as_user offline_access`, then add the portal redirect URL to the MCP app
registration). Without it, the Playground call reaches the MCP server with no user token and the per-user
fields are empty. This is a one-time portal action; the MCP server code is already correct (proven by
`scripts/test-spfx-chain.ps1`).

### 5.2 Design history (resolved)

Earlier iterations explored a "Shape 1 vs Shape 2" split and a Foundry **Toolbox** wrapper. Those are
**superseded**: there is now one architecture — proxy-side OBO for the SPFx path, and native MCP OAuth
identity passthrough for the Playground path. No Toolbox, no embedded profile tool inside the agent, no
`/api/ask`. Any leftover Foundry *toolbox* or extra *connections* in the project are orphans and can be
removed.
