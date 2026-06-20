# Architecture — identity & OBO design

This document explains **how the signed-in user's identity flows end-to-end** in this repo, and the
security reasoning behind it. For setup/run steps see [`README.md`](./README.md).

> **Scope.** What this repo *implements* is **§1 – §3** below: the SharePoint-profile-for-a-Foundry-agent
> OBO chain. **§4 (Broader vision)** sketches a larger "secure intranet AI assistant" (content ingestion,
> Azure AI Search, per-document ACL trimming) that this profile service would slot into — it is **context,
> not code in this repo**. **§5** is a reference matrix of which Foundry agent auth paths support OBO.

---

## 0. OBO in one picture — the hotel check-in analogy

Use this to explain the design to anyone non-technical. The whole flow is a hotel stay where the
guest checks in **once** and every desk afterwards issues a **single-purpose key card in the guest's name**.

**The one rule:** 🛂 **only the Passport Office issues cards.** The Front Desk and Concierge can **never**
make their own — every key card is requested from, and printed by, the Passport Office. To get a card, a
desk must show **two** things:

1. 🎫 its own **staff badge** — proves *which desk* is asking → the **client secret**
2. 🪪 the **guest's current card** — proves *"this guest sent me"* → the **user's token (the OBO assertion)**

| Hotel | This system |
|---|---|
| 🧳 Guest | The signed-in user |
| 🛂 Passport Office (issues every card) | Microsoft Entra (issues every token) |
| 🏨 Front Desk | Proxy (`src/OBOFunction`) — confidential client |
| 🛎️ Concierge | MCP server (`src/SharePointMcp`) — confidential client |
| 🏊 Pool / 🍽️ Restaurant | Microsoft Graph / SharePoint (the real data) |

```
🧳 Guest ──shows passport──▶ 🛂 Passport Office
                              🛂 issues 🟦 Card 1 (Front Desk), in the guest's name
   │  guest gives 🟦 Card 1 to the Front Desk
   ▼
🏨 Front Desk ──"this guest sent me" (🟦 Card 1 + 🎫 front-desk badge)──▶ 🛂 Passport Office
                              🛂 issues 🟩 Card 2 (Concierge), in the guest's name
   │  Front Desk passes 🟩 Card 2 to the Concierge
   ▼
🛎️ Concierge ──"this guest sent me" (🟩 Card 2 + 🎫 concierge badge)──▶ 🛂 Passport Office
                              🛂 issues 🟪 Card 3 (Pool) + 🟧 Card 4 (Restaurant)
   ▼
🏊 Pool        🍽️ Restaurant   ← each opens only with its own card, always in the guest's name
```

| Step | Who asks 🛂 | What they present | 🛂 issues |
|---|---|---|---|
| 1 | 🧳 Guest | the **passport** (sign-in) | 🟦 card for Front Desk |
| 2 | 🏨 Front Desk | "**this guest sent me**" (🟦 card) + 🎫 badge | 🟩 card for Concierge |
| 3 | 🛎️ Concierge | "**this guest sent me**" (🟩 card) + 🎫 badge | 🟪 Pool + 🟧 Restaurant cards |

**Two points to stress:**

- **The desks are middlemen, not card-printers.** Front Desk and Concierge never create access — each goes
  *back to the Passport Office* and says "this guest sent me." Technically: every service calls **Entra's
  token endpoint** to mint the next token; it can't forge one.
- **"This guest sent me" must be backed by the guest's real card.** A desk can't merely *claim* a guest sent
  it — it hands over the guest's **actual signed token** as proof. No guest card → the Passport Office refuses
  → no new card. That is what keeps every card in the **guest's** name instead of the hotel's.

> **Why this beats the "master key" (app-only) approach:** the alternative is giving the Concierge a master
> key that opens **every** guest's room with no guest present. Fewer steps, but the app can read **anyone's**
> data, anytime. OBO = **personal key cards** (per-user, least-privilege, auditable); app-only = **master key**.

**One-sentence version:** *No desk makes its own keys — each goes back to the Passport Office and says "this
guest sent me," showing the guest's current card plus its own staff badge. The Passport Office prints the
next door's card, always in the guest's name. That's how the guest's identity reaches the pool and
restaurant, while no single card opens the whole hotel.*

The sections below map this analogy to the real components, tokens, and app registrations.

---

## 0a. Master key — no OBO (app-only)

This is the alternative approach to reach the same SharePoint data. It is worth telling **as the same
hotel**, so the trade-off is obvious.

In this version there is **no check-in** and **no OBO**. Instead, the hotel manager gives the Concierge a
**master key**, signed off **once** by the building owner, that opens **every guest's room** — whether or
not the guest is present, and without the guest ever knowing.

| Hotel | Master-key system |
|---|---|
| 🛂 Passport Office | Microsoft Entra |
| 🏛️ Building owner (signs off the master key once) | **Tenant admin** granting **admin consent** |
| 🛎️ Concierge holding the master key | An Entra app with an **application permission** (`SharePoint User.Read.All`) |
| 🚪 Every guest's room | **Every** user's SharePoint profile |
| 🧳 The guest | **Not involved** — never checks in, never consents |

```
🏛️ Building owner ──"approve this once"──▶ 🛂 Passport Office
                                            🛂 issues 🗝️ a MASTER KEY to the Concierge (in the HOTEL's name)
   ▼
🛎️ Concierge ──🗝️ master key──▶ 🚪 any guest's room (no guest present, no per-guest check)
```

**How it differs from the check-in (OBO) story — same hotel, three changes:**

| | 🟦 Personal key cards (OBO / user tokens) | 🗝️ Master key (app-only + admin consent) |
|---|---|---|
| Who checks in? | The guest, every time | **Nobody** — approved once by the owner |
| Whose name is the key in? | The **guest's** | The **hotel's (the app's)** |
| Which rooms can it open? | Only **that guest's** room | **Every** guest's room |
| "This guest sent me"? | Required at every desk | **Never said** — no guest in the loop |
| SharePoint call it uses | `GetMyProperties` ("my own profile, as me") | `GetPropertiesFor(accountName)` ("look up *anyone* by name") |
| Microsoft's name for it | OAuth **identity passthrough** | Microsoft Entra **app / agent identity** |
| Per-user consent & authorization | ✅ enforced by the token chain | ❌ bypassed entirely |

**When the master key is a reasonable choice:** background jobs, batch enrichment, or system-to-system
reads where **there is no signed-in user** and an admin has deliberately accepted that the app can read
everyone. It is simpler (one consent, no check-in) and sometimes unavoidable.

**The trade-off:** the master key removes the guest from the loop. The app
can read **any** user's profile at **any** time, and audit logs show **the hotel**, not the guest, opening
the door. If the requirement is "each user sees only their own data, and the trace proves it," the master
key cannot provide it — that is precisely what the personal-key-card (OBO) model in §0 is for.

**Security & compliance risks (app-only):**
- **Over-broad access / no least privilege.** One credential reads *every* user's profile; a leaked secret exposes all of them at once (large blast radius).
- **No per-user authorization.** The app isn't constrained to the data the signed-in user may see — there is no user-scoped boundary to enforce.
- **Weak auditability.** Traces attribute every read to the app identity, not the end user, so "who accessed whose data" cannot be reconstructed.
- **Standing privilege & secret risk.** A long-lived, admin-consented app secret is a persistent high-value target; rotation and storage become critical.
- **Data-protection exposure.** Bulk readability of personal profile data raises GDPR/data-minimization and purpose-limitation concerns (processing beyond the acting user's need-to-know).
- **Consent bypass.** Per-user consent is skipped entirely, so users have no visibility into or control over access to their data.

> **This repo implements §0 (personal key cards / OBO).** §0a is documented only to make the contrast
> explicit, because the two are easy to confuse and have very different security and compliance postures.

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
