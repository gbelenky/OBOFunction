# Architecture — identity & OBO design

This document explains **how the signed-in user's identity flows end-to-end** in this repo, and the
security reasoning behind it. For setup/run steps see [`README.md`](./README.md).

> **Scope.** This repo demonstrates **how to obtain the signed-in user's profile (especially the
> `IntranetCountry` UPS field) via the OBO flow** — §1 – §3 below. The retrieved profile is injected
> into a **Microsoft Foundry hosted agent** as context ("Option A", §4), where it drives a
> **country-filtered FAQ / Q&A search** against Azure AI Search — see **§5**.

---

## 0. OBO in one picture — the hotel check-in analogy

Use this to explain the design to anyone non-technical. The whole flow is a hotel stay where the
guest checks in **once** and the front desk afterwards issues **single-purpose key cards in the
guest's name**.

**The one rule:** 🛂 **only the Passport Office issues cards.** The Front Desk can **never** make its
own — every key card is requested from, and printed by, the Passport Office. To get a card, the desk
must show **two** things:

1. 🎫 its own **staff badge** — proves *which desk* is asking → the **client secret**
2. 🪪 the **guest's current card** — proves *"this guest sent me"* → the **user's token (the OBO assertion)**

| Hotel | This system |
|---|---|
| 🧳 Guest | The signed-in user |
| 🛂 Passport Office (issues every card) | Microsoft Entra (issues every token) |
| 🏨 Front Desk | Proxy (`src/OBOFunction`) — the one confidential client |
| 🏊 Pool / 🍽️ Restaurant | SharePoint UPS / Microsoft Graph (the profile data) |
| 🤖 AI Lounge | Foundry hosted agent (`src/SharePointAgent`) |

```
🧳 Guest ──shows passport──▶ 🛂 Passport Office
                              🛂 issues 🟦 Card 1 (Front Desk), in the guest's name
   │  guest gives 🟦 Card 1 to the Front Desk
   ▼
🏨 Front Desk ──"this guest sent me" (🟦 Card 1 + 🎫 front-desk badge)──▶ 🛂 Passport Office
                              🛂 issues 🟪 Pool card + 🟧 Restaurant card + 🤖 AI-Lounge card,
                              all in the guest's name
   ▼
🏊 Pool   🍽️ Restaurant   🤖 AI Lounge   ← each opens only with its own card, always in the guest's name
```

| Step | Who asks 🛂 | What they present | 🛂 issues |
|---|---|---|---|
| 1 | 🧳 Guest | the **passport** (sign-in) | 🟦 card for Front Desk |
| 2 | 🏨 Front Desk | "**this guest sent me**" (🟦 card) + 🎫 badge | 🟪 Pool + 🟧 Restaurant + 🤖 AI-Lounge cards |

The Front Desk reads the guest's profile at the Pool and Restaurant **as the guest**, then escorts the
guest into the AI Lounge **as the guest** — carrying a one-page summary of who they are so the lounge
host can greet them by name. The lounge host never needs the guest's passport; the Front Desk already
did all the identity work.

**Two points to stress:**

- **The desk is a middleman, not a card-printer.** The Front Desk never creates access — it goes
  *back to the Passport Office* and says "this guest sent me." Technically: the proxy calls **Entra's
  token endpoint** to mint each downstream token; it can't forge one.
- **"This guest sent me" must be backed by the guest's real card.** The desk hands over the guest's
  **actual signed token** as proof. No guest card → the Passport Office refuses → no new card. That is
  what keeps every card in the **guest's** name instead of the hotel's.

> **Why this beats the "master key" (app-only) approach:** the alternative gives the hotel a master
> key that opens **every** guest's room with no guest present. OBO = **personal key cards** (per-user,
> least-privilege, auditable); app-only = **master key**.

**One-sentence version:** *The Front Desk makes no keys of its own — it goes back to the Passport
Office and says "this guest sent me," showing the guest's current card plus its own staff badge. The
Passport Office prints each next card in the guest's name. That's how the guest's identity reaches the
pool, the restaurant, and the AI lounge, while no single card opens the whole hotel.*

---

## 0a. Master key — no OBO (app-only)

This is the alternative approach to reach the same SharePoint data, told **as the same hotel** so the
trade-off is obvious. There is **no check-in** and **no OBO**: the hotel manager gives a service a
**master key**, signed off **once** by the building owner, that opens **every guest's room** — whether
or not the guest is present.

| Hotel | Master-key system |
|---|---|
| 🛂 Passport Office | Microsoft Entra |
| 🏛️ Building owner (signs off the master key once) | **Tenant admin** granting **admin consent** |
| 🗝️ Service holding the master key | An Entra app with an **application permission** (`SharePoint User.Read.All`) |
| 🚪 Every guest's room | **Every** user's SharePoint profile |
| 🧳 The guest | **Not involved** — never checks in, never consents |

**How it differs from the check-in (OBO) story:**

| | 🟦 Personal key cards (OBO / user tokens) | 🗝️ Master key (app-only + admin consent) |
|---|---|---|
| Who checks in? | The guest, every time | **Nobody** — approved once by the owner |
| Whose name is the key in? | The **guest's** | The **hotel's (the app's)** |
| Which rooms can it open? | Only **that guest's** room | **Every** guest's room |
| SharePoint call it uses | `GetMyProperties` ("my own profile, as me") | `GetPropertiesFor(accountName)` ("look up *anyone* by name") |
| Microsoft's name for it | OAuth **identity passthrough** | Microsoft Entra **app / agent identity** |
| Per-user consent & authorization | ✅ enforced by the token chain | ❌ bypassed entirely |

**Security & compliance risks (app-only):** over-broad access (one credential reads *every* profile —
large blast radius); no per-user authorization; weak auditability (reads attributed to the app, not the
user); standing privilege from a long-lived admin-consented secret; GDPR/data-minimization exposure;
consent bypass.

> **This repo implements §0 (personal key cards / OBO).** §0a is documented only to make the contrast
> explicit, because the two are easy to confuse and have very different security and compliance postures.

---

## 1. The components and the one identity chain

| Component | Role |
|---|---|
| **SPFx web part** | Auth boundary. Mints the user JWT (`AadHttpClient`, aud `api://<proxy-app>`). Holds no secrets. Agnostic — no profile or greeting logic. |
| **Proxy** (`src/OBOFunction`, App Service) | The only confidential client. Validates the JWT, OBO-exchanges it to **SharePoint + Graph** to build the profile, injects the profile as context, then OBO-exchanges to the **Foundry data plane** and calls the hosted agent **as the user**. |
| **Hosted agent** (`src/SharePointAgent`, Foundry) | Chat agent. Receives the profile as a host-injected `USER_PROFILE_JSON` context item and owns one **local `search_faq` tool** (Azure AI Search, agent identity). |

One request path, carrying the **same** signed-in user identity from the browser to SharePoint and to
the agent:

```
SPFx ──user JWT──► Proxy   (FIRST turn only)
   ├─ OBO: user JWT → SharePoint UPS token (https://<tenant>.sharepoint.com/.default)
   │        → GetMyProperties (+ custom IntranetCountry)              [as the user]
   ├─ OBO: user JWT → Graph token (https://graph.microsoft.com/.default)
   │        → /me                                                     [as the user]
   │        ⇒ build USER_PROFILE_JSON, inject as a developer-role item
   └─ OBO: user JWT → Foundry data-plane token (https://ai.azure.com/.default)
            → POST agent Responses (agent_reference)                  [as the user]
                 ▼
            Foundry Hosted Agent (SharePointProfileAgent)
                 └─ local search_faq ──► Azure AI Search faq-index    [agent identity, no OBO]
```

Follow-up turns inherit the conversation via `previous_response_id`, so the profile is resolved and
injected **once** (on the first turn).

**Why the proxy exists (three independent reasons):**
1. **SPFx can't be a confidential client.** It's a public, third-party SharePoint app; it cannot safely hold the OBO client secret.
2. **No CORS from the Foundry data plane** to a `*.sharepoint.com` origin — the browser can't call Foundry directly.
3. **Agent credentials must stay server-side.** A key or app token in the browser would leak.

The proxy is therefore the single place that holds the OBO secret. The user JWT and every OBO access
token are validated, used, and discarded — never logged, never stored.

## 2. One app registration (and why one is enough)

| App reg | Audience | OBO performed by it | Why it's needed |
|---|---|---|---|
| **Proxy API** | `api://<proxy-app>` | user → SharePoint, user → Graph, user → Foundry data plane | The token SPFx mints; the proxy is its confidential client and the only component performing OBO. |

Because the proxy does **all** the OBO itself (it both reads the profile and calls the agent), there is
no second downstream service that needs its own audience — so a single app registration suffices.

**Delegated permissions on the Proxy API:** Microsoft Graph `User.Read`, `openid`, `profile`,
`offline_access`; Office 365 SharePoint Online `AllSites.Read`, `User.Read.All`. Plus delegated
`user_impersonation` on the **Azure Machine Learning Services** API so the OBO to `https://ai.azure.com`
succeeds. All granted with **admin consent**.

**Pre-authorization:** the **SharePoint Online Client Extensibility** principal
`08e18876-6177-487e-b8b5-cf950c1e598c` is pre-authorized on the Proxy API for `access_as_user`, so SPFx
can mint the inbound token. The proxy's client secret lives in Key Vault, surfaced as config via the
managed identity.

## 3. Why custom SharePoint UPS fields require the user's identity

Custom User Profile Service attributes (Country / `IntranetCountry`, Responsibilities, Past projects,
Interests…) are readable only through a **SharePoint-consented token**:

- **User mode** (the proxy's OBO user token) → `GetMyProperties` returns the full UPS including custom
  `ExtendedProperties`. This is what the proxy uses on the first turn.
- **App-only mode** (a managed identity with no user token, or `DefaultAzureCredential` = Azure CLI app)
  lacks SharePoint delegated consent → `GetMyProperties` 401s → custom fields are empty. This is an
  **identity/consent** outcome, by design — not a code bug.

> **UI note:** custom UPS properties no longer render in the modern SPO profile/Delve UI (the classic
> profile page is retired). The values live in the UPS and are retrievable by API only — which is
> exactly what the proxy does.

---

## 4. Why "Option A" (proxy injects the profile) instead of OAuth passthrough

The natural-looking alternative is to let the agent fetch the profile itself via a tool, using Foundry's
**OAuth identity passthrough** so the agent calls SharePoint/Graph as the user.

**How passthrough actually works (per the Microsoft docs).** It is *not* tied to the Playground or any
"first-party" channel. On a user's first interaction, Agent Service **returns a consent link in the
agent's response**; the client surfaces that link, the user signs in and consents once, Agent Service
stores the per-user access/refresh tokens (scoped to user+agent), and subsequent tool calls run as that
user. **Any** client — Playground, Teams/M365 Copilot, *or a custom web app/SPFx* — can surface the
link, **provided**:

1. the agent is invoked **with the user's identity** (a user/OBO token) — **service-principal /
   app-only invocation is explicitly unsupported** for passthrough;
2. the user has at least the **Foundry User** role on the project, in the **same tenant** (no
   cross-tenant token exchange); and
3. the tool is configured as a **passthrough connection** (a portal MCP/A2A connection referenced by
   `project_connection_id`).

**Why it still didn't work in *our* chain — the real reason.** Not the channel, and not SPFx/SharePoint.
Our hosted agent is **code-first** and binds the MCP server **by URL** via `HostedMcpServerTool`, whose
converter can only emit `server_url` (+ optional `authorization`) — it **cannot emit a
`project_connection_id`**. With no passthrough *connection*, Agent Service never generated a consent
link (requirement #3 unmet), and our proxy→agent OBO additionally hit an MCP-audience mismatch. In other
words it was a **tool-authoring / connection-config limitation** of the URL-bound code agent, compounded
by an OBO-audience issue — *not* a platform rule that "custom web apps can't do passthrough."

A portal-configured passthrough connection on the agent (rather than a URL-bound code tool) could, in
principle, surface the consent link through this same SPFx → proxy → agent chain, because the proxy
already invokes the agent as the user (requirement #1) and the user has the project role (#2).

**Why we chose Option A anyway.** Even where passthrough is achievable it adds a per-user, interactive
consent ceremony and a portal-managed connection. For this profile use case the proxy *already* holds
the user's OBO token, so it simply resolves the profile itself (§1) and injects it as a
`USER_PROFILE_JSON` developer-role context item before invoking the agent. The agent treats it as
background knowledge — it never needs the user's token, no consent ceremony is required, and the user's
identity never leaks into agent traces.

> **References** —
> [Set up authentication for MCP tools](https://learn.microsoft.com/azure/ai-foundry/agents/how-to/tools/model-context-protocol) ·
> [Agent2Agent (A2A) authentication](https://learn.microsoft.com/azure/ai-foundry/agents/how-to/tools/a2a) ·
> [Agent identity](https://learn.microsoft.com/azure/ai-foundry/agents/concepts/agent-identity).
> The passthrough consent-link flow and its requirements (Foundry User role, same tenant, user-identity
> invocation, connection-based config) come from the first two; agent-identity (attended OBO vs
> unattended app-only) from the third.

---

## 5. Country-filtered FAQ / Q&A search (agent-owned, no OBO)

Once the profile is injected (§1), the user's **country** drives a search over a corporate FAQ knowledge
base. This is a **local, in-process tool** on the hosted agent (`src/SharePointAgent`): `search_faq`
(`Services/FaqSearchService.cs` + `Tools/FaqTools.cs`), registered via
`AIFunctionFactory.Create(faqTools.SearchFaqAsync, name: "search_faq")`.

**Why this tool does NOT use OBO** — and deliberately differs from the profile chain:

- It reads a **shared, non-personal** knowledge base (the FAQ index), not the user's personal data.
  There is no per-user authorization boundary to enforce, so no user token is required.
- It needs only **one value** from the profile — the country string — which the model passes as a tool
  argument. The tool itself never sees Graph, SharePoint, or any user token.
- It therefore runs with the **agent's own identity**: a read-only **query key** for local runs, or the
  agent's **Managed Identity** with the **`Search Index Data Reader`** role in production (keys omitted).

```
profile.country ──► model ──(local tool call: search_faq{country, question})──► Azure AI Search faq-index
                                                                                  filter: Location eq '<country>'
                                                                                       or Location eq 'Global'
```

**Country → index field mapping.** The index has no `Country` field; the filterable **`Location`** field
is used as the country/region (`AZURE_SEARCH_COUNTRY_FIELD=Location`). Valid values in the demo index are
**Europe, North America, Latin America, Global**. The search always tops results up with the full
country + `Global` set so universally-applicable and local-language entries (e.g. German *Urlaubsantrag*,
French *Demande de congé*) are always visible; the agent translates titles across languages and answers
in the user's language. Single quotes in the country value are escaped to prevent OData filter injection.

**Tool boundary.** The proxy is **tool-agnostic** — it never declares or executes the agent's tools. It
calls the hosted agent (via `agent_reference`), and the agent runs `search_faq` itself as part of the
turn. So the full SPFx → proxy → agent → `search_faq` path works end-to-end with the country coming from
the injected profile.
