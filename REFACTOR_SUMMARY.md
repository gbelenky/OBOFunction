# Architectural Refactor — Before/After

## The Refactor: From Profile Extraction Proxy → OBO Identity Pass-Through

This document shows what changed and why.

---

## ❌ OLD Architecture (Profile Extraction)

### What It Did
The proxy **fetched the user's profile** (SharePoint UPS + Graph) itself on the first turn, then **injected it into the agent as context**:

```
SPFx (Bearer token)
  ↓
Proxy validates JWT
  ├─ OBO to SharePoint UPS (GetMyProperties)
  ├─ OBO to Graph (/me)
  └─ Extracts profile (name, email, country)
  ↓
Injects profile into agent as "developer-role" context item
  ├─ Agent sees pre-resolved profile
  ├─ Agent doesn't fetch profile itself
  └─ search_faq tool uses injected profile
  ↓
Agent returns FAQ filtered by country
```

### Components
| Component | Role | Identity |
|---|---|---|
| **Proxy** | Fetches profile, injects context, calls agent | OBO (user's identity) |
| **Agent** | Receives profile, executes tools | Agent's own identity (Foundry managed) |
| **search_faq** | Filters FAQ by country (receives profile from proxy) | Agent's identity |
| **Toolbox connection** | None (not needed; proxy did all fetching) | N/A |

### Issues
1. **Profile resolution only on first turn** — subsequent turns didn't re-fetch (stale)
2. **Complex proxy logic** — multiple OBO exchanges, profile injection
3. **No user consent flow** — profile fetched silently
4. **Coupling** — proxy and agent tightly bound to profile schema

---

## ✅ NEW Architecture (OBO Identity Pass-Through)

### What It Does
The proxy **passes the user's identity through to the agent** via OBO; the **agent fetches its own profile** using a **Foundry Toolbox connection** on demand:

```
SPFx (Bearer token)
  ↓
Proxy validates JWT
  ├─ Extracts user's Bearer token (assertion)
  └─ Creates OBO credential (proxy app + user token)
  ↓
Proxy calls agent with OBO token (user's identity)
  ├─ Agent receives token AS THE USER
  └─ No profile injected; agent has user's identity
  ↓
Agent's first turn:
  ├─ Toolbox connection asks: "Get this user's profile"
  ├─ Toolbox mints THIS USER's Graph token (OBO within Foundry)
  ├─ Fetches profile (name, email, country) AS THE USER
  └─ (User may see sign-in prompt on first use of Toolbox)
  ↓
Agent calls search_faq tool
  ├─ Tool uses user's country from fetched profile
  └─ Returns country-filtered FAQ
  ↓
Agent returns answer (profile + FAQ result)
```

### Components
| Component | Role | Identity |
|---|---|---|
| **Proxy** | Validates JWT, creates OBO credential, calls agent | Confidential client (client secret holder) |
| **Agent** | Receives user's OBO token, fetches profile on demand, executes tools | User's identity (OBO pass-through) + Agent's identity (for tools) |
| **Toolbox connection** | Mints user's Graph token; resolves user's profile | User's delegated identity (via Toolbox) |
| **search_faq** | Filters FAQ by country (receives profile from agent's Toolbox) | Agent's own identity |

### Advantages
1. **User identity flows end-to-end** — agent executes AS THE USER
2. **Simpler proxy** — no profile extraction; pure identity pass-through
3. **User consent flow** — Toolbox connection may prompt on first use (standard OAuth)
4. **Flexible profile fetching** — agent can re-fetch if needed (though typically once per conversation)
5. **Least-privilege** — search_faq runs under agent's identity, not user's
6. **Multi-turn friendly** — profile context preserved in conversation (ProjectConversation)

---

## Code Changes

### Proxy: Before (Profile Extraction)

```csharp
// Old: Proxy fetches profile itself
var graphToken = await oboCredential.GetTokenAsync(
    new TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }), ct);

var graphClient = new GraphServiceClient(new BearerTokenAuthenticationProvider(...));
var user = await graphClient.Me.GetAsync();

var sharePointToken = await oboCredential.GetTokenAsync(
    new TokenRequestContext(new[] { "https://sharepoint.com/.default" }), ct);

var upsProfile = await GetUpsProfileAsync(sharePointToken);

// Inject into agent context
var contextItem = new UserProfileContextItem(user, upsProfile);
var agentRequest = new AgentChatRequest(message)
{
    ContextItems = new[] { contextItem }
};
```

### Proxy: After (OBO Pass-Through)

```csharp
// New: Proxy just creates OBO credential and calls agent
var oboCredential = new OnBehalfOfCredential(
    tenantId: _tenantId,
    clientId: _clientId,
    clientSecret: _clientSecret,
    userAssertion: userAssertion);

// Acquire Foundry-scoped token (will be used by agent)
var foundryToken = await oboCredential.GetTokenAsync(
    new TokenRequestContext(new[] { "https://ai.azure.com/.default" }), ct);

// Call agent with user's identity; no profile injection
var response = await CallAgentAsync(foundryToken, message);
```

### Configuration: Before vs After

| Setting | Old | New |
|---|---|---|
| `SharePoint:RootSiteUrl` | Required (for UPS fetch) | Not needed |
| `Foundry:ProjectEndpoint` | Required | **Required** |
| `Foundry:AgentName` | Required | **Required** |
| Proxy profile tools | `get_sharepoint_profile`, `get_user_profile` | **None** (agent owns profile) |
| Agent has Toolbox connection | No | **Yes** (SharePointProfileTools) |
| Agent receives context items | Yes (from proxy) | **No** (fetches on demand) |

---

## Deployment Artifacts

### What Changed in Git
- `src/OBOFunction/Services/AgentChatClient.cs` — Simplified to OBO pass-through
- `src/OBOFunction/Program.cs` — Removed profile extraction logic
- `.azure/OBOFunction/.env` — Agent name/endpoint added
- Documentation: `ARCHITECTURE.md` updated to reflect new design

### What Stayed the Same
- **Proxy endpoints** — `/api/agent/chat` format unchanged
- **Request/response DTOs** — `AgentChatRequest`, `AgentChatReply` unchanged
- **SPFx client** — No changes needed (calls proxy the same way)
- **Health probes** — `/liveness`, `/readiness` unchanged

---

## Testing Differences

### Old Flow Test
```
Proxy test:
✓ Verify profile extracted from SharePoint UPS
✓ Verify profile injected into agent context
✓ Verify agent receives full profile immediately

Agent test:
✓ Verify agent uses injected profile
✓ Verify no Toolbox connection called
```

### New Flow Test
```
Proxy test:
✓ Verify OBO token acquired
✓ Verify agent called with user's token
✓ Verify no profile extraction in proxy

Agent test:
✓ Verify Toolbox connection called on first turn
✓ Verify user's profile fetched via Toolbox
✓ Verify multi-turn context preserved
✓ Verify search_faq runs under agent's identity
```

---

## Migration Checklist

- [x] Refactored proxy to use `OnBehalfOfCredential`
- [x] Removed profile extraction from proxy
- [x] Deployed proxy to App Service
- [x] Verified proxy health endpoints
- [x] SPFx builds without changes
- [ ] Deploy SPFx to SharePoint
- [ ] Test greeting (Toolbox consent may appear on first use)
- [ ] Test profile query (should now come from agent's Toolbox)
- [ ] Test FAQ query (country filtering via agent)
- [ ] Verify App Insights traces show new flow

---

## Why This Refactor?

**Security:**
- ✅ User's identity flows end-to-end (least surprise)
- ✅ Proxy never touches user's profile data (separation of concerns)
- ✅ search_faq runs under agent's identity (least-privilege for data access)

**Flexibility:**
- ✅ Agent can re-fetch profile if needed (supports future use cases)
- ✅ Toolbox handles all auth flows (standard OAuth with consent)
- ✅ Easy to add more Toolbox connections (no proxy changes needed)

**Simplicity:**
- ✅ Proxy is now a pure identity pass-through (stateless, scalable)
- ✅ No complex profile injection logic in proxy
- ✅ Easier to test (fewer OBO exchanges in proxy)

**Standards:**
- ✅ Follows Microsoft Foundry's recommended pattern (Toolbox for data)
- ✅ Uses Azure Identity SDK correctly (OBO via `OnBehalfOfCredential`)
- ✅ Aligns with OAuth 2.0 spec (identity pass-through, not token laundering)
