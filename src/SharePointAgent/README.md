# SharePointAgent — Microsoft Foundry Hosted Agent (.NET 10)

A **hosted** Foundry agent (Microsoft Agent Framework) with **one local tool** and a **host-injected
profile**:

- **`search_faq`** — a **local, in-process** Azure AI Search tool (no MCP, no OBO). It queries the
  corporate FAQ index filtered by the user's **country** (the index `Location` field) and runs with the
  **agent's own identity**. Registered only when `AZURE_SEARCH_ENDPOINT` is set.
- **Profile** — the signed-in user's profile is **not** fetched by the agent. The proxy
  (`src/OBOFunction`, `POST /api/agent/chat`) resolves it via OBO and injects it as a
  `USER_PROFILE_JSON` developer/context message on the first turn ("Option A"). The agent treats it as
  background knowledge so it can greet by first name and filter FAQs by country. See
  [`../../ARCHITECTURE.md`](../../ARCHITECTURE.md).

## How it works

```
proxy POST /api/agent/chat ──(agent_reference, as the user)──► SharePointProfileAgent
   (first turn injects USER_PROFILE_JSON context)                  │
                                                                   └──(local in-process tool)──► search_faq
                                                                        Azure AI Search faq-index
                                                                        (agent identity; filter Location = profile country + Global)
```

- **Hosting:** `WebApplication` + `AddFoundryResponses(agent)` + `MapFoundryResponses()`; health at `/liveness`, `/readiness`. Runs unchanged locally and as a Foundry hosted agent.
- **Agent:** `AIProjectClient(endpoint, DefaultAzureCredential).AsAIAgent(model, name, instructions, tools)`.
- **FAQ tool:** `AIFunctionFactory.Create(faqTools.SearchFaqAsync, name: "search_faq")` over `FaqSearchService` (`Services/FaqSearchService.cs`). Builds the OData filter `Location eq '<country>' or Location eq 'Global'`, runs full-text search and returns compact camelCase JSON. The model is instructed to pass the `country` from the injected profile; globally-applicable entries are always included so results are never empty for a country with no region-specific FAQs.
- **Identity:** the agent never receives the user's token. It uses its own identity only for `search_faq` — a read-only **query key** for local runs, or its **Managed Identity** with the **`Search Index Data Reader`** role in production (key omitted). The personal profile data arrives pre-resolved from the proxy as context.

## Project layout

```
src/SharePointAgent/
  Program.cs                    # host + agent wiring + local search_faq tool registration
  Services/FaqSearchService.cs  # Azure AI Search query (filter Location = country + Global)
  Tools/FaqTools.cs             # [Description]-annotated search_faq surface for AIFunctionFactory
  agent.yaml                    # hosted-agent manifest (runtime dotnet_10)
  Properties/launchSettings.json
  SharePointAgent.csproj
  .env.sample                   # copy to .env and fill in
  .agentignore
```

`search_faq` is the only tool and only reads the search index — it never touches Graph/SharePoint or a
user token.

## Prerequisites

- **.NET 10 SDK**; `az login` (so `DefaultAzureCredential` works).
- A Microsoft Foundry project with a deployed chat model (`gpt-4.1-mini`).
- **Foundry Toolkit** VS Code extension and **azd** for the run/invoke loop (see `.vscode/extensions.json`).

## Configure

```pwsh
cd src/SharePointAgent
Copy-Item .env.sample .env
# edit .env
```

| Variable | Purpose |
|---|---|
| `FOUNDRY_PROJECT_ENDPOINT` | Foundry project endpoint (falls back to `AZURE_AI_PROJECT_ENDPOINT`). |
| `AZURE_AI_MODEL_DEPLOYMENT_NAME` | Chat model deployment (default `gpt-4.1-mini`). |
| `PORT` / `DEFAULT_AD_PORT` | Local port (default `8088`). |
| `AZURE_SEARCH_ENDPOINT` | Azure AI Search service URL hosting the FAQ index. **Omit to disable the `search_faq` tool.** |
| `AZURE_SEARCH_INDEX_NAME` | FAQ index name (default `faq-index`). |
| `AZURE_SEARCH_COUNTRY_FIELD` | Filterable index field used as the country/region filter (default `Location`). |
| `AZURE_SEARCH_API_KEY` | Optional read-only **query key** for local runs. Leave blank in prod and grant the agent's Managed Identity **`Search Index Data Reader`**. |
| `AZURE_SEARCH_INCLUDE_GLOBAL` | Also return `Location='Global'` entries (default `true`). |

## Run & debug locally

```pwsh
cd src/SharePointAgent
azd ai agent run                       # host on :8088 using .env + DefaultAzureCredential
# in another terminal:
azd ai agent invoke --local "How can I request a vacation?"
```

Open the **Agent Inspector** (Foundry Toolkit → `Foundry: Open Agent (Test Tool)`) against
`http://localhost:8088` to chat and inspect the `search_faq` tool call. Or press **F5**
→ *Debug SharePointAgent (host)* and set breakpoints. Health: `curl http://localhost:8088/liveness`.

> Locally there is no proxy to inject the profile, so the agent has no country. Either state your country
> in the chat (e.g. "I'm based in Europe") so `search_faq` filters correctly, or it returns the
> globally-applicable entries.

## Deploy to Foundry

Wired into `azure.yaml` as service **`sharepoint-agent`** (`host: azure.ai.agent`, runtime `dotnet_10`).

```pwsh
cd c:\src\OBOFunction
azd up                  # provisions + deploys both services
# or just the agent:
azd deploy sharepoint-agent
```

`agent.yaml` declares the hosted manifest (responses protocol, `entry_point: SharePointAgent.dll`, the
search env vars). No MCP, toolbox, or OAuth-passthrough connection is required — the profile is injected
by the proxy and `search_faq` runs under the agent's own identity.

## Notes

- Preview packages are pinned to a known-good set; experimental-API warnings are suppressed via `<NoWarn>` in the csproj.
- The agent depends only on `Microsoft.Agents.AI.Foundry.Hosting`, `Azure.AI.Projects`, `Microsoft.Extensions.AI`, and `Azure.Search.Documents` — no MCP packages.
