# SharePointAgent — Microsoft Foundry Hosted Agent (.NET 10)

A **hosted** Foundry agent (Microsoft Agent Framework) with two tools:

1. **`get_sharepoint_profile`** — the standalone **SharePointMcp** server, declared as a Foundry-native
   MCP tool (`HostedMcpServerTool`). The agent holds **no embedded profile logic** — it asks the MCP
   server, which does the OBO to Graph + SharePoint.
2. **`search_faq`** — a **local, in-process** Azure AI Search tool (no MCP, no OBO). It queries the
   corporate FAQ index filtered by the user's **country** (the index `Location` field) and runs with the
   **agent's own identity**. Registered only when `AZURE_SEARCH_ENDPOINT` is set.

This is the **Playground / autonomous demo** surface. The production SPFx path does **not** route auth
through this agent — the proxy (`src/OBOFunction`, `POST /api/agent/chat`) calls the model Responses
endpoint with a per-user OBO token directly. See [`../../ARCHITECTURE.md`](../../ARCHITECTURE.md).

## How it works

```
chat ──► SharePointProfileAgent ──(Foundry-native mcp tool)──► SharePointMcp /mcp
            │                                                      │  OBO (if a user token is present)
            │                                                      ▼
            │                                               Graph /me + SharePoint UPS
            │
            └──(local in-process tool)──► search_faq ──► Azure AI Search faq-index
                                            (agent identity; filter Location = profile country + Global)
```

- **Hosting:** `WebApplication` + `AddFoundryResponses(agent)` + `MapFoundryResponses()`; health at `/liveness`, `/readiness`. Runs unchanged locally and as a Foundry hosted agent.
- **Agent:** `AIProjectClient(endpoint, DefaultAzureCredential).AsAIAgent(model, name, instructions, tools)`.
- **Profile tool:** `new HostedMcpServerTool("SharePointMcp", MCP_SERVER_URL)` with `AllowedTools = ["get_sharepoint_profile"]` and `ApprovalMode = NeverRequire`.
- **FAQ tool:** `AIFunctionFactory.Create(faqTools.SearchFaqAsync, name: "search_faq")` over `FaqSearchService` (`Services/FaqSearchService.cs`). Builds the OData filter `Location eq '<country>' or Location eq 'Global'`, runs full-text search and returns compact camelCase JSON. The model is instructed to pass the `country` from the loaded profile; globally-applicable entries are always included so results are never empty for a country with no region-specific FAQs.
- **Identity into the MCP server:**
  - **Playground (autonomous):** no user token → the MCP server falls back to its managed identity (app-only) → per-user fields empty. Requires a Foundry **OAuth identity-passthrough** connection to supply a real user token.
  - **Local dev:** set `MCP_USER_AUTHORIZATION` to a static dev bearer token and it is forwarded on the tool's `Authorization` header.

> **Why no `project_connection_id`:** the public `HostedMcpServerTool` → OpenAI `McpTool` converter only
> emits `server_url` (+ optional `authorization`) or a built-in `connector_id`; it cannot emit a Foundry
> `project_connection_id`. So a code agent binds the MCP server by **URL**, and per-user identity arrives
> via the `authorization` header (proxy path) or a portal-configured OAuth passthrough connection (Playground).

## Project layout

```
src/SharePointAgent/
  Program.cs                 # host + agent wiring + HostedMcpServerTool + local search_faq tool
  Services/FaqSearchService.cs  # Azure AI Search query (filter Location = country + Global)
  Tools/FaqTools.cs          # [Description]-annotated search_faq surface for AIFunctionFactory
  agent.yaml                 # hosted-agent manifest (runtime dotnet_10)
  Properties/launchSettings.json
  SharePointAgent.csproj
  .env.sample                # copy to .env and fill in
  .agentignore
```

Profile logic lives in `src/SharePointMcp` (called via the MCP tool); the FAQ tool is the only local
tool and only reads the search index — it never touches Graph/SharePoint or a user token.

## Prerequisites

- **.NET 10 SDK**; `az login` (so `DefaultAzureCredential` works).
- A Microsoft Foundry project with a deployed chat model (`gpt-4.1-mini`).
- A running **SharePointMcp** server URL for `MCP_SERVER_URL` (local `http://localhost:8089/mcp` or the deployed `https://app-mcp-<token>.azurewebsites.net/mcp`).
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
| `MCP_SERVER_URL` | The SharePointMcp `/mcp` endpoint. |
| `MCP_USER_AUTHORIZATION` | Optional static dev bearer token forwarded to the MCP tool (local only). |
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
azd ai agent invoke --local "Who am I and what are my skills?"
```

Open the **Agent Inspector** (Foundry Toolkit → `Foundry: Open Agent (Test Tool)`) against
`http://localhost:8088` to chat and inspect the `get_sharepoint_profile` tool call. Or press **F5**
→ *Debug SharePointAgent (host)* and set breakpoints. Health: `curl http://localhost:8088/liveness`.

## Deploy to Foundry

Wired into `azure.yaml` as service **`sharepoint-agent`** (`host: azure.ai.agent`, runtime `dotnet_10`).

```pwsh
cd c:\src\OBOFunction
azd up                  # provisions + deploys all three services
# or just the agent:
azd deploy sharepoint-agent
```

`agent.yaml` declares the hosted manifest (responses protocol, `entry_point: SharePointAgent.dll`, env vars
including `MCP_SERVER_URL`). For the Playground to return per-user data, configure the project's MCP
connection as **OAuth identity passthrough** (see ARCHITECTURE.md §5.1).

## Notes

- Preview packages are pinned to a known-good set; experimental-API warnings are suppressed via `<NoWarn>` in the csproj.
- `ModelContextProtocol` is pinned to **1.2.0** (newer majors force `Microsoft.Extensions.AI.Abstractions` 10.5.2, which breaks `AgentServer.Responses` at runtime).
