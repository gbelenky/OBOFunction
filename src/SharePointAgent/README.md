# SharePointAgent — Microsoft Foundry Hosted Agent (.NET 10)

A **hosted** Foundry agent built with the **Microsoft Agent Framework**. It exposes a single
chat agent whose **embedded `get_sharepoint_profile` tool** runs at the start of every
conversation to load the signed-in user's Microsoft Graph + SharePoint profile, then uses
that context (name, job title, department, office, skills, interests, responsibilities) in
its answers.

It runs **unchanged** both locally (for dev/debug) and as a deployed Foundry hosted agent,
using `Microsoft.Agents.AI.Foundry.Hosting` (`AddFoundryResponses` / `MapFoundryResponses`).

## How it works

```
chat ──► SharePointProfileAgent ──(session start)──► get_sharepoint_profile tool
                                                          │
                                  DefaultAzureCredential ──┼──► Microsoft Graph /me
                                                          └──► SharePoint UPS (PeopleManager/GetMyProperties)
                                                                   │
                                                          merged UserProfile (JSON) ──► model context
```

- **Hosting:** `WebApplication` + `AddFoundryResponses(agent)` + `MapFoundryResponses()`. Health probes at `/liveness` and `/readiness`.
- **Agent:** `AIProjectClient(endpoint, DefaultAzureCredential).AsAIAgent(model, name, instructions, tools)`.
- **Embedded tool:** `get_sharepoint_profile` (`AIFunctionFactory`). Agent instructions enforce calling it first each session.
- **Auth (local-first):** `DefaultAzureCredential` (your `az login` / VS Code identity) calls Graph and SharePoint. Locally this returns **your own** profile — correct for dev. In production the agent's managed identity is used. There is **no OBO** here: a hosted agent has no inbound user assertion to exchange.

## Project layout

```
src/SharePointAgent/
  Program.cs                       # host + agent wiring
  Models/UserProfile.cs            # merged Graph + SharePoint DTO
  Services/ISharePointProfileService.cs
  Services/SharePointProfileService.cs   # Graph /me + SharePoint UPS merge
  Tools/ProfileTools.cs            # get_sharepoint_profile AIFunction
  agent.yaml                       # hosted-agent manifest (runtime dotnet_10)
  .env.sample                      # copy to .env and fill in
  .agentignore
```

## Prerequisites

- **.NET 10 SDK**
- Azure CLI signed in: `az login` (so `DefaultAzureCredential` works)
- A Microsoft Foundry project with a deployed chat model (e.g. `gpt-4o-mini`)
- **Foundry Toolkit** VS Code extension (`ms-windows-ai-studio.windows-ai-studio`) and **Azure Developer CLI** (`azd`) for the run/invoke loop
- Recommended extensions are listed in `.vscode/extensions.json`

## Configure

```pwsh
cd src/SharePointAgent
Copy-Item .env.sample .env
# edit .env:
#   FOUNDRY_PROJECT_ENDPOINT=https://<project>.services.ai.azure.com/api/projects/<project>
#   AZURE_AI_MODEL_DEPLOYMENT_NAME=gpt-4o-mini
#   SHAREPOINT_TENANT_HOSTNAME=mngenv168112.sharepoint.com
```

| Variable | Purpose |
|---|---|
| `FOUNDRY_PROJECT_ENDPOINT` | Foundry project endpoint (model inference). Falls back to `AZURE_AI_PROJECT_ENDPOINT`. |
| `AZURE_AI_MODEL_DEPLOYMENT_NAME` | Chat model deployment name. |
| `SHAREPOINT_TENANT_HOSTNAME` | SharePoint host; used to build the UPS root site URL. |
| `SHAREPOINT_ROOT_SITE_URL` | Optional full root site URL override. |
| `PORT` / `DEFAULT_AD_PORT` | Local server port (default `8088`). |

## Run & debug locally

### Option A — Foundry Toolkit / azd (recommended)

```pwsh
cd src/SharePointAgent
azd ai agent run                       # starts the host on :8088 using .env + DefaultAzureCredential
# in another terminal:
azd ai agent invoke --local "Who am I and what are my skills?"
```

Open the **Agent Inspector** from the Foundry Toolkit (`Foundry: Open Agent (Test Tool)`) and point it at `http://localhost:8088` to chat interactively and inspect the `get_sharepoint_profile` tool call.

> VS Code tasks `azd ai agent run (agent)` and `azd ai agent invoke (agent)` wrap these commands.

### Option B — Plain debugging with breakpoints

1. Press **F5** → **"Debug SharePointAgent (host)"** (loads `.env`, launches on `:8088`, attaches the .NET debugger).
2. Set breakpoints in `Tools/ProfileTools.cs` or `Services/SharePointProfileService.cs`.
3. Drive it with `azd ai agent invoke --local "..."` or the Agent Inspector.

Health check while running: `curl http://localhost:8088/liveness` → `Healthy`.

## Deploy to Foundry

The agent is wired into the repo's `azure.yaml` as service **`sharepoint-agent`** (`host: azure.ai.agent`, runtime `dotnet_10`). After provisioning a Foundry project:

```pwsh
cd c:\src\OBOFunction
azd up                                 # provisions + deploys both the Function and the agent
# or just the agent:
azd deploy sharepoint-agent
```

`agent.yaml` declares the hosted manifest (responses protocol, `entry_point: SharePointAgent.dll`, env vars). In production the agent uses its **managed identity** — grant it Graph (`User.Read.All` / `User.Read`) and SharePoint read permissions for the profile fetch.

## Notes

- Preview packages are pinned to a known-good set (`Microsoft.Agents.AI.Foundry.Hosting 1.4.0-preview.260505.1`, `Azure.AI.Projects 2.1.0-beta.1`, `Microsoft.Extensions.AI 10.5.0`). Experimental-API warnings are suppressed via `<NoWarn>` in the csproj.
- SharePoint UPS fetch is best-effort locally — if the dev credential can't get a SharePoint token, the Graph profile is still returned and `sharePointProfile` is `null`.
