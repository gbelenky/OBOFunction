# Project: OBOFunction

## Current architecture

This repo is the **BFF proxy** for a SharePoint SPFx client.

```text
SPFx -> src/OBOFunction -> HostedSecureMcpAgent (Foundry, external to this repo)
```

## Source of truth

- Setup and repo scope: `../README.md`
- Architecture and BFF justification: `../ARCHITECTURE.md`

## What this repo contains

1. **Proxy** — `src/OBOFunction`
   - ASP.NET Core (.NET 8)
   - validates the inbound SPFx JWT
   - extracts the user bearer token
   - performs OBO to `https://ai.azure.com/.default`
   - calls the configured Foundry hosted agent
   - exposes:
     - `POST /api/agent/chat`
     - `GET /liveness`
     - `GET /readiness`

2. **SPFx sample** — `spfx-sample/spfx-profile-agent`
   - reference web part using `AadHttpClient`
   - talks only to the proxy

## What this repo does not contain

- no local hosted-agent service
- no local MCP/toolbox/tool implementation for the hosted agent
- no profile-fetching service in the proxy
- no Azure Functions runtime model

## Current Foundry target

The proxy must target:

- `HostedSecureMcpAgent`

Current config shape:

- `Foundry:ProjectEndpoint`
- `Foundry:AgentName`
- `Foundry:AgentResponsesUrl`
- `Foundry:ApiVersion`
- `Foundry:TokenScope`

## BFF rule

Do not redesign this into browser -> Foundry direct calls.

The proxy is required because:

1. SPFx is a public client and cannot safely hold the OBO secret.
2. The proxy is the JWT-validation and CORS boundary.
3. The proxy owns Foundry endpoint/agent routing and retry behavior.
4. The browser must remain agnostic of hosted-agent internals.

## Implementation rules

- Keep the proxy tool-agnostic: it calls the hosted agent only.
- Do not reintroduce local hosted-agent code into this repo unless explicitly requested.
- Do not reintroduce Key Vault runtime dependency for the proxy secret unless explicitly requested.
- Do not add browser-direct Foundry calls.
- Keep first-turn requests free of `previous_response_id` when there is no prior response id.
- Never log user JWTs, OBO tokens, or raw secrets.
- SPFx sample UI must use only public, supported libraries. Use Fluent UI v9 (`@fluentui/react-components`); never use internal-only `@fluentui-copilot/*` packages.

## Git workflow

- **Do NOT commit or push. The user commits and pushes themselves.** Only run `git commit` / `git push` when the user explicitly asks. Making changes is fine; leave them unstaged/uncommitted otherwise.

## Definition of done for changes

- Build succeeds
- Tests succeed
- Docs match the current live setup
- `azure.yaml` and infra match the actual deployed architecture
