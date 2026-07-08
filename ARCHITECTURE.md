# Architecture

This document captures the **current** design only.

## 1. Live topology

```text
SharePoint page
  -> SPFx web part
  -> BFF proxy (src/OBOFunction)
  -> Foundry Responses endpoint
  -> HostedSecureMcpAgent
  -> agent-managed downstream tools/connections
```

This repo owns the **proxy** and the **SPFx sample**. It does **not** own the live hosted agent
implementation anymore.

## 2. Identity flow

1. The SPFx web part uses `AadHttpClient` to call `api://<proxy-app>`.
2. The proxy validates the inbound JWT.
3. The proxy extracts the incoming bearer token from the `Authorization` header.
4. The proxy uses `OnBehalfOfCredential` with:
   - `AzureAd__TenantId`
   - `AzureAd__ClientId`
   - `AzureAd__ClientSecret`
   - the inbound user token as the user assertion
5. The proxy acquires a user-scoped token for:
   - `https://ai.azure.com/.default`
6. The proxy calls the configured Foundry Responses endpoint:
   - currently `HostedSecureMcpAgent`
7. The agent continues the workflow using its own configuration and tools.

The proxy does **not** fetch the SharePoint or Graph profile itself anymore.

## 3. Why the BFF proxy exists

The proxy is mandatory for four concrete reasons.

### Confidential-client requirement

OBO requires a confidential client. SPFx runs in the browser and cannot safely hold the
app-registration secret. The proxy is therefore the only valid OBO broker.

### Security boundary

The proxy keeps these server-side:

- the OBO client secret
- the Foundry endpoint selection
- any future routing/policy logic
- distributed tracing and server-side diagnostics

### Stable contract for the client

The SharePoint client gets one stable integration surface:

- `POST /api/agent/chat`

That means the browser does not need to understand:

- Foundry API versions
- agent names
- conversation-cursor quirks
- retry semantics
- downstream agent changes

### Operational control

The proxy gives the team one place to:

- validate JWT audience/issuer
- constrain CORS to the SharePoint tenant
- inspect failures in App Insights
- add safe retries or payload normalization

The recent live fix is a concrete example: the proxy removed null `previous_response_id` from
first-turn requests, fixing a real 400 from Foundry without any SPFx change.

## 4. Proof from the live setup

The current live configuration points the proxy at:

- `Foundry__AgentName = HostedSecureMcpAgent`
- `Foundry__AgentResponsesUrl = .../agents/HostedSecureMcpAgent/endpoint/protocols/openai/responses?api-version=v1`

The SharePoint page was verified live through the browser path:

1. SPFx loaded on the SharePoint page
2. SPFx called `https://app-proxy-z6vb2tjg2j4ye.azurewebsites.net/api/agent/chat`
3. The proxy returned `200`
4. The user saw:
   - a greeting
   - the user profile from the agent path

That proves the BFF proxy is both:

- **architecturally necessary**, and
- **the actual live integration point**

## 5. Proxy responsibilities

The proxy owns:

- JWT validation
- OBO to Foundry
- request normalization
- multi-turn cursor forwarding (`previousResponseId`)
- error shaping (`ProblemDetails`)
- telemetry correlation

The proxy does **not** own:

- profile fetching
- local MCP server execution
- local FAQ tool execution
- hosted-agent implementation details

## 6. Current repository scope

### In scope

- `src/OBOFunction`
- `infra/`
- `spfx-sample/`
- proxy tests

### Out of scope

- local Foundry hosted-agent code
- local tool implementations for the hosted agent
- historical rollout and migration notes for superseded architectures

Those stale artifacts were removed so the repo now reflects the live setup.

## 7. Observability

The proxy emits OpenTelemetry spans through Azure Monitor / Application Insights.

Important tags include:

- `gen_ai.agent.name`
- `gen_ai.response.id`
- `gen_ai.previous_response.id`
- `gen_ai.response.status`
- `enduser.id`

This keeps the proxy side diagnosable while avoiding token or profile leakage in logs.

## 8. Deployment model

`azure.yaml` now deploys only:

- `api`

The hosted agent is expected to exist already in Foundry and be referenced by the proxy through
configuration.
