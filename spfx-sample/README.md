# SPFx sample

This sample demonstrates the **browser side** of the current integration:

```text
SPFx -> BFF proxy -> HostedSecureMcpAgent
```

The hosted-agent source is published at
[`gbelenky/HostedSecureMcpAgent`](https://github.com/gbelenky/HostedSecureMcpAgent).

The SPFx web part:

- calls the proxy with `AadHttpClient`
- never calls Foundry directly
- never performs OBO itself
- stays agnostic of hosted-agent internals

## Current live wiring

| Setting | Value |
| --- | --- |
| `PROXY_RESOURCE` | `api://7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a` |
| `PROXY_BASE` | `https://app-proxy-z6vb2tjg2j4ye.azurewebsites.net` |

Source file:

- `spfx-profile-agent/src/webparts/profileAgent/services/ProxyClient.ts`

## Why SPFx talks only to the proxy

The browser cannot replace the proxy because:

1. SPFx is a public client and cannot safely hold the OBO secret.
2. The proxy owns JWT validation and CORS policy.
3. The proxy hides Foundry endpoint and agent-target details from the browser.
4. Runtime fixes can be applied server-side without redeploying the client.

## Web part behavior

On load:

```json
{ "message": "", "greeting": true }
```

Chat turn:

```json
{ "message": "What is my profile?", "previousResponseId": "caresp_..." }
```

The web part stores the `responseId` it gets back and sends it as `previousResponseId` on the next
turn.

## Build

SPFx 1.23 requires Node 22.

```powershell
nvm install 22
nvm use 22
cd C:\src\OBOFunction\spfx-sample\spfx-profile-agent
npm install
npm run build
```

## Local workbench

```powershell
npm run start
```

## Deploy

1. Upload `sharepoint/solution/spfx-profile-agent.sppkg` to the tenant App Catalog.
2. Approve the proxy API permission request if needed.
3. Add the web part to a SharePoint page.

## Live verification target

The SharePoint page used for live verification in this environment:

- `https://mngenv168112.sharepoint.com/sites/gbelelenky-dev-sp/SitePages/CollabHome.aspx`

## Expected outcomes

- Greeting returns a short welcome
- “What is my profile?” returns the signed-in user profile
- Follow-up turns continue correctly through `previousResponseId`
