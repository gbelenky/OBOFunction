#requires -Version 7.0
<#
.SYNOPSIS
  Tests the SharePoint-profile identity chain the same way the real SPFx web part
  exercises it, WITHOUT needing SharePoint / Foundry in the loop.

.DESCRIPTION
  The production data path is:

      SPFx web part ──(user JWT, proxy audience)──▶ Proxy
                                                     │  OBO exchange
                                                     ▼
                                            MCP-audience user token
                                                     │  (Foundry passthrough delivers this)
                                                     ▼
                                            MCP server  ──OBO──▶ Microsoft Graph + SharePoint UPS

  This script reproduces every identity hop faithfully:
    1. Acquire a USER token for the PROXY audience          (exactly what SPFx AadHttpClient gets)
    2. Proxy performs On-Behalf-Of  ->  MCP-audience token  (exactly what Foundry passthrough will deliver)
    3. Call the MCP server's get_sharepoint_profile tool WITH that user token
    4. The MCP server runs its own OBO -> Graph /me + SharePoint UPS -> full per-user profile

  A successful run returns the signed-in user's real profile with "ResolvedVia": "user"
  and all custom UPS attributes (skills, interests, country, ...). This is the closest
  possible test to the SPFx scenario; the only piece NOT covered is the Foundry portal
  connection that delivers the user token to the MCP server (the documented passthrough blocker).

.PARAMETER Tenant          Entra tenant id.
.PARAMETER ProxyApp        Proxy API app registration (audience SPFx requests).
.PARAMETER McpApp          MCP server app registration (OBO target audience).
.PARAMETER McpUrl          MCP server /mcp endpoint.
.PARAMETER KeyVault        Key Vault holding the proxy client secret.
.PARAMETER SecretName      Secret name for the proxy client secret.

.EXAMPLE
  az login --tenant cbe03044-c23b-46df-93a5-c018d51915d8
  ./scripts/test-spfx-chain.ps1
#>
param(
  [string]$Tenant     = 'cbe03044-c23b-46df-93a5-c018d51915d8',
  [string]$ProxyApp   = '7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a',
  [string]$McpApp     = '4646a978-f7bc-46a5-9988-ff864cca0410',
  [string]$McpUrl     = 'https://app-mcp-z6vb2tjg2j4ye.azurewebsites.net/mcp',
  [string]$KeyVault   = 'kv-z6vb2tjg2j4ye',
  [string]$SecretName = 'AzureAd--ClientSecret'
)

$ErrorActionPreference = 'Stop'

Write-Host '=== STEP 1: user token for PROXY audience (what SPFx AadHttpClient gets) ===' -ForegroundColor Cyan
$userTok = az account get-access-token --scope "api://$ProxyApp/access_as_user" --query accessToken -o tsv
if ([string]::IsNullOrWhiteSpace($userTok)) { Write-Error "No user token. Run 'az login --tenant $Tenant' first."; return }
Write-Host "  user token acquired (len $($userTok.Length))" -ForegroundColor Green

Write-Host '=== STEP 2: proxy OBO exchange -> MCP-audience token (what Foundry passthrough delivers) ===' -ForegroundColor Cyan
$secret = az keyvault secret show --vault-name $KeyVault --name $SecretName --query value -o tsv
if ([string]::IsNullOrWhiteSpace($secret)) { Write-Error "Could not read proxy secret '$SecretName' from $KeyVault."; return }

$body = @{
  grant_type            = 'urn:ietf:params:oauth:grant-type:jwt-bearer'
  client_id             = $ProxyApp
  client_secret         = $secret
  assertion             = $userTok
  scope                 = "api://$McpApp/.default"
  requested_token_use   = 'on_behalf_of'
}
$obo = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" `
  -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -SkipHttpErrorCheck
if (-not $obo.access_token) { Write-Host 'OBO FAILED:' -ForegroundColor Red; $obo | ConvertTo-Json -Depth 5; return }
$mcpTok = $obo.access_token
Write-Host "  MCP-audience token acquired (len $($mcpTok.Length))" -ForegroundColor Green

Write-Host '=== STEP 3: call MCP server get_sharepoint_profile WITH the user token (simulates passthrough) ===' -ForegroundColor Cyan
$headers = @{
  'Content-Type'  = 'application/json'
  'Accept'        = 'application/json, text/event-stream'
  'Authorization' = "Bearer $mcpTok"
}
$callBody = '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"get_sharepoint_profile","arguments":{}}}'
$resp = Invoke-WebRequest -Uri $McpUrl -Method Post -Headers $headers -Body $callBody -SkipHttpErrorCheck
$raw = if ($resp.Content -is [byte[]]) { [System.Text.Encoding]::UTF8.GetString($resp.Content) } else { "$($resp.Content)" }

Write-Host "MCP HTTP status: $($resp.StatusCode)" -ForegroundColor Yellow

# SSE response: extract the JSON payload from the 'data:' line, then unwrap the tool text.
$dataLine = ($raw -split "`n" | Where-Object { $_ -like 'data:*' } | Select-Object -First 1)
if ($dataLine) {
  $json = ($dataLine -replace '^data:\s*', '') | ConvertFrom-Json
  $profileText = $json.result.content[0].text
  if ($json.result.isError -or $profileText -like '*An error occurred invoking*') {
    Write-Host 'TOOL RETURNED AN ERROR (app-only identity or OBO failure):' -ForegroundColor Red
    Write-Host $profileText
  } else {
    Write-Host '=== PROFILE (per-user, ResolvedVia should be "user") ===' -ForegroundColor Green
    $profileText
  }
} else {
  Write-Host 'Unexpected response:' -ForegroundColor Red
  $raw
}
