#requires -Version 7.0
# Refreshes OBO_TOKEN env var for REST Client.
# Usage:  . ./scripts/refresh-token.ps1
# (note the leading dot — dot-source so the env var sticks in your shell)

$apiClientId = '7ce28b8f-cb0e-4a07-8cfb-dfe8f36d644a'
$resource = "api://$apiClientId"

$token = az account get-access-token --resource $resource --query accessToken -o tsv
if ([string]::IsNullOrWhiteSpace($token)) {
    Write-Error "Failed to get token. Did you run 'az login --tenant <tenant-id>' first?"
    return
}

$env:OBO_TOKEN = $token
Write-Host "OBO_TOKEN refreshed for resource $resource (valid ~60 min)" -ForegroundColor Green
Write-Host "Now restart VS Code (or reload window) so REST Client sees the env var." -ForegroundColor Yellow
