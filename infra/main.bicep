targetScope = 'subscription'

@minLength(1)
@maxLength(64)
@description('Name of the azd environment, used in resource naming.')
param environmentName string

@minLength(1)
@description('Azure region for all resources.')
param location string

@description('Entra ID tenant for the Function App registration.')
param aadTenantId string

@description('Function App registration (client) ID.')
param aadClientId string

@secure()
@description('Proxy API app-registration client secret. Optional: when empty, the secret is not seeded and OBO is disabled until set. Never stored in app settings directly.')
param aadClientSecret string = ''

@description('SharePoint tenant hostname for CORS, e.g. contoso.sharepoint.com.')
param sharepointTenantHostname string

@description('Foundry project endpoint URL.')
param foundryProjectEndpoint string = ''

@description('Foundry agent ID (asst_xxx).')
param foundryAgentId string = ''

@description('Principal ID of the user/SP running azd, for KV access during seeding.')
param principalId string = ''

var resourceToken = toLower(uniqueString(subscription().id, environmentName, location))
var tags = {
  'azd-env-name': environmentName
  purpose: 'demo'
  owner: 'gbelenky'
  project: 'obo-function'
}

resource rg 'Microsoft.Resources/resourceGroups@2024-03-01' = {
  name: 'rg-${environmentName}'
  location: location
  tags: tags
}

module resources 'modules/resources.bicep' = {
  name: 'resources'
  scope: rg
  params: {
    location: location
    resourceToken: resourceToken
    tags: tags
    aadTenantId: aadTenantId
    aadClientId: aadClientId
    aadClientSecret: aadClientSecret
    sharepointTenantHostname: sharepointTenantHostname
    foundryProjectEndpoint: foundryProjectEndpoint
    foundryAgentId: foundryAgentId
    principalId: principalId
  }
}

output AZURE_LOCATION string = location
output AZURE_TENANT_ID string = aadTenantId
output AZURE_RESOURCE_GROUP string = rg.name
output SERVICE_API_NAME string = resources.outputs.proxyAppName
output SERVICE_API_HOSTNAME string = resources.outputs.proxyAppHostname
output SERVICE_SHAREPOINT_MCP_NAME string = resources.outputs.mcpAppName
output SERVICE_SHAREPOINT_MCP_HOSTNAME string = resources.outputs.mcpAppHostname
output MCP_SERVER_URL string = '${resources.outputs.mcpAppHostname}/mcp'
output AZURE_KEY_VAULT_NAME string = resources.outputs.keyVaultName
output AZURE_KEY_VAULT_ENDPOINT string = resources.outputs.keyVaultEndpoint
output APPLICATIONINSIGHTS_CONNECTION_STRING string = resources.outputs.appInsightsConnectionString
output API_APP_RESOURCE_ID string = 'api://${aadClientId}'
