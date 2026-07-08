targetScope = 'subscription'

@minLength(1)
@maxLength(64)
@description('Name of the azd environment, used in resource naming.')
param environmentName string

@minLength(1)
@description('Azure region for all resources.')
param location string

@description('Entra ID tenant for the proxy app registration.')
param aadTenantId string

@description('Proxy app registration (client) ID.')
param aadClientId string

@secure()
@description('Proxy API app-registration client secret. Optional: when empty, OBO is disabled until set.')
param aadClientSecret string = ''

@description('SharePoint tenant hostname for CORS, e.g. contoso.sharepoint.com.')
param sharepointTenantHostname string

@description('Foundry project endpoint URL.')
param foundryProjectEndpoint string = ''

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
  }
}

output AZURE_LOCATION string = location
output AZURE_TENANT_ID string = aadTenantId
output AZURE_RESOURCE_GROUP string = rg.name
output SERVICE_API_NAME string = resources.outputs.proxyAppName
output SERVICE_API_HOSTNAME string = resources.outputs.proxyAppHostname
output APPLICATIONINSIGHTS_CONNECTION_STRING string = resources.outputs.appInsightsConnectionString
output API_APP_RESOURCE_ID string = 'api://${aadClientId}'
