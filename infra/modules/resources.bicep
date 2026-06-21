@description('Azure region.')
param location string

@description('Resource token for globally-unique names.')
param resourceToken string

@description('Tags applied to all resources.')
param tags object

param aadTenantId string
param aadClientId string
@secure()
param aadClientSecret string
param sharepointTenantHostname string
param foundryProjectEndpoint string
param principalId string

var planName = 'plan-${resourceToken}'
var proxyAppName = 'app-proxy-${resourceToken}'
var laName = 'log-${resourceToken}'
var aiName = 'appi-${resourceToken}'
var kvName = 'kv-${take(resourceToken, 20)}'
var uamiName = 'id-${resourceToken}'

// Role definition IDs
var kvSecretsUserRoleId = '4633458b-17de-408a-b874-0445c86b69e6' // Key Vault Secrets User
var kvSecretsOfficerRoleId = 'b86a8fe4-44ce-4948-aee5-eccb2c155cd7' // Key Vault Secrets Officer (for azd principal to seed)

var hasClientSecret = !empty(aadClientSecret)

// User-assigned managed identity shared by both web apps
module uami 'br/public:avm/res/managed-identity/user-assigned-identity:0.4.1' = {
  name: 'uami-deploy'
  params: {
    name: uamiName
    location: location
    tags: tags
  }
}

// Log Analytics workspace
module la 'br/public:avm/res/operational-insights/workspace:0.9.1' = {
  name: 'la-deploy'
  params: {
    name: laName
    location: location
    tags: tags
    skuName: 'PerGB2018'
    dataRetention: 30
  }
}

// Application Insights
module ai 'br/public:avm/res/insights/component:0.4.2' = {
  name: 'ai-deploy'
  params: {
    name: aiName
    location: location
    tags: tags
    workspaceResourceId: la.outputs.resourceId
    kind: 'web'
    applicationType: 'web'
  }
}

// Key Vault — UAMI gets read; azd principal gets officer for seeding
module kv 'br/public:avm/res/key-vault/vault:0.10.2' = {
  name: 'kv-deploy'
  params: {
    name: kvName
    location: location
    tags: tags
    sku: 'standard'
    enableRbacAuthorization: true
    enableSoftDelete: true
    enablePurgeProtection: false
    publicNetworkAccess: 'Enabled'
    roleAssignments: concat(
      [
        {
          principalId: uami.outputs.principalId
          roleDefinitionIdOrName: kvSecretsUserRoleId
          principalType: 'ServicePrincipal'
        }
      ],
      empty(principalId) ? [] : [
        {
          principalId: principalId
          roleDefinitionIdOrName: kvSecretsOfficerRoleId
          principalType: 'User'
        }
      ])
    secrets: hasClientSecret ? [
      {
        name: 'AzureAd--ClientSecret'
        value: aadClientSecret
      }
    ] : []
  }
}

// Shared App Service plan (Linux, B1 for demo)
module plan 'br/public:avm/res/web/serverfarm:0.4.1' = {
  name: 'plan-deploy'
  params: {
    name: planName
    location: location
    tags: tags
    skuName: 'B1'
    skuCapacity: 1
    kind: 'linux'
    reserved: true
  }
}

// SharePoint Proxy Web API (OBO) — azd service "api"
module proxyApp 'br/public:avm/res/web/site:0.11.1' = {
  name: 'proxy-deploy'
  params: {
    name: proxyAppName
    location: location
    tags: union(tags, { 'azd-service-name': 'api' })
    kind: 'app,linux'
    serverFarmResourceId: plan.outputs.resourceId
    managedIdentities: {
      userAssignedResourceIds: [
        uami.outputs.resourceId
      ]
    }
    siteConfig: {
      linuxFxVersion: 'DOTNETCORE|8.0'
      alwaysOn: true
      ftpsState: 'Disabled'
      minTlsVersion: '1.2'
      http20Enabled: true
      cors: {
        allowedOrigins: [
          'https://${sharepointTenantHostname}'
        ]
        supportCredentials: true
      }
    }
    httpsOnly: true
    appSettingsKeyValuePairs: union(
      {
        APPLICATIONINSIGHTS_CONNECTION_STRING: ai.outputs.connectionString
        APPLICATIONINSIGHTS_AUTHENTICATION_STRING: 'ClientId=${uami.outputs.clientId};Authorization=AAD'

        'AzureAd__Instance': 'https://login.microsoftonline.com/'
        'AzureAd__TenantId': aadTenantId
        'AzureAd__ClientId': aadClientId
        'AzureAd__Audience': 'api://${aadClientId}'

        'SharePoint__RootSiteUrl': 'https://${sharepointTenantHostname}'
        'SharePoint__TenantHostname': sharepointTenantHostname

        'Foundry__ProjectEndpoint': foundryProjectEndpoint
        'Foundry__AgentName': 'SharePointProfileAgent'
        // The proxy calls ONLY the hosted agent (tool-agnostic). The agent owns its own tools.
        'Foundry__AgentResponsesUrl': '${foundryProjectEndpoint}/agents/SharePointProfileAgent/endpoint/protocols/openai/responses?api-version=v1'

        'KeyVault__Uri': kv.outputs.uri
        'AZURE_CLIENT_ID': uami.outputs.clientId
      },
      hasClientSecret ? {
        'AzureAd__ClientSecret': '@Microsoft.KeyVault(VaultName=${kvName};SecretName=AzureAd--ClientSecret)'
      } : {}
    )
  }
}

output proxyAppName string = proxyApp.outputs.name
output proxyAppHostname string = 'https://${proxyApp.outputs.defaultHostname}'
output keyVaultName string = kv.outputs.name
output keyVaultEndpoint string = kv.outputs.uri
output appInsightsConnectionString string = ai.outputs.connectionString
output userAssignedIdentityClientId string = uami.outputs.clientId
output userAssignedIdentityPrincipalId string = uami.outputs.principalId
