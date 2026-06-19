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
param foundryAgentId string
param principalId string

var storageName = 'st${resourceToken}'
var planName = 'plan-${resourceToken}'
var functionAppName = 'func-${resourceToken}'
var laName = 'log-${resourceToken}'
var aiName = 'appi-${resourceToken}'
var kvName = 'kv-${take(resourceToken, 20)}'
var uamiName = 'id-${resourceToken}'

// Role definition IDs
var kvSecretsUserRoleId = '4633458b-17de-408a-b874-0445c86b69e6' // Key Vault Secrets User
var kvSecretsOfficerRoleId = 'b86a8fe4-44ce-4948-aee5-eccb2c155cd7' // Key Vault Secrets Officer (for azd principal to seed)
var storageBlobDataOwnerRoleId = 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b'

// User-assigned managed identity for the Function App
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

// Storage account (required by Functions)
module storage 'br/public:avm/res/storage/storage-account:0.14.3' = {
  name: 'storage-deploy'
  params: {
    name: storageName
    location: location
    tags: tags
    skuName: 'Standard_LRS'
    kind: 'StorageV2'
    allowBlobPublicAccess: false
    minimumTlsVersion: 'TLS1_2'
    publicNetworkAccess: 'Enabled'
    networkAcls: {
      defaultAction: 'Allow'
      bypass: 'AzureServices'
    }
    roleAssignments: [
      {
        principalId: uami.outputs.principalId
        roleDefinitionIdOrName: storageBlobDataOwnerRoleId
        principalType: 'ServicePrincipal'
      }
    ]
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
    secrets: [
      {
        name: 'AzureAd--ClientSecret'
        value: aadClientSecret
      }
    ]
  }
}

// Function App (Flex Consumption-style via AVM site module)
module functionApp 'br/public:avm/res/web/site:0.11.1' = {
  name: 'func-deploy'
  params: {
    name: functionAppName
    location: location
    tags: union(tags, { 'azd-service-name': 'api' })
    kind: 'functionapp,linux'
    serverFarmResourceId: plan.outputs.resourceId
    managedIdentities: {
      userAssignedResourceIds: [
        uami.outputs.resourceId
      ]
    }
    siteConfig: {
      linuxFxVersion: 'DOTNET-ISOLATED|8.0'
      alwaysOn: false
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
    appSettingsKeyValuePairs: {
      AzureWebJobsStorage__accountName: storage.outputs.name
      AzureWebJobsStorage__credential: 'managedidentity'
      AzureWebJobsStorage__clientId: uami.outputs.clientId
      FUNCTIONS_EXTENSION_VERSION: '~4'
      FUNCTIONS_WORKER_RUNTIME: 'dotnet-isolated'
      APPLICATIONINSIGHTS_CONNECTION_STRING: ai.outputs.connectionString
      APPLICATIONINSIGHTS_AUTHENTICATION_STRING: 'ClientId=${uami.outputs.clientId};Authorization=AAD'

      'AzureAd__Instance': 'https://login.microsoftonline.com/'
      'AzureAd__TenantId': aadTenantId
      'AzureAd__ClientId': aadClientId
      'AzureAd__Audience': 'api://${aadClientId}'
      'AzureAd__ClientSecret': '@Microsoft.KeyVault(VaultName=${kvName};SecretName=AzureAd--ClientSecret)'

      'Graph__Scopes': 'User.Read'
      'SharePoint__RootSiteUrl': 'https://${sharepointTenantHostname}'
      'SharePoint__Scopes': 'AllSites.Read User.Read.All'

      'Foundry__ProjectEndpoint': foundryProjectEndpoint
      'Foundry__AgentId': foundryAgentId

      'SharePoint__TenantHostname': sharepointTenantHostname

      'KeyVault__Uri': kv.outputs.uri

      'AZURE_CLIENT_ID': uami.outputs.clientId
    }
  }
  dependsOn: [
    storage
    ai
    kv
  ]
}

// App Service plan (Linux, B1 for demo — swap for FlexConsumption when AVM module supports it cleanly)
module plan 'br/public:avm/res/web/serverfarm:0.4.1' = {
  name: 'plan-deploy'
  params: {
    name: planName
    location: location
    tags: tags
    skuName: 'B1'
    skuCapacity: 1
    kind: 'Linux'
    reserved: true
  }
}

output functionAppName string = functionApp.outputs.name
output functionAppHostname string = 'https://${functionApp.outputs.defaultHostname}'
output keyVaultName string = kv.outputs.name
output keyVaultEndpoint string = kv.outputs.uri
output appInsightsConnectionString string = ai.outputs.connectionString
output userAssignedIdentityClientId string = uami.outputs.clientId
output userAssignedIdentityPrincipalId string = uami.outputs.principalId
