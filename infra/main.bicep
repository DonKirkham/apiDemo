@description('Deployment location')
param location string = resourceGroup().location

@description('Globally unique Function App name')
param functionAppName string

@description('Globally unique storage account name')
param storageAccountName string

@description('SharePoint App Registration client id')
param sharePointClientId string

@secure()
@description('SharePoint App Registration client secret')
param sharePointClientSecret string

@description('Azure AD tenant id')
param tenantId string

@description('SharePoint resource scope for client credentials')
param sharePointResource string

@description('Comma-separated allowed SPFx origins')
param allowedCallerDomains string

@description('Audience for inbound Entra token validation')
param entraAudience string

@description('Comma-separated client app ids allowed to call Entra endpoint')
param allowedClientAppIds string = ''

resource storage 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  name: storageAccountName
  location: location
  sku: {
    name: 'Standard_LRS'
  }
  kind: 'StorageV2'
  properties: {
    minimumTlsVersion: 'TLS1_2'
    supportsHttpsTrafficOnly: true
    allowBlobPublicAccess: false
  }
}

resource hostingPlan 'Microsoft.Web/serverfarms@2023-12-01' = {
  name: '${functionAppName}-plan'
  location: location
  sku: {
    name: 'Y1'
    tier: 'Dynamic'
  }
  kind: 'functionapp'
  properties: {
    reserved: true
  }
}

resource appInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: '${functionAppName}-appi'
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
  }
}

resource functionApp 'Microsoft.Web/sites@2023-12-01' = {
  name: functionAppName
  location: location
  kind: 'functionapp,linux'
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    httpsOnly: true
    serverFarmId: hostingPlan.id
    siteConfig: {
      linuxFxVersion: 'Node|20'
      minTlsVersion: '1.2'
      appSettings: [
        {
          name: 'FUNCTIONS_EXTENSION_VERSION'
          value: '~4'
        }
        {
          name: 'FUNCTIONS_WORKER_RUNTIME'
          value: 'node'
        }
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1'
        }
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storage.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${listKeys(storage.id, storage.apiVersion).keys[0].value}'
        }
        {
          name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
          value: appInsights.properties.ConnectionString
        }
        {
          name: 'ALLOWED_CALLER_DOMAINS'
          value: allowedCallerDomains
        }
        {
          name: 'AZURE_TENANT_ID'
          value: tenantId
        }
        {
          name: 'SHAREPOINT_APP_CLIENT_ID'
          value: sharePointClientId
        }
        {
          name: 'SHAREPOINT_APP_CLIENT_SECRET'
          value: sharePointClientSecret
        }
        {
          name: 'SHAREPOINT_RESOURCE'
          value: sharePointResource
        }
        {
          name: 'ENTRA_AUDIENCE'
          value: entraAudience
        }
        {
          name: 'ALLOWED_CLIENT_APP_IDS'
          value: allowedClientAppIds
        }
      ]
    }
  }
}

output functionAppName string = functionApp.name
