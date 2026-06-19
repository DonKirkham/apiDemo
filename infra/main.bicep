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

@secure()
@description('PEM (certificate + private key) for SharePoint app-only certificate auth. Leave empty to use the client secret instead.')
param sharePointCertPem string = ''

@description('Always Ready (pre-warmed) instances kept running to avoid cold starts. 1 keeps a single instance warm; 0 disables and scales fully to zero. Always Ready instances are billed continuously, even when idle.')
@minValue(0)
@maxValue(40)
param alwaysReadyInstances int = 1

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
    name: 'FC1'
    tier: 'FlexConsumption'
  }
  kind: 'functionapp'
  properties: {
    reserved: true
  }
}

resource blobService 'Microsoft.Storage/storageAccounts/blobServices@2023-01-01' = {
  parent: storage
  name: 'default'
}

resource deploymentContainer 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-01-01' = {
  parent: blobService
  name: 'app-package'
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
    functionAppConfig: {
      deployment: {
        storage: {
          type: 'blobContainer'
          value: '${storage.properties.primaryEndpoints.blob}${deploymentContainer.name}'
          authentication: {
            type: 'StorageAccountConnectionString'
            storageAccountConnectionStringName: 'AzureWebJobsStorage'
          }
        }
      }
      scaleAndConcurrency: {
        maximumInstanceCount: 40
        instanceMemoryMB: 2048
        // Keep N instances pre-warmed so HTTP calls skip cold start. 'http' covers
        // all HTTP-triggered functions; empty array (param 0) scales fully to zero.
        alwaysReady: alwaysReadyInstances > 0 ? [
          {
            name: 'http'
            instanceCount: alwaysReadyInstances
          }
        ] : []
      }
      runtime: {
        name: 'node'
        version: '22'
      }
    }
    siteConfig: {
      minTlsVersion: '1.2'
      appSettings: concat([
        {
          name: 'AzureWebJobsStorage'
          value: 'DefaultEndpointsProtocol=https;AccountName=${storage.name};EndpointSuffix=${environment().suffixes.storage};AccountKey=${storage.listKeys().keys[0].value}'
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
      ], empty(sharePointCertPem) ? [] : [
        {
          name: 'SHAREPOINT_CERT_PEM'
          value: sharePointCertPem
        }
      ])
    }
  }
}

output functionAppName string = functionApp.name
