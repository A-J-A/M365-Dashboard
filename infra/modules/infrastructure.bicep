// ============================================================================
// M365 Dashboard - Infrastructure Module
// ============================================================================

@description('Name prefix for all resources')
param namePrefix string

@description('Azure region')
param location string

@description('Environment')
param environment string

@description('Entra ID Tenant ID')
param entraIdTenantId string

@description('Entra ID Client ID')
param entraIdClientId string

@secure()
@description('Entra ID Client Secret')
param entraIdClientSecret string

@description('SQL Admin Login')
param sqlAdminLogin string

@secure()
@description('SQL Admin Password')
param sqlAdminPassword string

@description('Container image')
param containerImage string

@description('Object ID of the deploying user - granted Key Vault Secrets Officer for management access')
param deployingUserObjectId string = ''

// ============================================================================
// Variables
// ============================================================================
var uniqueSuffix = uniqueString(resourceGroup().id)
var containerRegistryName = '${namePrefix}${uniqueSuffix}acr'
var containerAppEnvName = '${namePrefix}-${environment}-env'
var containerAppName = '${namePrefix}-${environment}-app'
var sqlServerName = '${namePrefix}-${environment}-sql'
var sqlDatabaseName = '${namePrefix}-db'
var logAnalyticsName = '${namePrefix}-${environment}-logs'
var keyVaultName = '${namePrefix}${uniqueSuffix}kv'
var mapsAccountName = '${namePrefix}-${environment}-maps'

// ============================================================================
// Log Analytics Workspace (required for Container Apps)
// ============================================================================
resource logAnalytics 'Microsoft.OperationalInsights/workspaces@2022-10-01' = {
  name: logAnalyticsName
  location: location
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    retentionInDays: 30
  }
}

// ============================================================================
// Azure Container Registry
// ============================================================================
resource containerRegistry 'Microsoft.ContainerRegistry/registries@2023-07-01' = {
  name: containerRegistryName
  location: location
  sku: {
    name: 'Basic'
  }
  properties: {
    adminUserEnabled: true
  }
}

// ============================================================================
// Azure SQL Server & Database
// ============================================================================
resource sqlServer 'Microsoft.Sql/servers@2023-05-01-preview' = {
  name: sqlServerName
  location: location
  properties: {
    administratorLogin: sqlAdminLogin
    administratorLoginPassword: sqlAdminPassword
    version: '12.0'
    minimalTlsVersion: '1.2'
    publicNetworkAccess: 'Enabled'
  }
}

// Allow Azure services to access SQL
resource sqlFirewallAzure 'Microsoft.Sql/servers/firewallRules@2023-05-01-preview' = {
  parent: sqlServer
  name: 'AllowAzureServices'
  properties: {
    startIpAddress: '0.0.0.0'
    endIpAddress: '0.0.0.0'
  }
}

resource sqlDatabase 'Microsoft.Sql/servers/databases@2023-05-01-preview' = {
  parent: sqlServer
  name: sqlDatabaseName
  location: location
  sku: {
    name: 'Basic'
    tier: 'Basic'
    capacity: 5
  }
  properties: {
    collation: 'SQL_Latin1_General_CP1_CI_AS'
    maxSizeBytes: 2147483648 // 2GB
  }
}

// ============================================================================
// Azure Maps Account
// ============================================================================
resource mapsAccount 'Microsoft.Maps/accounts@2023-06-01' = {
  name: mapsAccountName
  location: 'global'
  sku: {
    name: 'G2'
  }
  kind: 'Gen2'
  properties: {
    disableLocalAuth: false
  }
}

// ============================================================================
// Key Vault (for secrets)
// ============================================================================
resource keyVault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: keyVaultName
  location: location
  properties: {
    sku: {
      family: 'A'
      name: 'standard'
    }
    tenantId: subscription().tenantId
    enableRbacAuthorization: true
    enableSoftDelete: true
    softDeleteRetentionInDays: 7
  }
}

// Store all secrets in Key Vault using AzureAd-- naming convention
// (double-dash maps to colon hierarchy in .NET config: AzureAd--TenantId => AzureAd:TenantId)
resource secretTenantId 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'AzureAd--TenantId'
  properties: {
    value: entraIdTenantId
  }
}

resource secretClientId 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'AzureAd--ClientId'
  properties: {
    value: entraIdClientId
  }
}

resource secretEntraClientSecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'AzureAd--ClientSecret'
  properties: {
    value: entraIdClientSecret
  }
}

resource secretAudience 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'AzureAd--Audience'
  properties: {
    value: 'api://${entraIdClientId}'
  }
}

resource secretConnectionString 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'ConnectionStrings--DefaultConnection'
  properties: {
    value: sqlConnectionString
  }
}

resource secretMapsKey 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  parent: keyVault
  name: 'AzureMaps--SubscriptionKey'
  properties: {
    value: mapsAccount.listKeys().primaryKey
  }
}

// ============================================================================
// Container Apps Environment
// ============================================================================
// Note: sqlConnectionString is also stored in Key Vault above.
// It is defined before containerAppEnv to ensure it is available for the secret resource.
resource containerAppEnv 'Microsoft.App/managedEnvironments@2023-05-01' = {
  name: containerAppEnvName
  location: location
  properties: {
    appLogsConfiguration: {
      destination: 'log-analytics'
      logAnalyticsConfiguration: {
        customerId: logAnalytics.properties.customerId
        sharedKey: logAnalytics.listKeys().primarySharedKey
      }
    }
  }
}

// ============================================================================
// Container App
// ============================================================================
var sqlConnectionString = 'Server=tcp:${sqlServer.properties.fullyQualifiedDomainName},1433;Initial Catalog=${sqlDatabaseName};Persist Security Info=False;User ID=${sqlAdminLogin};Password=${sqlAdminPassword};MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;'

resource containerApp 'Microsoft.App/containerApps@2023-05-01' = {
  name: containerAppName
  location: location
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    managedEnvironmentId: containerAppEnv.id
    configuration: {
      ingress: {
        external: true
        targetPort: 8080
        transport: 'http'
        allowInsecure: false
      }
      registries: [
        {
          server: containerRegistry.properties.loginServer
          username: containerRegistry.listCredentials().username
          passwordSecretRef: 'acr-password'
        }
      ]
      secrets: [
        {
          // ACR pull credentials - needed for image pull, not stored in Key Vault
          name: 'acr-password'
          value: containerRegistry.listCredentials().passwords[0].value
        }
      ]
    }
    template: {
      containers: [
        {
          name: 'm365dashboard'
          image: !empty(containerImage) ? containerImage : 'mcr.microsoft.com/azuredocs/containerapps-helloworld:latest'
          resources: {
            cpu: json('0.5')
            memory: '1Gi'
          }
          env: [
            {
              // Only non-secret config is set as env vars.
              // All secrets (TenantId, ClientId, ClientSecret, ConnectionString, etc.)
              // are loaded at runtime by Program.cs via Key Vault + Managed Identity.
              name: 'ASPNETCORE_ENVIRONMENT'
              value: 'Production'
            }
            {
              // Explicitly set the listening port to match the Container App targetPort.
              // Container Apps injects PORT env var but ASP.NET needs ASPNETCORE_URLS.
              name: 'ASPNETCORE_URLS'
              value: 'http://+:8080'
            }
            {
              name: 'KeyVault__Uri'
              value: keyVault.properties.vaultUri
            }
            {
              name: 'AzureAd__Instance'
              value: 'https://login.microsoftonline.com/'
            }
            {
              name: 'AzureMaps__ClientId'
              value: mapsAccount.properties.uniqueId
            }
          ]
        }
      ]
      scale: {
        minReplicas: 1
        maxReplicas: 3
        rules: [
          {
            name: 'http-scaling'
            http: {
              metadata: {
                concurrentRequests: '100'
              }
            }
          }
        ]
      }
    }
  }
}

// Grant Container App managed identity access to Key Vault secrets (read-only)
resource keyVaultRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: keyVault
  name: guid(keyVault.id, containerApp.id, 'Key Vault Secrets User')
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', '4633458b-17de-408a-b874-0445c86b69e6') // Key Vault Secrets User
    principalId: containerApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// Grant the deploying user Key Vault Secrets Officer access so they can view/manage secrets
// This avoids needing to manually grant access after deployment
resource keyVaultUserRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!empty(deployingUserObjectId)) {
  scope: keyVault
  name: guid(keyVault.id, deployingUserObjectId, 'Key Vault Secrets Officer')
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', 'b86a8fe4-44ce-4948-aee5-eccb2c155cd7') // Key Vault Secrets Officer
    principalId: deployingUserObjectId
    principalType: 'User'
  }
}

// ============================================================================
// Outputs
// ============================================================================
output containerAppUrl string = 'https://${containerApp.properties.configuration.ingress.fqdn}'
output containerRegistryName string = containerRegistry.name
output containerRegistryLoginServer string = containerRegistry.properties.loginServer
output sqlServerFqdn string = sqlServer.properties.fullyQualifiedDomainName
output sqlDatabaseName string = sqlDatabase.name
output keyVaultUri string = keyVault.properties.vaultUri
output keyVaultName string = keyVault.name
output mapsAccountName string = mapsAccount.name
output mapsClientId string = mapsAccount.properties.uniqueId
