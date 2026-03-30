// ============================================================================
// M365 Dashboard - Azure Infrastructure
// Deploy with: az deployment sub create --location <region> --template-file main.bicep --parameters main.bicepparam
// ============================================================================

targetScope = 'subscription'

@description('Name prefix for all resources')
param namePrefix string = 'm365dash'

@description('Azure region for resources')
param location string = 'uksouth'

@description('Environment (dev, prod)')
@allowed(['dev', 'prod'])
param environment string = 'prod'

@description('Entra ID (Azure AD) Tenant ID')
param entraIdTenantId string

@description('Entra ID App Registration Client ID')
param entraIdClientId string

@secure()
@description('Entra ID App Registration Client Secret')
param entraIdClientSecret string

@description('Object ID of the deploying user - granted Key Vault Secrets Officer for management access')
param deployingUserObjectId string = ''

@description('SQL Server Administrator Login')
param sqlAdminLogin string = 'sqladmin'

@secure()
@description('SQL Server Administrator Password')
param sqlAdminPassword string

@description('Container image to deploy (leave empty for initial setup)')
param containerImage string = ''

// ============================================================================
// Resource Group
// ============================================================================
resource rg 'Microsoft.Resources/resourceGroups@2023-07-01' = {
  name: '${namePrefix}-${environment}-rg'
  location: location
  tags: {
    environment: environment
    application: 'M365 Dashboard'
  }
}

// ============================================================================
// Deploy Infrastructure
// ============================================================================
module infrastructure 'modules/infrastructure.bicep' = {
  scope: rg
  name: 'infrastructure'
  params: {
    namePrefix: namePrefix
    location: location
    environment: environment
    entraIdTenantId: entraIdTenantId
    entraIdClientId: entraIdClientId
    entraIdClientSecret: entraIdClientSecret
    sqlAdminLogin: sqlAdminLogin
    sqlAdminPassword: sqlAdminPassword
    containerImage: containerImage
    deployingUserObjectId: deployingUserObjectId
  }
}

// ============================================================================
// Outputs
// ============================================================================
output resourceGroupName string = rg.name
output containerAppUrl string = infrastructure.outputs.containerAppUrl
output containerRegistryName string = infrastructure.outputs.containerRegistryName
output containerRegistryLoginServer string = infrastructure.outputs.containerRegistryLoginServer
output sqlServerFqdn string = infrastructure.outputs.sqlServerFqdn
output sqlDatabaseName string = infrastructure.outputs.sqlDatabaseName
output mapsAccountName string = infrastructure.outputs.mapsAccountName
output mapsClientId string = infrastructure.outputs.mapsClientId
output keyVaultUri string = infrastructure.outputs.keyVaultUri
output keyVaultName string = infrastructure.outputs.keyVaultName
