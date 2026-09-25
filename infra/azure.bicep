@maxLength(20)
@minLength(4)
@description('Base name used to generate resource names.')
param resourceBaseName string

@description('App Service plan SKU (e.g., B1, S1).')
param webAppSKU string

@maxLength(42)
@description('Display name for your bot.')
param botDisplayName string

@description('App Service plan name (defaults to base name).')
param serverfarmsName string = resourceBaseName

@description('Web App name (defaults to base name).')
param webAppName string = resourceBaseName

@description('User-assigned managed identity name (defaults to base name).')
param identityName string = resourceBaseName

@description('Deployment location.')
param location string = resourceGroup().location

@description('Bot AAD app clientId created by aadApp/create (== BOT_ID).')
param botAppId string

@secure()
@description('Bot AAD app client secret created by aadApp/create (== SECRET_BOT_PASSWORD).')
param botAppPassword string

// Optional MSI for Azure resource access (NOT used for Bot Framework auth)
resource identity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: identityName
  location: location
}

resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  name: serverfarmsName
  location: location
  kind: 'app'
  sku: {
    name: webAppSKU
  }
}

resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  name: webAppName
  location: location
  kind: 'app'
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: true
      appSettings: [
        { name: 'WEBSITE_RUN_FROM_PACKAGE', value: '1' }
        { name: 'WEBSITE_NODE_DEFAULT_VERSION', value: '~18' }
        { name: 'RUNNING_ON_AZURE', value: '1' }

        // ✅ Use the SAME AAD app the pipeline created for the bot identity
        { name: 'MicrosoftAppId', value: botAppId }
        { name: 'MicrosoftAppPassword', value: botAppPassword }
        { name: 'MicrosoftAppType', value: 'SingleTenant' }
        { name: 'MicrosoftAppTenantId', value: tenant().tenantId }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
  // Keep MSI if you need it; it does NOT affect bot auth
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${identity.id}': {}
    }
  }
  tags: {
    'botDisplayName': botDisplayName
  }
}

// === Outputs consumed by your YAML / env ===
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ENDPOINT string = 'https://${webApp.properties.defaultHostName}'
