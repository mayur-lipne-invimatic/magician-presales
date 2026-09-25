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

@description('Tenant ID for the single-tenant bot registration (== BOT_TENANT_ID).')
param botAppTenantId string

@maxLength(42)
@minLength(4)
@description('Azure Bot resource name. Separate from the web app name.')
param botServiceName string

@description('Azure Bot SKU. The one SingleTenant bot here that receives channel messages is S1.')
param botServiceSku string = 'S1'

// Optional MSI for Azure resource access (NOT used for Bot Framework auth). The bot cannot
// authenticate as this identity: a UserAssignedMSI registration has only a managed-identity
// service principal and no Entra app registration, so Teams rejects the manifest's RSC block
// with "Requested permission is not recognized" and ChannelMessage.Read.Group is unavailable.
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
        { name: 'MicrosoftAppTenantId', value: botAppTenantId }
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

// Azure Bot registration (replaces the dev.botframework.com registration).
resource botService 'Microsoft.BotService/botServices@2022-09-15' = {
  kind: 'azurebot'
  location: 'global'
  name: botServiceName
  properties: {
    displayName: botDisplayName
    endpoint: 'https://${webApp.properties.defaultHostName}/api/messages'
    msaAppId: botAppId
    msaAppType: 'SingleTenant'
    msaAppTenantId: botAppTenantId
  }
  sku: {
    name: botServiceSku
  }
}

resource botServiceMsTeamsChannel 'Microsoft.BotService/botServices/channels@2022-09-15' = {
  parent: botService
  location: 'global'
  name: 'MsTeamsChannel'
  properties: {
    channelName: 'MsTeamsChannel'
    // Adding the channel through the portal makes you accept the Teams terms of service;
    // a template-created channel defaults to unaccepted. The working bots have it accepted.
    properties: {
      acceptedTerms: true
      isEnabled: true
    }
  }
}

// === Outputs consumed by your YAML / env ===
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ENDPOINT string = 'https://${webApp.properties.defaultHostName}'
