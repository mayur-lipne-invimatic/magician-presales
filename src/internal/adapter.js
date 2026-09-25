const {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
    ConfigurationServiceClientCredentialFactory,
} = require("botbuilder");
const {
    ManagedIdentityServiceClientCredentialsFactory,
    JwtTokenProviderFactory,
} = require("botframework-connector");
const config = require("./config");

// One adapter for the whole process: the inbound /api/messages turn and the outbound
// proactive sends both go through it. index.js used to build a second BotFrameworkAdapter
// of its own, which is how the two drifted apart and only one of them worked on Azure.
// BotFrameworkAdapter also can't be used at all now - it only knows appId/appPassword,
// and a UserAssignedMSI bot has no password.
function createCredentialsFactory() {
    if (config.MicrosoftAppType === "UserAssignedMSI") {
        // Tokens come from the user-assigned managed identity attached to the App Service,
        // resolved by MicrosoftAppId (the identity's client id). Nothing to keep secret.
        return new ManagedIdentityServiceClientCredentialsFactory(
            config.MicrosoftAppId,
            new JwtTokenProviderFactory()
        );
    }

    return new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: config.MicrosoftAppId,
        MicrosoftAppPassword: config.MicrosoftAppPassword,
        MicrosoftAppType: config.MicrosoftAppType,
        MicrosoftAppTenantId: config.MicrosoftAppTenantId,
    });
}

const adapter = new CloudAdapter(
    new ConfigurationBotFrameworkAuthentication(
        {
            MicrosoftAppType: config.MicrosoftAppType,
            MicrosoftAppTenantId: config.MicrosoftAppTenantId,
        },
        createCredentialsFactory()
    )
);

module.exports = { adapter };
