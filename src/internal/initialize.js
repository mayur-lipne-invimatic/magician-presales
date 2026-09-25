const {
    BotBuilderCloudAdapter
} = require("@microsoft/teamsfx");
const ConversationBot = BotBuilderCloudAdapter.ConversationBot;
const {
    LoopsCommandHandler
} = require("../commands/loopsCommandHandler");
const config = require("./config");

// Create the command bot and register the command handlers for your app.
// You can also use the commandApp.command.registerCommands to register other commands
// if you don't want to register all of them in the constructor

const commandApp = new ConversationBot({
    adapterConfig: {
        MicrosoftAppId: config.MicrosoftAppId,
        MicrosoftAppPassword: config.MicrosoftAppPassword,
        MicrosoftAppType: "SingleTenant",
        MicrosoftAppTenantId: config.MicrosoftAppTenantId
    },
    command: {
        enabled: true,
        commands: [new LoopsCommandHandler()],
    }
});

module.exports = {
    commandApp
};
