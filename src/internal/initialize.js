const {
    BotBuilderCloudAdapter
} = require("@microsoft/teamsfx");
const ConversationBot = BotBuilderCloudAdapter.ConversationBot;
const {
    LoopsCommandHandler
} = require("../commands/loopsCommandHandler");
const { adapter } = require("./adapter");

// Create the command bot and register the command handlers for your app.
// You can also use the commandApp.command.registerCommands to register other commands
// if you don't want to register all of them in the constructor
const commandApp = new ConversationBot({
    // Hand ConversationBot the shared adapter rather than an `adapterConfig` for it to
    // build one from. When it builds its own it can only do appId/appPassword, so a
    // managed-identity bot ends up with no way to mint an outbound token: inbound webhook
    // calls still get processed (a missing app id falls back to an unauthenticated,
    // emulator-style accept), but anything this adapter sends on its own turn context -
    // like the context.updateActivity() that rewrites a card after a like/dislike tap -
    // gets a 401 "Authorization has been denied for this request" from the Bot Connector.
    adapter,
    command: {
        enabled: true,
        commands: [new LoopsCommandHandler()],
    }
});

module.exports = {
    commandApp,
};
