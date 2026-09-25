// src/index.js
const restify = require("restify");
const { commandApp } = require("./internal/initialize");
const { TeamsBot } = require("./teamsBot");
const { BotFrameworkAdapter } = require("botbuilder");
const { sendMessageToTeams } = require("./messages/sendMessageToTeams");
const {
    CloudAdapter,
    ConfigurationBotFrameworkAuthentication,
    ConfigurationServiceClientCredentialFactory
} = require("botbuilder");
const { TurnContext } = require("botbuilder");

const MICROSOFT_APP_ID = process.env.MicrosoftAppId || process.env.BOT_ID;
const MICROSOFT_APP_PASSWORD = process.env.MicrosoftAppPassword || process.env.SECRET_BOT_PASSWORD;
const MICROSOFT_APP_TENANT_ID = process.env.MicrosoftAppTenantId || process.env.BOT_TENANT_ID;
 
if (!MICROSOFT_APP_ID || !MICROSOFT_APP_PASSWORD) {
  console.error("[FATAL] Set MicrosoftAppId/MicrosoftAppPassword (or BOT_ID/SECRET_BOT_PASSWORD).");
  process.exit(1);
}

const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`App listening`);
  console.log("Bot AppId (last 8):", MICROSOFT_APP_ID.slice(-8));
});

const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: MICROSOFT_APP_ID,
    MicrosoftAppPassword: MICROSOFT_APP_PASSWORD,
    MicrosoftAppType: 'SingleTenant',
    MicrosoftAppTenantId: MICROSOFT_APP_TENANT_ID
});

	
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
    {},
    credentialsFactory
);

// 3. Create the CloudAdapter
const adapter = new CloudAdapter(botFrameworkAuthentication);
// Catch-all for errors
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${error}`);
    await context.sendActivity('The bot encountered an error or bug.');
};

const teamsBot = new TeamsBot();

// Bot messages
server.post("/api/messages", async (req, res) => {
  await commandApp.requestHandler(req, res, async (turnContext) => {
    await teamsBot.run(turnContext);
  });
});

// Proactive send (server generates its own BF token; client bearer ignored)
server.post("/api/sendToTeams", async (req, res) => {
  try {
    console.log("Request received: ", req.body);
    const { conversationReference, cardContent, messageText, activityId, isUser, updateMode } = req.body || {};
    if (!conversationReference || (!cardContent && !messageText)) {
      return res.send(400, { error: "Provide conversationReference and messageText or cardContent" });
    }
    const details = await sendMessageToTeams(
      adapter,
      conversationReference,
      cardContent,
      messageText,
      activityId,
      isUser,
      updateMode
    );
    res.send(200, { status: "Message processed successfully", details });
  } catch (err) {
    console.error("Error sending message to Teams:", err);
    res.send(err.statusCode || 500, err.body || err.message || "Internal Server Error");
  }
});
