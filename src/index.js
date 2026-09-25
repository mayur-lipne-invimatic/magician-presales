// src/index.js
const restify = require("restify");
const { commandApp } = require("./internal/initialize");
const { TeamsBot } = require("./teamsBot");
const { sendMessageToTeams } = require("./messages/sendMessageToTeams");
const { adapter } = require("./internal/adapter");
const config = require("./internal/config");

const MICROSOFT_APP_ID = config.MicrosoftAppId;

// A UserAssignedMSI bot has no password - the managed identity supplies its own tokens.
if (!MICROSOFT_APP_ID) {
    console.error("[FATAL] Set MicrosoftAppId (or BOT_ID).");
    process.exit(1);
}
if (config.MicrosoftAppType !== "UserAssignedMSI" && !config.MicrosoftAppPassword) {
    console.error("[FATAL] Set MicrosoftAppPassword (or BOT_PASSWORD/SECRET_BOT_PASSWORD).");
    process.exit(1);
}

const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`App listening`);
    console.log("Bot AppId (last 8):", MICROSOFT_APP_ID.slice(-8));
    console.log("Bot AppType:", config.MicrosoftAppType);
});

const teamsBot = new TeamsBot();

// Bot messages
server.post("/api/messages", async (req, res) => {
    const body = req.body || {};
    console.log("Incoming activity:", JSON.stringify({
        type: body.type,
        name: body.name,
        eventType: body.channelData && body.channelData.eventType,
        conversationType: body.conversation && body.conversation.conversationType,
        conversationId: body.conversation && body.conversation.id,
        hasText: Boolean(body.text),
        text: body.text ? String(body.text).slice(0, 200) : undefined,
    }));
    try {
        // Let TeamsFx ConversationBot process the request; optionally run TeamsBot logic inside the callback.
        await commandApp.requestHandler(req, res, async (turnContext) => {
            console.log("Bot logic running for activity type:", turnContext.activity.type);
            await teamsBot.run(turnContext);
        });
    } catch (err) {
        console.error("Error handling /api/messages:", err);
        throw err;
    }
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
