// src/messages/sendMessageToTeams.js
const { MessageFactory } = require("botbuilder");
const { MicrosoftAppCredentials } = require("botframework-connector");

/**
 * Ensure we have a ConversationReference-like object the adapter can use.
 * Fills in bot.id from your AppId if it's missing.
 */
function normalizeConversationReference(input, appId) {
  const conv = input.conversation || {};
  const bot = input.bot || {};
  const ref = {
    serviceUrl: input.serviceUrl,                  // e.g., https://smba.trafficmanager.net/amer/
    channelId: "msteams",
    bot: {
      id: bot.id || (appId ? `28:${appId}` : undefined),
      name: bot.name || "Bot",
    },
    conversation: {
      id: conv.id,                                 // '19:...@thread.tacv2'
      tenantId: conv.tenantId || input.tenantId,  // Teams tenant GUID
      isGroup: true,
      conversationType: conv.conversationType || "channel",
    },
  };
  return ref;
}

/**
 * Send or update a message in a Teams channel using the BotFramework Adapter.
 * IMPORTANT: The adapter will mint a Bot Framework token using your bot's AppId/secret.
 *            Any client Authorization header is ignored at this layer.
 *
 * @param {BotFrameworkAdapter} adapter
 * @param {object} conversationReference - must contain serviceUrl + conversation.id (thread) + tenantId
 * @param {object} [cardContent]         - optional attachment (e.g., Adaptive Card)
 * @param {string} [messageText]         - optional text
 * @param {string} [activityId]          - when updateMode=true, the activity id to update
 * @param {boolean} [isUser]             - ignored; always sends as the bot
 * @param {boolean} [updateMode]         - if true, updates an existing activity
 * @returns {Promise<object>}
 */
async function sendMessageToTeams(
  adapter,
  conversationReference,
  cardContent,
  messageText,
  activityId,
  isUser,
  updateMode
) {
  const appId = process.env.MicrosoftAppId || process.env.BOT_ID;

  // Build a safe reference
  const ref = normalizeConversationReference(conversationReference, appId);

  // Guardrails
  if (!ref.serviceUrl) throw new Error("conversationReference.serviceUrl is required.");
  if (!ref.conversation?.id) throw new Error("conversationReference.conversation.id (threadId) is required.");
  if (!ref.conversation?.tenantId)
    throw new Error("conversationReference.conversation.tenantId (Teams tenant) is required.");
  if (!ref.bot?.id)
    throw new Error("Bot id missing. Ensure manifest botId and MicrosoftAppId match the installed bot.");

  let result = {};

  MicrosoftAppCredentials.trustServiceUrl(ref.serviceUrl);
  await adapter.continueConversationAsync(process.env.MicrosoftAppId, ref, async (context) => {
    try {
      // Build activity
      const activity = cardContent
        ? { type: "message", attachments: [cardContent] }
        : MessageFactory.text(messageText ?? "");

      // Teams targeting: channel (thread) & tenant
      activity.channelData = {
        channel: { id: ref.conversation.id },
        tenant: { id: ref.conversation.tenantId },
      };

      if (updateMode && activityId) {
        const updateResp = await context.updateActivity({ id: activityId, ...activity });
        result = { mode: "update", activityId, updateResp };
      } else {
        const sendResp = await context.sendActivity(activity); // returns { id: '<activityId>' }
        result = { mode: "send", activityId: sendResp?.id };
      }
    } catch (e) {
      // Build activity
      const activity = cardContent
        ? { type: "message", attachments: [cardContent] }
        : MessageFactory.text(messageText ?? "");

      // Teams targeting: channel (thread) & tenant
      activity.channelData = {
        channel: { id: ref.conversation.id },
        tenant: { id: ref.conversation.tenantId },
      };

      if (updateMode && activityId) {
        const updateResp = await context.updateActivity({ id: activityId, ...activity });
        result = { mode: "update", activityId, updateResp };
      } else {
        const sendResp = await context.sendActivity(activity); // returns { id: '<activityId>' }
        result = { mode: "send", activityId: sendResp?.id };
      }
    }
    
  });

  return {
    ok: true,
    ...result,
    serviceUrl: ref.serviceUrl,
    threadId: ref.conversation.id,
    tenantId: ref.conversation.tenantId,
  };
}

module.exports = { sendMessageToTeams };
