// Local dev's .localConfigs (generated per m365agents.local.yml) sets BOT_ID/BOT_PASSWORD/
// BOT_TYPE. The deployed Azure Web App (per infra/azure.bicep's appSettings) only sets
// MicrosoftAppId/MicrosoftAppPassword/MicrosoftAppType - it never sets BOT_ID/BOT_PASSWORD/
// BOT_TYPE at all. Reading only the BOT_* names (as this used to) meant every field here came
// back undefined on Azure, so the CloudAdapter commandApp builds from this config had no real
// credentials - inbound webhook calls still got processed (missing app id falls back to an
// unauthenticated/emulator-style accept), but any outbound call it made on its own turn context
// (like context.updateActivity() after a like/dislike tap) had nothing to authenticate with,
// so the Bot Connector rejected it with 401. Mirrors the fallback index.js already uses at
// the top of the file for its own separate adapter.
const config = {
  MicrosoftAppId: process.env.MicrosoftAppId || process.env.BOT_ID,
  MicrosoftAppType: process.env.MicrosoftAppType || process.env.BOT_TYPE,
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId || process.env.BOT_TENANT_ID,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword || process.env.BOT_PASSWORD || process.env.SECRET_BOT_PASSWORD,
};

module.exports = config;
