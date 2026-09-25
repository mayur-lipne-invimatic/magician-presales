const config = {
  MicrosoftAppId: process.env.MicrosoftAppId || process.env.BOT_ID,
  MicrosoftAppType: process.env.BOT_TYPE || 'SingleTenant',
  MicrosoftAppTenantId: process.env.MicrosoftAppTenantId || process.env.BOT_TENANT_ID,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword || process.env.SECRET_BOT_PASSWORD,
};

module.exports = config;
