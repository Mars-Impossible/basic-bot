// index.js is used to setup and configure your bot

// Import required packages
const express = require("express");
const mockTeamsApi = require("./mockTeamsApi");

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
} = require("botbuilder");
const { TeamsBot } = require("./teamsBot");
const config = require("./config");

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory(config);

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights. See https://aka.ms/bottelemetry for telemetry
  //       configuration instructions.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Only send error message for user messages, not for other message types so the bot doesn't spam a channel or chat.
  if (context.activity.type === "message") {
    // Send a message to the user
    await context.sendActivity(`The bot encountered an unhandled error:\n ${error.message}`);
    await context.sendActivity("To continue to run this bot, please fix the bot source code.");
  }
};

// Create the bot that will handle incoming messages.
const bot = new TeamsBot();

// Create express application.
const expressApp = express();
expressApp.use(express.json());

// 使用 mock API 路由
expressApp.use(mockTeamsApi);

// 添加一个简单的健康检查端点
expressApp.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

const server = expressApp.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${expressApp.name} listening to`, server.address());
  console.log('Mock API endpoints available at:');
  console.log('- GET    /api/teams/mapping');
  console.log('- POST   /api/teams/session');
  console.log('- DELETE /api/teams/session');
  console.log('- PUT    /api/teams/session/activity');
});

// Listen for incoming requests.
expressApp.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await bot.run(context);
  });
});

// Gracefully shutdown HTTP server
["exit", "uncaughtException", "SIGINT", "SIGTERM", "SIGUSR1", "SIGUSR2"].forEach((event) => {
  process.on(event, () => {
    server.close();
  });
});
