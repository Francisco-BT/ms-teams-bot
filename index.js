const express = require("express");
const morgan = require("morgan");
const {
  ConfigurationBotFrameworkAuthentication,
  CloudAdapter,
} = require("botbuilder");

const { BotActivityHandler } = require("./src/botActivitiHandler");

const server = express();
const PORT = 3978;
server.use(morgan("dev"));
server.use(express.json());

server.listen(PORT, () => {
  console.log(`\n${server.name} listening on ${PORT}`);
  console.log(
    "\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator"
  );
  console.log('\nTo talk to your bot, open the emulator select "Open Bot"');
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppType: "",
  MicrosoftAppId: "",
  MicrosoftAppPassword: "",
  MicrosoftAppTenantId: "",
});
const adapter = new CloudAdapter(botFrameworkAuthentication);
const bot = new BotActivityHandler();

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights. See https://aka.ms/bottelemetry for telemetry
  //       configuration instructions.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity("The bot encountered an error or bug.");
  await context.sendActivity(
    "To continue to run this bot, please fix the bot source code."
  );
};

server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});
