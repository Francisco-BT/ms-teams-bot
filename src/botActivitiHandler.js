const {
  TeamsActivityHandler,
  MessageFactory,
  CardFactory,
  ActionTypes,
  TurnContext,
} = require("botbuilder");

class BotActivityHandler extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (ctx, next) => {
      // console.log("ctx: ", ctx.activity);
      const text = ctx.activity.text;

      // console.log("original: ", ctx.activity.text);
      const modifiedText = TurnContext.removeMentionText(
        ctx.activity,
        ctx.activity.recipient.id
      );
      // console.log("modifiedText: ", modifiedText);

      if (modifiedText !== text) {
        await ctx.sendActivity(`Hi ${ctx.activity.from.name}`);
      }

      await next();
    });
  }

  async onInstallationUpdateAddActivity(context) {
    try {
      console.log("context -> ", context.activity);
      if (context.activity.conversation.conversationType === "channel") {
        return await context.sendActivity(
          MessageFactory.text(
            `Welcome to Microsoft Teams conversationUpdate events demo bot. This bot is configured in ${context.activity.conversation.name}`
          )
        );
      } else {
        return await context.sendActivity(
          MessageFactory.text(
            "Welcome to Microsoft Teams conversationUpdate events demo bot."
          )
        );
      }
    } catch (error) {
      console.error("error in installantion update activity: ", error);
    }
  }

  async sendIntroCard(context) {
    const card = CardFactory.heroCard(
      "Welcome to Bot Framework!",
      "Welcome to Welcome Users bot sample! This Introduction card is a great way to introduce your Bot to the user and suggest some things to get them started. We use this opportunity to recommend a few next steps for learning more creating and deploying bots.",
      ["https://aka.ms/bf-welcome-card-image"],
      [
        {
          type: ActionTypes.OpenUrl,
          title: "Get an overview",
          value:
            "https://docs.microsoft.com/en-us/azure/bot-service/?view=azure-bot-service-4.0",
        },
        {
          type: ActionTypes.OpenUrl,
          title: "Ask a question",
          value: "https://stackoverflow.com/questions/tagged/botframework",
        },
        {
          type: ActionTypes.OpenUrl,
          title: "Learn how to deploy",
          value:
            "https://docs.microsoft.com/en-us/azure/bot-service/bot-builder-howto-deploy-azure?view=azure-bot-service-4.0",
        },
      ]
    );

    await context.sendActivity({ attachments: [card] });
  }
}

module.exports.BotActivityHandler = BotActivityHandler;
