const { TeamsActivityHandler, TurnContext, CardFactory } = require("botbuilder");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();

      if (txt === "learn") {
        const card = CardFactory.adaptiveCard({
          type: "AdaptiveCard",
          version: "1.0",
          body: [
            {
              type: "TextBlock",
              text: "Learn Adaptive Card and Commands",
              weight: "bolder",
              size: "medium"
            },
            {
              type: "TextBlock",
              text: "Now you have triggered a command that sends this card! Go to documentations to learn more about Adaptive Card and Commands in Teams Bot. Click on \"I like this\" below if you think this is helpful.",
              wrap: true
            },
            {
              type: "FactSet",
              facts: [
                {
                  title: "Like Count:",
                  value: "0"
                }
              ]
            }
          ],
          actions: [
            {
              type: "Action.Submit",
              title: "I Like This!",
              data: { action: "like" }
            },
            {
              type: "Action.OpenUrl",
              title: "Adaptive Card Docs",
              url: "https://learn.microsoft.com/en-us/adaptive-cards/"
            },
            {
              type: "Action.OpenUrl",
              title: "Bot Command Docs",
              url: "https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/create-a-bot-commands-menu"
            }
          ]
        });

        await context.sendActivity({ attachments: [card] });
      } else {
        try {
          // 使用 httpbin.org 的 echo API
          const response = await fetch('https://httpbin.org/post', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ message: txt })
          });
          
          const data = await response.json();
          await context.sendActivity(`API Response: ${data.json.message}`);
        } catch (error) {
          console.error('Error calling echo API:', error);
          await context.sendActivity('Sorry, there was an error processing your request.');
        }
      }
      
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(
            `Hi there! I'm a Teams bot that will echo what you said to me.`
          );
          break;
        }
      }
      await next();
    });
  }
}

module.exports.TeamsBot = TeamsBot;
