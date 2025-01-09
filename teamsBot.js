const { TeamsActivityHandler, TurnContext, CardFactory, ActivityTypes } = require("botbuilder");
const { ConversationState, MemoryStorage } = require('botbuilder');

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    // 设置存储和状态管理
    this.storage = new MemoryStorage(); //不适合生产环境
    this.conversationState = new ConversationState(this.storage);
    this.historyAccessor = this.conversationState.createProperty('history');

    this.onMessage(async (context, next) => {
      // 获取当前对话的历史记录
      const history = await this.historyAccessor.get(context, []);
      
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
      } else if (txt === "citation") {
        await context.sendActivity({
          type: ActivityTypes.Message,
          text: "This is a message with citations [1][2]. You can click on the citation numbers to view more information.",
          entities: [{
            type: "https://schema.org/Message",
            "@type": "Message",
            "@context": "https://schema.org",
            // additionalType: ["AIGeneratedContent"],  // 启用 AI 标签
            citation: [
              {
                "@type": "Claim",
                position: 1,
                appearance: {
                  "@type": "DigitalDocument",
                  name: "Teams Bot documentation",
                  url: "https://learn.microsoft.com/en-us/microsoftteams/platform/bots/design/bots",
                  abstract: "Teams Bot documentation",
                  keywords: ["Teams", "Bot", "documentation"]
                }
              },
              {
                "@type": "Claim",
                position: 2,
                appearance: {
                  "@type": "DigitalDocument",
                  name: "Citation example",
                  url: "https://example.com/citation",
                  abstract: "This is the detailed explanation of the second citation",
                  keywords: ["example", "citation", "Teams"]
                }
              }
            ]
          }],
          channelData: {
            feedbackLoopEnabled: true  // 启用反馈按钮
          }
        });
      } else if (txt === "delete history") {
        await context.sendActivities([
          { type: 'typing' },
          { type: 'message', text: 'History cleared!' }
        ]);
        await this.historyAccessor.set(context, []);
      } else {
        // 添加新消息到历史记录
        history.push({
          role: 'user',
          content: txt
        });
        
        // 保持最多3条记录
        if (history.length > 3) {
          history.shift();
        }
        
        // 保存更新后的历史记录
        await this.historyAccessor.set(context, history);

        try {
          const response = await fetch('https://httpbin.org/post', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify({ 
              message: txt,
              history: history 
            })
          });
          
          const data = await response.json();
          let historyMessage = 'History:\n';
          history.forEach((msg, index) => {
            historyMessage += `${index + 1}. ${msg.content}\n`;
          });
          
          // 发送typing状态和消息
          await context.sendActivities([
            { type: 'typing' },
            { type: 'message', text: historyMessage }
          ]);
          
        } catch (error) {
          console.error('Error:', error);
          await context.sendActivities([
            { type: 'typing' },
            { type: 'message', text: 'Sorry, there was an error.' }
          ]);
        }
      }
      
      // 保存状态
      await this.conversationState.saveChanges(context);
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
