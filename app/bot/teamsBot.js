const { TeamsActivityHandler, TurnContext, CardFactory, ActivityTypes } = require("botbuilder");
const { ConversationState, MemoryStorage } = require('botbuilder');
const { contextSearch } = require('../api/search');  // 引入 contextSearch

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
            }
            // {
            //   type: "FactSet",
            //   facts: [
            //     {
            //       title: "Like Count:",
            //       value: "0"
            //     }
            //   ]
            // }
          ],
          actions: [
            {
              type: "Action.ShowCard",
              title: "Show Details",
              card: {
                type: "AdaptiveCard",
                body: [
                  {
                    type: "TextBlock",
                    text: "Additional information here"
                  },
                  {
                    type: "Input.Text",
                    id: "comment",
                    placeholder: "Add your comment"
                  }
                ],
                actions: [
                  {
                    type: "Action.Submit",
                    title: "Submit Comment",
                    data: { action: "submitComment" }
                  }
                ]
              }
            },
            {
              type: "Action.ToggleVisibility",
              title: "Show/Hide Details",
              targetElements: ["detailsSection"]
            },
            {
              type: "Action.OpenUrl",
              title: "Adaptive Card Docs",
              url: "https://learn.microsoft.com/en-us/adaptive-cards/"
            },
          //   {
          //     // Teams特有的执行动作
          //     type: "Action.Execute",
          //     title: "Echo History",
          //     verb: "processData",  // 后端处理时的标识符
          //     data: {
          //         operationType: "analyze",
          //         parameters: {
          //             source: "userAction"
          //         }
          //     }
          // },
          // {
          //     // 打开任务模块
          //     type: "Action.Submit",
          //     title: "Open Task Module",
          //     data: {
          //         msteams: {
          //             type: "task/fetch",
          //             taskModule: {
          //                 title: "Task Module",
          //                 height: "medium",
          //                 width: "medium"
          //             }
          //         }
          //     }
          // }
          ]
        });

        await context.sendActivity({ attachments: [card] });
      } else if (txt === "news") {
        const heroCard = CardFactory.heroCard(
          'Seattle Center Monorail',
          'Seattle Center Monorail',
          [{
            url: 'https://upload.wikimedia.org/wikipedia/commons/thumb/4/49/Seattle_monorail01_2008-02-25.jpg/1024px-Seattle_monorail01_2008-02-25.jpg'
          }],
          [{
            type: 'openUrl',
            title: 'Official website',
            value: 'https://www.seattlemonorail.com'
          },
          {
            type: 'openUrl',
            title: 'Wikipeda page',
            value: 'https://en.wikipedia.org/wiki/Seattle_Center_Monorail'
          }],
          {
            text: 'The Seattle Center Monorail is an elevated train line between Seattle Center (near the Space Needle) and downtown Seattle. It was built for the 1962 World\'s Fair. Its original two trains, completed in 1961, are still in service.'
          }
        );

        await context.sendActivity({ attachments: [heroCard] });
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

  // 处理搜索查询
  async handleTeamsMessagingExtensionQuery(context, query) {
    const searchQuery = query.parameters[0].value;
    try {
      // 使用 contextSearch 获取搜索结果
      const results = await contextSearch(searchQuery);
      
      // 确保返回的是数组
      if (!Array.isArray(results)) {
        console.log('Invalid response format:', results);
        throw new Error('Invalid response format');
      }

      return {
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: results.map(item => ({
            contentType: 'application/vnd.microsoft.card.adaptive',
            content: {
              type: 'AdaptiveCard',
              version: '1.0',
              body: [
                {
                  type: 'TextBlock',
                  text: item.title || item.name,
                  weight: 'bolder',
                  size: 'medium'
                },
                {
                  type: 'TextBlock',
                  text: `Type: ${item.targetType}`,
                  wrap: true
                },
                {
                  type: 'FactSet',
                  facts: [
                    {
                      title: 'ID:',
                      value: item.relatedId
                    },
                    {
                      title: 'Menu ID:',
                      value: item.tagMenuId
                    }
                  ]
                }
              ],
              actions: [
                {
                  type: 'Action.Submit',
                  title: 'Share',
                  data: {
                    msteams: {
                      type: 'messageBack',
                      text: `Sharing: ${item.title || item.name}`
                    },
                    itemId: item.relatedId
                  }
                }
              ]
            }
          }))
        }
      };
    } catch (error) {
      console.error('Search error:', error);
      return {
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: []
        }
      };
    }
  }
}

module.exports.TeamsBot = TeamsBot;
