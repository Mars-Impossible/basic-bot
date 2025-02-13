const { TeamsActivityHandler, TurnContext, CardFactory, ActivityTypes } = require("botbuilder");
const { ConversationState, MemoryStorage } = require('botbuilder');
const { chatWithSSE } = require('./api/chat');
const fetch = require('node-fetch');
const { contextSearch, keySearch, queryContactList, queryAccountList, queryFundList, queryActivityList, queryDocumentList } = require('./api/search'); 
const  aiChatConfig  = require('./store/aiChatConfig');
const { 
  createContactCard,
  createAccountCard,
  createFundCard,
  createActivityCard,
  createDocumentCard,
  createErrorCard,
  buildDetailUrl
} = require('./ui/searchCard');

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    // 设置存储和状态管理
    this.storage = new MemoryStorage(); //不适合生产环境
    this.conversationState = new ConversationState(this.storage);
    this.historyAccessor = this.conversationState.createProperty('history');

    // 添加通用的 suggestedActions
    this.commonSuggestedActions = {
      actions: [
        {
          type: 'imBack',
          title: '👍 Helpful',
          value: 'helpful'
        },
        {
          type: 'imBack',
          title: '👎 Not Helpful',
          value: 'not helpful'
        },
        {
          type: 'imBack',
          title: '🔄 Regenerate',
          value: 'regenerate'
        }
      ]
    };

    // 创建一个帮助方法来添加 suggestedActions
    this.createActivityWithSuggestions = (content) => {
      return {
        ...content,
        suggestedActions: this.commonSuggestedActions
      };
    };

    this.onMessage(async (context, next) => {
      // 获取用户信息
      const userInfo = {
        id: context.activity.from.id,
        name: context.activity.from.name,
        aadObjectId: context.activity.from.aadObjectId, // Azure AD Object ID
        userPrincipalName: context.activity.from.userPrincipalName // 如果可用
      };
      
      // 注释掉历史记录获取
      // const history = await this.historyAccessor.get(context, []);
      
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText ? removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim() : "";

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

        await context.sendActivity(this.createActivityWithSuggestions({ 
          attachments: [card]
        }));
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

        await context.sendActivity(this.createActivityWithSuggestions({ 
          attachments: [heroCard] 
        }));
      } else if (txt === "citation") {
        await context.sendActivity(this.createActivityWithSuggestions({
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
        }));
      } else if (txt === "/search") {
        // 初始搜索卡片
        const searchCard = CardFactory.adaptiveCard({
          type: "AdaptiveCard",
          version: "1.0",
          body: [
            {
              type: "TextBlock",
              text: "AI Search",
              weight: "bolder",
              size: "medium"
            },
            {
              type: "Input.Text",
              id: "searchQuery",
              placeholder: "Enter your search query",
              isRequired: true
            }
          ],
          actions: [
            {
              type: "Action.Submit",
              title: "Search",
              data: { action: "aiSearch" }
            }
          ]
        });

        await context.sendActivity(this.createActivityWithSuggestions({ 
          attachments: [searchCard] 
        }));
      } else if (context.activity.value && context.activity.value.action === "aiSearch") {
        // 处理搜索提交
        const query = context.activity.value.searchQuery;
        try {
          const results = await contextSearch(query);
          
          // 将结果按类型分组
          const groupedResults = results.reduce((acc, result) => {
            const type = result.targetType;
            if (!acc[type]) acc[type] = [];
            acc[type].push(result);
            return acc;
          }, {});

          // 创建结果卡片
          const resultCard = CardFactory.adaptiveCard({
            type: "AdaptiveCard",
            version: "1.0",
            body: [
              {
                type: "TextBlock",
                text: "Search Results",
                weight: "bolder",
                size: "medium"
              }
            ],
            actions: Object.entries(groupedResults).map(([type, items]) => ({
              type: "Action.ShowCard",
              title: `${aiChatConfig.targetTypes.find(t => t.id === parseInt(type))?.name || 'Unknown'} (${items.length})`,
              card: {
                type: "AdaptiveCard",
                body: items.map(item => ({
                  type: "TextBlock",
                  text: item.name || item.title,
                  wrap: true
                }))
              }
            }))
          });

          await context.sendActivity(this.createActivityWithSuggestions({ 
            attachments: [resultCard]
          }));
        } catch (error) {
          console.error('Search error:', error);
          await context.sendActivity('Sorry, there was an error processing your search.');
        }
      } else {
        // 注释掉历史记录更新
        // history.push({
        //   role: 'user',
        //   content: txt,
        //   userInfo: userInfo,
        //   timestamp: new Date().toISOString()
        // });
        
        // 注释掉历史记录长度控制
        // if (history.length > 3) {
        //   history.shift();
        // }
        // await this.historyAccessor.set(context, history);

        // 发送 typing 状态
        await context.sendActivity({ type: 'typing' });

        // 发送初始响应
        const initialResponse = await context.sendActivity({
          type: 'message',
          text: 'Thinking...'
        });

        let responseText = '';

        await chatWithSSE({
          message: txt,
          onUpdate: async (text) => {
            try {
              await context.updateActivity(this.createActivityWithSuggestions({
                id: initialResponse.id,
                type: 'message',
                text: text
              }));
            } catch (error) {
              console.error('Update error:', error);
            }
          },
          onFinish: async (finalText) => {
            try {
              await context.updateActivity(this.createActivityWithSuggestions({
                id: initialResponse.id,
                type: 'message',
                text: finalText
              }));
            } catch (error) {
              console.error('Final update error:', error);
            }
          },
          onError: async (error) => {
            console.error('Chat error:', error);
            await context.updateActivity({
              id: initialResponse.id,
              type: 'message',
              text: 'Sorry, there was an error processing your request.'
            });
          }
        });

        await next();
      }
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(this.createActivityWithSuggestions({
            text: `Hi there! I'm an ai assistant for you. You can ask me anything.`
          }));
          break;
        }
      }
      await next();
    });
  }

  // 处理搜索查询
  async handleTeamsMessagingExtensionQuery(context, query) {
    const searchQuery = query.parameters[0].value?.trim();

    if (!searchQuery || searchQuery.length < 3) {
      return null;
    }

    try {
      // 根据命令ID选择搜索方法
      const searchMethod = query.commandId === 'aiSearch' ? contextSearch : keySearch;
      console.log('Search method:', query.commandId);
      const startTime = Date.now();
      console.log(`Start time: ${startTime}`);

      const results = await searchMethod(searchQuery);
      console.log('Search results:', results);  
      const endTime = Date.now();
      console.log(`Search time: ${endTime - startTime}ms`);

      // 处理搜索结果
      const attachments = results.map(result => {
        // 获取 targetType 对应的名称
        console.log('result.targetType:', result.targetType);
        console.log('aiChatConfig.targetTypes:', aiChatConfig.targetTypes);
        const targetTypeName = aiChatConfig.targetTypes.find(t => t.id === result.targetType)?.name || 'Unknown Type';

        const heroCard = CardFactory.heroCard(
          result.name || result.title || 'No Title',
          targetTypeName  // 添加 subtitle
        );

        const preview = CardFactory.heroCard(
          result.name || result.title || 'No Title',
          targetTypeName  // preview 也添加 subtitle
        );
        preview.content.tap = { type: 'invoke', value: result };

        return { ...heroCard, preview };
      });

      return {
        composeExtension: {
          type: 'result',
          attachmentLayout: 'list',
          attachments: attachments
        }
      };
    } catch (error) {
      console.error('Search error:', error);  // 已有的错误日志
      throw error;
    }
  }

  // 处理选中项，生成与原代码相同的 Adaptive Card
  async handleTeamsMessagingExtensionSelectItem(context, obj) {
    console.log('Selected item data:', obj);
    try {
      let data;
      const detailUrl = buildDetailUrl(obj);  
      
      switch (obj.targetType) {
        case 1: // Account
        data = await queryAccountList({
          tagMappingMenuId: obj.tagMenuId,
          keywords: obj.relatedId
        });
        return createAccountCard(data[0], detailUrl);

        case 2: // Contact
          data = await queryContactList({
            tagMappingMenuId: obj.tagMenuId,
            keywords: obj.relatedId
          });
          return createContactCard(data[0], detailUrl);
          
        case 3: // Fund
          data = await queryFundList({
            tagMappingMenuId: obj.tagMenuId,
            keywords: obj.relatedId
          });
          return createFundCard(data[0], detailUrl);
          
        case 4: // Activity
          data = await queryActivityList({
            tagMappingMenuId: obj.tagMenuId,
            keywords: obj.relatedId
          });
          return createActivityCard(data[0], detailUrl);
          
        case 5: // Document
          data = await queryDocumentList({
            tagMappingMenuId: obj.tagMenuId,
            keywords: obj.relatedId
          });
          return createDocumentCard(data[0], detailUrl);
          
        default:
          throw new Error(`Unsupported target type: ${obj.targetType}`);
      }
    } catch (error) {
      console.error('Error in handleTeamsMessagingExtensionSelectItem:', error);
      return createErrorCard('Failed to load details');
    }
  }
}

module.exports.TeamsBot = TeamsBot;
