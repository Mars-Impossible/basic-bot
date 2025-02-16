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
} = require('./ui/searchResultCard');
const { createSearchCard } = require('./ui/searchQueryCard');

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
            }
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
        const searchCard = createSearchCard();
        await context.sendActivity(this.createActivityWithSuggestions({ 
          attachments: [searchCard] 
        }));
      } else if (context.activity.value && context.activity.value.action === "aiSearch") {
        const query = context.activity.value.searchQuery;
        const isAISearch = context.activity.value.searchMode === "true";
        const selectedTypes = context.activity.value.searchTypes ? context.activity.value.searchTypes.split(',') : [];

        // 验证是否选择了至少一个类型
        if (selectedTypes.length === 0) {
          const errorCard = CardFactory.adaptiveCard({
            type: "AdaptiveCard",
            version: "1.0",
            body: [
              {
                type: "TextBlock",
                text: "Error",
                weight: "bolder",
                size: "medium",
                color: "attention"
              },
              {
                type: "TextBlock",
                text: "Please select at least one type to search",
                wrap: true
              }
            ]
          });
          await context.sendActivity({ attachments: [errorCard] });
          return;
        }
        
        try {
          const modulesFilterStr = selectedTypes
            .map(type => `TargetTypes=${type}`)
            .join('&');
          
          const searchFunction = isAISearch ? contextSearch : keySearch;
          const results = await searchFunction(query, modulesFilterStr);
          
          // 添加调试日志
          console.log('Search results:', results);
          
          // 修改结果分组，使用正确的字段名
          const groupedResults = results.reduce((acc, result) => {
            if (selectedTypes.includes(result.targetType.toString())) {
              const type = result.targetType;
              if (!acc[type]) acc[type] = [];
              acc[type].push({
                id: result.relatedId,  // 使用 relatedId 而不是 id
                name: result.name || result.title,
                title: result.title,
                targetType: result.targetType,
                tagMappingMenuId: result.tagMappingMenuId  // 确保使用正确的字段名
              });
            }
            return acc;
          }, {});

          // 添加调试日志
          console.log('Grouped results:', groupedResults);
          
          const updatedCard = createSearchCard(query, isAISearch, selectedTypes.join(','));
          
          updatedCard.content.actions = Object.entries(groupedResults).map(([type, items]) => {
            console.log(`Processing type ${type} items:`, items); // 调试日志
            return {
              type: "Action.ShowCard",
              title: `${aiChatConfig.targetTypes.find(t => t.id === parseInt(type))?.name || 'Unknown'} (${items.length})`,
              card: {
                type: "AdaptiveCard",
                body: items.map(item => {
                  const detailUrl = buildDetailUrl({
                    targetType: parseInt(type),
                    relatedId: item.id,  // 这里使用的是上面保存的 relatedId
                    tagMenuId: item.tagMappingMenuId
                  });
                  
                  // 添加调试日志
                  console.log('Building URL for item:', {
                    type,
                    relatedId: item.id,
                    tagMenuId: item.tagMappingMenuId,
                    url: detailUrl
                  });

                  return {
                    type: "Container",
                    selectAction: {
                      type: "Action.OpenUrl",
                      url: detailUrl
                    },
                    items: [
                      {
                        type: "TextBlock",
                        text: item.name || item.title,
                        wrap: true,
                        color: "accent"
                      }
                    ]
                  };
                })
              }
            };
          });

          await context.updateActivity({
            type: 'message',
            id: context.activity.replyToId,
            attachments: [updatedCard]
          });

        } catch (error) {
          console.error('Search error:', error);
          await context.sendActivity('Sorry, there was an error processing your search.');
        }
      } else {

        // 发送 typing 状态
        await context.sendActivity({ type: 'typing' });

        // 发送初始响应
        const initialResponse = await context.sendActivity({
          type: 'message',
          text: '...',
          suggestedActions: this.commonSuggestedActions
        });

        let lastUpdateTime = 0;  
        const updateInterval = 200;  

        await chatWithSSE({
          message: txt,
          onUpdate: async (text) => {
            try {
              const now = Date.now();
              if (now - lastUpdateTime >= updateInterval) {
                await context.updateActivity({
                  id: initialResponse.id,
                  type: 'message',
                  text: text
                });
                lastUpdateTime = now;
              }
            } catch (error) {
              console.error('Update error:', error);
            }
          },
          onFinish: async (finalText) => {
            try {
              console.log('[Teams] Finishing with final text length:', finalText.length);
              
              await context.updateActivity({
                id: initialResponse.id,
                type: 'message',
                text: finalText,
              });
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
