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
      // 获取会话相关的唯一标识符
      const conversationId = context.activity.conversation.id;  // 完整的会话ID
      const conversationProperties = {
        conversationId: context.activity.conversation.id,       // 会话ID
        channelId: context.activity.channelId,                 // 通道ID (例如: 'msteams')
        tenantId: context.activity.conversation.tenantId,      // Teams 租户ID
        userId: context.activity.from.id,                      // 用户ID
        aadObjectId: context.activity.from.aadObjectId,        // Azure AD 对象ID
      };

      console.log('Conversation Properties:', conversationProperties);
      
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
          
          
          // 修改结果分组，使用正确的字段名
          const groupedResults = results.reduce((acc, result) => {
            if (selectedTypes.includes(result.targetType.toString())) {
              const type = result.targetType;
              if (!acc[type]) acc[type] = [];
              const maxResults = context.activity.value.maxResultCount || 5; // 从 Input.Number 获取值

              if (acc[type].length < maxResults) {  // 使用动态的限制数量
                // 清理文本格式的函数
                const cleanFormatting = (text) => {
                  return text
                    .replace(/[""]/g, '') // 移除双引号
                    .replace(/\*\*/g, '') // 移除markdown加粗
                    .replace(/\[|\]/g, '') // 移除方括号
                    .replace(/\(.*?\)/g, '') // 移除括号及其内容
                    .trim();
                };

                const text = cleanFormatting(result.text || '');
                let truncatedText = '';
                
                if (text.length > 100) {
                  // 将文本分割成单词
                  const words = text.split(' ');
                  let currentLength = 0;
                  
                  // 逐个添加单词，直到接近但不超过100个字符
                  for (const word of words) {
                    if (currentLength + word.length + 1 <= 97) { // 97 留出空间给 "..."
                      truncatedText += (truncatedText ? ' ' : '') + word;
                      currentLength += word.length + (truncatedText ? 1 : 0);
                    } else {
                      break;
                    }
                  }
                  truncatedText += ' ...';
                } else {
                  truncatedText = text;
                }
                
                const truncateName = (text) => {
                  if (!text) return '';
                  
                  // 清理特殊格式
                  let cleanText = text
                    .replace(/[""]/g, '') // 移除双引号
                    .replace(/\*\*/g, '') // 移除markdown加粗
                    .replace(/\[|\]/g, '') // 移除方括号
                    .replace(/\(.*?\)/g, '') // 移除括号及其内容
                    .trim();
                  
                  if (cleanText.length <= 60) return cleanText;
                  
                  const words = cleanText.split(' ');
                  let truncatedName = '';
                  let currentLength = 0;
                  
                  for (const word of words) {
                    if (currentLength + word.length + 1 <= 57) { // 57 留出空间给 "..."
                      truncatedName += (truncatedName ? ' ' : '') + word;
                      currentLength += word.length + (truncatedName ? 1 : 0);
                    } else {
                      break;
                    }
                  }
                  return truncatedName + ' ...';
                };

                acc[type].push({
                  id: result.relatedId,
                  name: truncateName(result.name || result.title),
                  title: result.title,
                  percentage: result.percentage || 0,
                  text: truncatedText,
                  targetType: result.targetType,
                  tagMappingMenuId: result.tagMappingMenuId
                });
              }
            }
            return acc;
          }, {});

          
          const updatedCard = createSearchCard(query, isAISearch, selectedTypes.join(','));
          
          updatedCard.content.actions = Object.entries(groupedResults).map(([type, items]) => {
            return {
              type: "Action.ShowCard",
              title: `${aiChatConfig.targetTypes.find(t => t.id === parseInt(type))?.name || 'Unknown'} (${items.length})`,
              style: "default",
              card: {
                type: "AdaptiveCard",
                body: [
                  {
                    type: "TextBlock",
                    text: "Search Results",
                    weight: "bolder",
                    size: "medium",
                    spacing: "medium"
                  },
                  {
                    type: "ActionSet",
                    actions: items.map(item => {
                      const detailUrl = buildDetailUrl({
                        targetType: parseInt(type),
                        relatedId: item.id,  // 这里使用的是上面保存的 relatedId
                        tagMenuId: item.tagMappingMenuId
                      });
                      return {
                        type: "Action.OpenUrl",
                        title: `${Math.round((item.percentage || 0) * 100).toString().padStart(2, ' ')}% | ${item.name || item.title}`,
                        url: detailUrl,
                        tooltip: `${item.text}\n${'─'.repeat(40)}`
                      };
                    })
                  }
                ]
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

    // 监听对话更新事件
    // this.onConversationUpdate(async (context, next) => {
    //   // 检查是否是对话删除事件
    //   if (context.activity.channelData?.eventType === 'teamChatDeleted') {
    //     const conversationId = context.activity.conversation.id;
        
    //     console.log('Chat deletion detected:', {
    //       conversationId: conversationId,
    //       userId: context.activity.from.id,
    //       aadObjectId: context.activity.from.aadObjectId,
    //       timestamp: new Date().toISOString()
    //     });

    //     try {
    //       // 1. 从数据库中获取映射关系
    //       const mapping = await this.getConversationMapping(conversationId);
          
    //       if (mapping) {
    //         // 2. 删除公司系统中的聊天记录
    //         await this.deleteCompanySystemChat(mapping.companySessionId);
            
    //         // 3. 删除或标记映射关系为已删除
    //         await this.deleteConversationMapping(conversationId);
            
    //         console.log('Successfully deleted associated data:', {
    //           teamsConversationId: conversationId,
    //           companySessionId: mapping.companySessionId
    //         });
    //       }
    //     } catch (error) {
    //       console.error('Error handling conversation deletion:', error);
    //     }
    //   }
      
    //   await next();
    // });
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
      const startTime = Date.now();

      const results = await searchMethod(searchQuery);
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
