const { TeamsActivityHandler, TurnContext, CardFactory, ActivityTypes } = require("botbuilder");
const { ConversationState, MemoryStorage } = require('botbuilder');
const { chatWithSSE, deleteHistory } = require('./api/chat');
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
const NodeCache = require('node-cache');
const axios = require('axios');

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

    // 添加内存缓存
    this.mappingCache = new NodeCache({ 
      stdTTL: 3600, // 1小时过期
      checkperiod: 600 // 每10分钟检查过期
    });

    this.onMessage(async (context, next) => {
      // 判断聊天类型
      // const isGroupChat = context.activity.conversation.conversationType === 'channel';
      // const chatType = isGroupChat ? 'group' : 'personal';
      // console.log('Chat Type:', chatType);
      console.log('Conversation Type:', context.activity.conversation.conversationType);
      
      // 获取会话相关的唯一标识符
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

      // 检查是否是单独的 @command
      const soloCommands = {
        '@accounts': { 
          usage: '<span style="color: #D4AF37;">💡 Usage: @accounts your question ...</span>',
          type: 1  
        },
        '@contacts': { 
          usage: '<span style="color: #D4AF37;">💡 Usage: @contacts your question ...</span>',
          type: 2  
        },
        '@funds': {   // 修改这里
          usage: '<span style="color: #D4AF37;">💡 Usage: @funds your question ...</span>',
          type: 3  
        },
        '@activities': {   // 修改这里
          usage: '<span style="color: #D4AF37;">💡 Usage: @activities your question ...</span>',
          type: 4  
        },
        '@documents': { 
          usage: '<span style="color: #D4AF37;">💡 Usage: @documents your question ...</span>',
          type: 5  
        }
      };

      // 检查命令类型
      const commandMatch = Object.keys(soloCommands).find(cmd => txt.startsWith(cmd));
      if (commandMatch) {
        const searchTerm = txt.slice(commandMatch.length).trim();
        
        // 如果是单独的命令，显示使用说明
        if (!searchTerm) {
          await context.sendActivity(this.createActivityWithSuggestions({ 
            text: soloCommands[commandMatch].usage,
            textFormat: 'xml'
          }));
          return;
        }

        // 有搜索词，执行搜索
        try {
          const targetType = soloCommands[commandMatch].type;
          const modulesFilterStr = `TargetTypes=${targetType}`;
          const results = await contextSearch(searchTerm, modulesFilterStr);
          
          // 检查是否有结果
          if (!results || results.length === 0) {
            await context.sendActivity(this.createActivityWithSuggestions({ 
              text: `<span style="color: #D4AF37;">No results found for "${searchTerm}" in ${commandMatch.slice(1)}</span>`,
              textFormat: 'xml'  // 启用 HTML 格式化
            }));
            return;
          }

          // 文本清理函数
          const cleanFormatting = (text) => {
            return text
              .replace(/[""]/g, '') // 移除双引号
              .replace(/\*\*/g, '') // 移除markdown加粗
              .replace(/\[|\]/g, '') // 移除方括号
              .replace(/\(.*?\)/g, '') // 移除括号及其内容
              .replace(/^\d+\.\s+/, '') // 移除开头的数字编号（如 "6. "）
              .replace(/\s+/g, ' ') // 将多个空格替换为单个空格
              .trim();
          };

          // 名称截断函数
          const truncateName = (text) => {
            if (!text) return '';
            let cleanText = cleanFormatting(text);
            if (cleanText.length <= 57) return cleanText;
            
            const words = cleanText.split(' ');
            let truncatedName = '';
            let currentLength = 0;
            
            for (const word of words) {
              if (currentLength + word.length + 1 <= 53) {
                truncatedName += (truncatedName ? ' ' : '') + word;
                currentLength += word.length + (truncatedName ? 1 : 0);
              } else {
                break;
              }
            }
            return truncatedName + ' ...';
          };

          // 创建结果卡片
          const resultCard = CardFactory.adaptiveCard({
            type: "AdaptiveCard",
            version: "1.4",
            body: [
              {
                type: "Container",
                items: results.map(item => {
                  // 处理描述文本
                  const text = cleanFormatting(item.text || '');
                  let truncatedText = '';
                  
                  if (text.length > 70) {
                    const words = text.split(' ');
                    let currentLength = 0;
                    
                    for (const word of words) {
                      if (currentLength + word.length + 1 <= 67) {
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

                  return {
                    type: "Container",
                    style: "emphasis",  
                    selectAction: {
                      type: "Action.OpenUrl",
                      url: buildDetailUrl({
                        targetType: targetType,
                        relatedId: item.relatedId,
                        tagMenuId: item.tagMappingMenuId
                      }),
                      tooltip: `${truncatedText}\n${'─'.repeat(40)}`
                    },
                    items: [
                      {
                        type: "TextBlock",
                        text: `${Math.round((item.percentage || 0) * 100).toString().padStart(2, ' ')}% | ${truncateName(item.name || item.title)}`,
                        wrap: true,
                        size: "medium",
                        weight: "bolder",
                        spacing: "small" 
                      },
                      {
                        type: "TextBlock",
                        text: truncatedText,
                        wrap: true,
                        size: "small",
                        color: "light",
                        spacing: "small"
                      }
                    ],
                    spacing: "small", 
                    padding: "default"  
                  };
                })
              }
            ]
          });

          await context.sendActivity({ attachments: [resultCard] });
        } catch (error) {
          console.error('Quick search error:', error);
          await context.sendActivity(this.createActivityWithSuggestions({ 
            text: `❗ <span style="color: #FF4444;">Sorry, there was an error processing your search.</span>`,
            textFormat: 'xml'
          }));
        }
        return;
      }

      // 添加对 help 命令的处理
      if (txt === "/help" || txt === "/?") {
        const helpCard = CardFactory.adaptiveCard({
          type: "AdaptiveCard",
          version: "1.4",
          style: "default",
          backgroundImage: {
            url: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSI0MCIgaGVpZ2h0PSI0MCIgdmlld0JveD0iMCAwIDQwIDQwIj48ZyBmaWxsPSJub25lIiBmaWxsLXJ1bGU9ImV2ZW5vZGQiPjxwYXRoIGZpbGw9IiMxYjFiMWIiIGQ9Ik0wIDBoNDB2NDBIMHoiLz48cGF0aCBkPSJNMCAwaDQwdjQwSDB6IiBmaWxsPSIjMjEyMTIxIiBmaWxsLW9wYWNpdHk9Ii44Ii8+PC9nPjwvc3ZnPg==",
            fillMode: "repeat"
          },
          body: [
            {
              type: "Container",
              style: "emphasis",
              items: [
                {
                  type: "TextBlock",
                  text: "Available Commands",
                  size: "large",
                  weight: "bolder",
                  color: "light",
                  horizontalAlignment: "center",
                  spacing: "medium"
                },
                {
                  type: "Container",
                  style: "default",
                  items: [
                    {
                      type: "FactSet",
                      facts: [
                        {
                          title: "**`/help`** or **`/?`**",
                          value: "Display this help message"
                        },
                        {
                          title: "**`/search`**",
                          value: "Open the advanced search interface"
                        },
                        {
                          title: "**`/clear`**",
                          value: "Clear current conversation history"
                        },
                        {
                          title: "**`@accounts`**",
                          value: "Quick search for accounts"
                        },
                        {
                          title: "**`@contacts`**",
                          value: "Quick search for contacts"
                        },
                        {
                          title: "**`@activities`**",
                          value: "Quick search for activities"
                        },
                        {
                          title: "**`@funds`**",
                          value: "Quick search for funds"
                        },
                        {
                          title: "**`@documents`**",
                          value: "Quick search for documents"
                        }
                      ]
                    }
                  ],
                  style: "emphasis",
                  bleed: true,
                  spacing: "padding"
                }
              ]
            },
            {
              type: "Container",
              items: [
                {
                  type: "TextBlock",
                  text: "💡 **Tip**: You can use these commands anytime during our conversation",
                  wrap: true,
                  color: "accent",
                  size: "small",
                  horizontalAlignment: "center"
                }
              ],
              spacing: "medium"
            }
          ],
          "$schema": "http://adaptivecards.io/schemas/adaptive-card.json"
        });

        await context.sendActivity(this.createActivityWithSuggestions({ 
          attachments: [helpCard] 
        }));
        return;
      }

      // 添加删除历史记录的命令处理
      if (txt === "/clear") {
        try {
          const conversationContext = {
            userId: context.activity.from.id,
            userName: context.activity.from.name,
            aadObjectId: context.activity.from.aadObjectId,
            conversationId: context.activity.conversation.id
          };

          const result = await deleteHistory(conversationContext);
          
          if (result.success) {
            await context.sendActivity("Chat history has been cleared. You can start a new conversation.");
          } else {
            await context.sendActivity(result.error.message || "No chat history found.");
          }
        } catch (error) {
          console.error('Error deleting chat history:', error);
          await context.sendActivity("Failed to clear chat history. Please try again.");
        }
        return;
      }

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
        const count = parseInt(context.activity.value.maxResultCount) || 10;

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
          const results = await searchFunction(query, modulesFilterStr, count);
          
          // 修改结果分组，使用正确的字段名
          const groupedResults = results.reduce((acc, result) => {
            if (selectedTypes.includes(result.targetType.toString())) {
              const type = result.targetType;
              if (!acc[type]) acc[type] = [];

              // 清理文本格式的函数
              const cleanFormatting = (text) => {
                return text
                  .replace(/[""]/g, '') // 移除双引号
                  .replace(/\*\*/g, '') // 移除markdown加粗
                  .replace(/\[|\]/g, '') // 移除方括号
                  .replace(/\(.*?\)/g, '') // 移除括号及其内容
                  .replace(/^\d+\.\s+/, '') // 移除开头的数字编号（如 "6. "）
                  .replace(/\s+/g, ' ') // 将多个空格替换为单个空格
                  .trim();
              };

              const text = cleanFormatting(result.text || '');
              let truncatedText = '';
              
              if (text.length > 77) {  
                const words = text.split(' ');
                let currentLength = 0;
                
                for (const word of words) {
                  if (currentLength + word.length + 1 <= 74) { 
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
                
                if (cleanText.length <= 57) return cleanText;
                
                const words = cleanText.split(' ');
                let truncatedName = '';
                let currentLength = 0;
                
                for (const word of words) {
                  if (currentLength + word.length + 1 <= 53) { 
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
            return acc;
          }, {});

          // 将 groupedResults 转换为数组并按第一个项目的 percentage 排序
          const sortedGroups = Object.entries(groupedResults)
            .sort((a, b) => {
              const aFirstPercentage = a[1][0]?.percentage || 0;
              const bFirstPercentage = b[1][0]?.percentage || 0;
              return bFirstPercentage - aFirstPercentage; // 降序排序
            });

          const updatedCard = createSearchCard(query, isAISearch, selectedTypes.join(','));

          updatedCard.content.actions = sortedGroups.map(([type, items]) => {
            return {
              type: "Action.ShowCard",
              title: `${aiChatConfig.targetTypes.find(t => t.id === parseInt(type))?.name || 'Unknown'} (${items.length})`,
              style: "default", //style修改后会报错
              card: {
                type: "AdaptiveCard",
                body: items.map(item => {
                      const detailUrl = buildDetailUrl({
                        targetType: parseInt(type),
                        relatedId: item.id,
                        tagMenuId: item.tagMappingMenuId
                      });
                      return {
                        type: "Container",
                        selectAction: {
                          type: "Action.OpenUrl",
                          url: detailUrl,
                          tooltip: `${item.text}\n${'─'.repeat(40)}`
                        },
                        items: [
                          {
                            type: "TextBlock",
                            text: `${Math.round((item.percentage || 0) * 100).toString().padStart(2, ' ')}% | ${item.name || item.title}`,
                            wrap: true,
                            size: "medium",
                            weight: "bolder",
                            spacing: "none"
                          },
                          {
                            type: "TextBlock",
                            text: item.text,
                            wrap: true,
                            size: "small",
                            color: "light",
                            spacing: "none"
                          }
                        ],
                        spacing: "medium"
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
          await context.sendActivity(this.createActivityWithSuggestions({ 
            text: `❗ <span style="color: #FF4444;">Sorry, there was an error processing your search.</span>`,
            textFormat: 'xml'
          }));
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

        // 创建完整的 conversationContext 对象
        const conversationContext = {
          userId: context.activity.from.id,
          userName: context.activity.from.name,
          aadObjectId: context.activity.from.aadObjectId,
          conversationId: context.activity.conversation.id,
          activity: context.activity
        };

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
          },
          conversationContext
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

  async getConversationMapping(conversationId) {
    // 先从缓存获取
    let mapping = this.mappingCache.get(conversationId);
    
    if (!mapping) {
      // 缓存未命中，从数据库获取
      mapping = await this.getConversationMappingFromDB(conversationId);
      
      if (mapping) {
        // 写入缓存
        this.mappingCache.set(conversationId, mapping);
      }
    }
    
    return mapping;
  }

  async deleteCompanySystemChat(companySessionId) {
    const maxRetries = 3;
    let retryCount = 0;

    while (retryCount < maxRetries) {
      try {
        await this.deleteChat(companySessionId);
        return;
      } catch (error) {
        retryCount++;
        console.error(`Failed to delete chat (attempt ${retryCount}):`, error);
        
        if (retryCount === maxRetries) {
          throw new Error('Failed to delete chat after maximum retries');
        }
        
        // 等待一段时间后重试
        await new Promise(resolve => setTimeout(resolve, 1000 * retryCount));
      }
    }
  }
}

module.exports.TeamsBot = TeamsBot;
