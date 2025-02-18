const request = require('./request');
const { fetchEventSource } = require('@fortaine/fetch-event-source');
const axios = require('axios');

// 设置 mocking API 的基础 URL
const MOCK_API_BASE_URL = 'http://localhost:3978';

// 直接设置 token
const AUTH_TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJodHRwOi8vc2NoZW1hcy54bWxzb2FwLm9yZy93cy8yMDA1LzA1L2lkZW50aXR5L2NsYWltcy9uYW1laWRlbnRpZmllciI6IjVmODAxN2ZhLTViNDYtNGY1Ny1hMzFjLTA1NjkxMTg1NzE2YiIsImh0dHA6Ly9zY2hlbWFzLnhtbHNvYXAub3JnL3dzLzIwMDUvMDUvaWRlbnRpdHkvY2xhaW1zL25hbWUiOiJ4aWFvaGUiLCJuYmYiOjE3MzkxODQ5NzYsImV4cCI6MTczOTc4OTc3NiwiaXNzIjoiaHR0cDovL2xvY2FsaG9zdDo1MDAwIiwiYXVkIjoiaHR0cDovL2xvY2FsaG9zdDo1MDAwIn0.V5c3CtSufbBBnc0ht8RLNEzgYx6Psm3mzWwvFRdGnH0";

const throttle = (func, limit) => {
    let inThrottle;
    let lastResult;
    return (...args) => {
        if (!inThrottle) {
            inThrottle = true;
            lastResult = func(...args);
            setTimeout(() => inThrottle = false, limit);
        }
        return lastResult;
    }
}

// 创建新的聊天会话
async function createNewSession() {
    try {
        const session = {
            topic: "new conversation",  // 默认主题
            conversationType: 1,              // Just chat为1
        };

        const res = await request({
            url: "/ChatSessions",
            method: "POST",
            data: session
        });
        console.log("create new session", res);
        return res;
    } catch (error) {
        console.log("[Request Error] fail to create new session");
        throw error;
    }
}

// 删除聊天会话
async function deleteChatSession(sessionId) {
    try {
        await request({
            url: `/ChatSessions/${sessionId}`,
            method: "DELETE"
        });
    } catch (error) {
        console.log("[Request Error] Fail to delete chat session", error);
        throw error;
    }
}

// 修改 getOrCreateSessionId 函数
async function getOrCreateSessionId(conversationContext) {
    try {
        if (!conversationContext || !conversationContext.userId) {
            throw new Error('User ID is required');
        }

        // 1. 先尝试获取现有的映射
        const existingMapping = await axios.get(`${MOCK_API_BASE_URL}/api/teams/mapping`, {
            params: { teamsUserId: conversationContext.userId }
        });

        if (existingMapping.data.success && existingMapping.data.data) {
            console.log("existing mapping", existingMapping.data.data);
            return existingMapping.data.data.internalSessionId;
        }

        // 2. 如果不存在映射，创建新的内部会话
        const internalSessionId = await createNewSession();
        
        // 创建新的映射关系
        const mappingData = {
            teamsUserId: conversationContext.userId,
            aadObjectId: conversationContext.aadObjectId,
            internal_session_id: internalSessionId,
            userName: conversationContext.userName
        };

        console.log('Creating mapping with data:', mappingData);  // 添加日志

        await axios.post(`${MOCK_API_BASE_URL}/api/teams/session`, mappingData);

        return internalSessionId;
    } catch (error) {
        console.error('Error in getOrCreateSessionId:', error);
        throw error;
    }
}

// 修改 chatWithSSE 函数
async function chatWithSSE({ message, onUpdate, onFinish, onError, conversationContext }) {
    try {
        //Todo: 不应该每次都请求mapping，应该添加缓存
        const sessionId = await getOrCreateSessionId(conversationContext);
        
        // 更新最后活动时间和 conversationId
        await axios.put(`${MOCK_API_BASE_URL}/api/teams/session`, {
            teamsUserId: conversationContext.userId,
            teamsConversationId: conversationContext.conversationId
        });
        // const sessionId = "aa151f22-0494-4665-8d87-d6391fb06ab1";

        const requestData = {
            chatSessionId: sessionId,
            message: message,
            ImageUrls: []
        };

        let responseText = '';
        let finished = false;

        const throttledUpdate = throttle((text) => {
            onUpdate?.(text);
        },400); 

        await fetchEventSource('https://marsai.arencore.me/api/Chats/SSEChat', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${AUTH_TOKEN}`,
                'Accept': 'application/json, text/event-stream',
                'Origin': 'https://newchat.arencore.me',
                'Referer': 'https://newchat.arencore.me/',
                'X-Requested-With': 'XMLHttpRequest'
            },
            body: JSON.stringify(requestData),
            
            onopen(res) {
                if (!res.ok) {
                    console.error('[SSE] Connection failed:', res.status);
                    throw new Error(`Failed to connect: ${res.status}`);
                }
                console.log('[SSE] Connection opened:', res.status);
            },
            
            onmessage(msg) {
                if (finished) return;
                
                try {
                    const json = JSON.parse(msg.data);
                    
                    if (json.MessageType === 1 && json.Content) {
                        responseText += json.Content;
                        throttledUpdate(responseText);
                    } else if (json.MessageType === 10) {
                        finished = true;
                        onFinish?.(responseText);
                    }
                } catch (error) {
                    onError?.(error);
                }
            },
            
            onclose() {
                console.log('[SSE] Connection closed, finished:', finished);
                if (!finished) {
                    finished = true;
                    onFinish?.(responseText);
                }
            },
            
            onerror(err) {
                console.error('[SSE] Connection error:', err);
                onError?.(err);
            }
        });
    } catch (error) {
        console.error('[Chat API] Error:', error);
        onError?.(error);
    }
}

// 修改 deleteHistory 函数
async function deleteHistory(userId) {
    try {
        // 1. 获取当前的 mapping
        const mappingResponse = await axios.get(`${MOCK_API_BASE_URL}/api/teams/mapping`, {
            params: { teamsUserId: userId }
        });

        console.log('Mapping response:', mappingResponse.data);  // 添加日志

        // 如果 mapping 不存在，直接返回成功（因为已经删除了）
        if (!mappingResponse.data.success) {
            if (mappingResponse.data.error.code === 'MAPPING_NOT_FOUND') {
                return { success: true };
            }
            return {
                success: false,
                error: mappingResponse.data.error
            };
        }

        const mapping = mappingResponse.data.data;

        // 2. 删除内部会话
        await deleteChatSession(mapping.internalSessionId);
        
        // 3. 删除 mapping 关系
        const deleteResponse = await axios.delete(`${MOCK_API_BASE_URL}/api/teams/session`, {
            params: { teamsUserId: userId }
        });

        console.log('Delete response:', deleteResponse.data);  // 添加日志

        return { success: true };
    } catch (error) {
        console.error('Error in deleteHistory:', {
            message: error.message,
            response: error.response?.data
        });
        return {
            success: false,
            error: {
                code: 'DELETE_ERROR',
                message: error.message
            }
        };
    }
}

// 导出函数
module.exports = {
    chatWithSSE,
    createNewSession,
    deleteChatSession,
    deleteHistory
};