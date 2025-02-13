const request = require('./request');
const { fetchEventSource } = require('@fortaine/fetch-event-source');

// 直接设置 token
const AUTH_TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJodHRwOi8vc2NoZW1hcy54bWxzb2FwLm9yZy93cy8yMDA1LzA1L2lkZW50aXR5L2NsYWltcy9uYW1laWRlbnRpZmllciI6IjkyMzY0Yjc1LWNhNjYtNDc4NC04MTlmLWU5ODRkM2ZjYThhYyIsImh0dHA6Ly9zY2hlbWFzLnhtbHNvYXAub3JnL3dzLzIwMDUvMDUvaWRlbnRpdHkvY2xhaW1zL25hbWUiOiJqb2hubnkgd2FuZyIsIm5iZiI6MTczNjkzMjcxOCwiZXhwIjoxNzM3NTM3NTE4LCJpc3MiOiJodHRwOi8vbG9jYWxob3N0OjUwMDAiLCJhdWQiOiJodHRwOi8vbG9jYWxob3N0OjUwMDAifQ.zowTdszEnKrZw3JeVb8QRwYuxRDathEDmrBGc-EdRSc";

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

async function chatWithSSE({ message, onUpdate, onFinish, onError }) {
    try {
        const requestData = {
            chatSessionId: '8971d9ad-e81e-4ca9-9e16-ae0dfce2e444',
            message: message,
            ImageUrls: []
        };

        let responseText = '';
        let finished = false;

        const throttledUpdate = throttle((text) => {
            onUpdate?.(text);
        },400); 

        console.log('[SSE] Starting chat with message:', message);

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

module.exports = {
    chatWithSSE
};