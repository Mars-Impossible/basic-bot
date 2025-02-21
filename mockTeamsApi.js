const express = require('express');
const { v4: uuidv4 } = require('uuid');
const router = express.Router();

// 设置基础 URL
const BASE_URL = 'http://localhost:3978';

// 修改数据存储结构，使用复合键 Map
const userMappingStore = new Map();

// 辅助函数：生成复合键
function generateCompositeKey(teamsUserId, aadObjectId, conversationId) {
    return `${teamsUserId}:${aadObjectId}:${conversationId}`;
}

// 辅助函数：验证 Teams 会话 ID 格式
function validateTeamsConversationId(id) {
    return typeof id === 'string' && id.length > 0;
}

// 添加请求日志中间件
router.use((req, res, next) => {
    console.log(`[${new Date().toISOString()}] ${BASE_URL}${req.path}`);
    next();
});

// 1. 获取映射关系 - 使用复合键查询
router.get('/api/teams/mapping', (req, res) => {
    try {
        const { teamsUserId, aadObjectId, conversationId } = req.query;

        if (!teamsUserId || !aadObjectId || !conversationId) {
            return res.status(400).json({
                success: false,
                error: {
                    code: 'INVALID_PARAMETERS',
                    message: 'Missing required parameters: teamsUserId, aadObjectId, and conversationId'
                }
            });
        }

        const compositeKey = generateCompositeKey(teamsUserId, aadObjectId, conversationId);
        const mapping = userMappingStore.get(compositeKey);
        
        if (!mapping || mapping.is_deleted) {
            return res.json({
                success: false,
                error: {
                    code: 'MAPPING_NOT_FOUND',
                    message: 'User mapping not found'
                }
            });
        }

        res.json({
            success: true,
            data: {
                internalSessionId: mapping.internal_session_id,
                teamsUserId: teamsUserId,
                aadObjectId: aadObjectId,
                teamsConversationId: conversationId,
                userName: mapping.user_name
            }
        });
    } catch (error) {
        console.error('Error in GET /api/teams/mapping:', error);
        res.status(500).json({ success: false, error: { code: 'INTERNAL_ERROR', message: 'Internal server error' }});
    }
});

// 2. 创建新mapping
router.post('/api/teams/session', (req, res) => {
    try {
        const {
            teamsUserId,
            userName,
            aadObjectId,
            internal_session_id,
            teamsConversationId
        } = req.body;

        // 验证必需字段
        if (!teamsUserId || !aadObjectId || !internal_session_id || !teamsConversationId) {
            return res.status(400).json({
                success: false,
                error: {
                    code: 'MISSING_REQUIRED_FIELDS',
                    message: 'Missing required fields'
                }
            });
        }

        const compositeKey = generateCompositeKey(teamsUserId, aadObjectId, teamsConversationId);

        // 检查是否存在未删除的映射
        const existingMapping = userMappingStore.get(compositeKey);
        if (existingMapping && !existingMapping.is_deleted) {
            return res.status(400).json({
                success: false,
                error: {
                    code: 'MAPPING_EXISTS',
                    message: 'Active user mapping already exists'
                }
            });
        }

        // 创建新的映射
        userMappingStore.set(compositeKey, {
            internal_session_id,
            aad_object_id: aadObjectId,
            user_name: userName,
            teams_conversation_id: teamsConversationId,
            created_at: new Date(),
            last_activity_at: new Date(),
            is_deleted: false
        });

        res.json({
            success: true,
            data: {
                internalSessionId: internal_session_id
            }
        });
    } catch (error) {
        console.error('Error in POST /api/teams/session:', error);
        res.status(500).json({ success: false, error: { code: 'INTERNAL_ERROR', message: 'Internal server error' }});
    }
});

// 3. 更新 mapping - 更新最后活动时间
router.put('/api/teams/session', (req, res) => {
    try {
        const { teamsUserId, aadObjectId, teamsConversationId } = req.body;

        if (!teamsUserId || !aadObjectId || !teamsConversationId) {
            return res.status(400).json({
                success: false,
                error: {
                    code: 'MISSING_REQUIRED_FIELDS',
                    message: 'Missing required fields'
                }
            });
        }

        const compositeKey = generateCompositeKey(teamsUserId, aadObjectId, teamsConversationId);
        const mapping = userMappingStore.get(compositeKey);
        
        if (!mapping || mapping.is_deleted) {
            return res.status(404).json({
                success: false,
                error: {
                    code: 'MAPPING_NOT_FOUND',
                    message: 'User mapping not found'
                }
            });
        }

        // 更新最后活动时间
        mapping.last_activity_at = new Date();
        userMappingStore.set(compositeKey, mapping);

        res.json({ success: true });
    } catch (error) {
        console.error('Error in PUT /api/teams/session:', error);
        res.status(500).json({ success: false, error: { code: 'INTERNAL_ERROR', message: 'Internal server error' }});
    }
});

// 4. 删除 mapping
router.delete('/api/teams/session', (req, res) => {
    try {
        const { teamsUserId, aadObjectId, conversationId } = req.query;

        if (!teamsUserId || !aadObjectId || !conversationId) {
            return res.status(400).json({
                success: false,
                error: {
                    code: 'MISSING_REQUIRED_FIELDS',
                    message: 'Missing required parameters'
                }
            });
        }

        const compositeKey = generateCompositeKey(teamsUserId, aadObjectId, conversationId);
        const mapping = userMappingStore.get(compositeKey);
        
        if (!mapping || mapping.is_deleted) {
            return res.json({
                success: false,
                error: {
                    code: 'MAPPING_NOT_FOUND',
                    message: 'Mapping not found or already deleted'
                }
            });
        }

        // 标记为已删除
        mapping.is_deleted = true;
        userMappingStore.set(compositeKey, mapping);

        res.json({ success: true });
    } catch (error) {
        console.error('Error in DELETE /api/teams/session:', error);
        res.status(500).json({ success: false, error: { code: 'INTERNAL_ERROR', message: 'Internal server error' }});
    }
});

// 添加错误处理中间件
router.use((error, req, res, next) => {
    console.error('API Error:', error);
    res.status(500).json({
        success: false,
        error: {
            code: 'INTERNAL_ERROR',
            message: 'Internal server error'
        }
    });
});

module.exports = router; 