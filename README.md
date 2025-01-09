# Teams Bot 

## setup bot
- 在项目根目录下,运行
```
npm install
```
- 打开侧边栏,选择 debug in Teams
- 在teams中选择添加到channel或者group chat中。

## 基础命令
- 在groupchat和channel中，以下命令均需要`@basic-botlocal`
- `learn` - 显示自适应卡片(Adaptive Card)教程,包含详细说明和文档链接
- `news` - 显示西雅图单轨列车的Adaptive Card,包含图片和相关链接
- `citation` - 发送一条带有引用链接的消息示例
- `delete history` - 清除当前对话的历史记录

## 消息历史
- Bot会自动记录最近3条对话内容
- 发送任何不属于以上命令的消息时,Bot会显示最近的对话历史

## Message Extension
Bot支持消息扩展搜索功能,可以搜索并分享文档(目前显示示例数据)。

