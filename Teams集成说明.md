# Teams集成配置指南

## 🎯 功能概述

ITC报表系统现在支持向Microsoft Teams频道发送通知，可以：
- 🚨 发送紧急问题警报到Teams频道
- ⚠️ 发送常规待处理任务通知
- ✅ 发送处理完成通知（可选）
- 👥 支持@提及特定团队成员

## 📋 配置步骤

### 1. 创建Teams传入Webhook

1. **打开Teams应用**，进入要接收通知的频道
2. **点击频道名旁边的"..."** → **连接器**
3. **找到"传入Webhook"** → **配置**
4. **填写名称**（如：ITC报表通知）和**上传图标**（可选）
5. **复制生成的Webhook URL**（格式类似：`https://pg.webhook.office.com/webhookb2/...`）

### 2. 修改配置文件

编辑 `teams_config.json` 文件：

```json
{
    "enabled": true,  # 启用Teams通知
    "default_webhook": "itc_notifications",
    "webhooks": {
        "itc_notifications": "你的Webhook URL",
        "urgent_alerts": "紧急警报Webhook URL",
        "general_notifications": "一般通知Webhook URL"
    },
    "team_members": {
        "itc_team": [
            "your.name@pg.com",
            "colleague@pg.com"
        ],
        "managers": [
            "manager@pg.com"
        ]
    }
}
```

### 3. 修改主程序配置

在 `BatRun_ITCreport_downloader_rev1.py` 中调整：

```python
# Teams通知控制
TEAMS_ENABLED = True     # 启用Teams通知
TEAMS_SEND_COMPLETION = False  # 只在有问题时发送
```

## 🎨 Teams消息样式

### 紧急问题 🚨
- **红色主题**
- **标题**：🚨 ITC报表发现紧急问题
- **内容**：详细的问题统计和建议行动

### 常规待处理 ⚠️
- **橙色主题**
- **标题**：⚠️ ITC报表发现待处理任务
- **内容**：待处理项目数量和时间要求

### 无问题完成 ✅
- **蓝色主题**
- **标题**：ℹ️ ITC报表检查完成 - 无紧急问题
- **内容**：处理摘要和系统状态

## 🔧 配置选项说明

### 主要设置
```python
TEAMS_ENABLED = True/False      # 是否启用Teams通知
TEAMS_SEND_COMPLETION = True/False  # 是否发送完成通知
```

### Webhook配置
- `itc_notifications`: 常规通知
- `urgent_alerts`: 紧急警报（可配置到不同频道）
- `general_notifications`: 一般完成通知

### 通知规则
- **有紧急问题**: 总是发送到Teams
- **有常规问题**: 总是发送到Teams
- **无问题**: 仅在`TEAMS_SEND_COMPLETION=True`时发送

## 📱 使用场景

### 推荐配置 - 生产环境
```python
EMAIL_ENABLED = True
EMAIL_AUTO_SEND = True
SEND_COMPLETION_EMAIL = False

TEAMS_ENABLED = True
TEAMS_SEND_COMPLETION = False  # 只在有问题时通知Teams
```

### 推荐配置 - 测试环境
```python
EMAIL_ENABLED = True
EMAIL_AUTO_SEND = False  # 预览模式

TEAMS_ENABLED = True
TEAMS_SEND_COMPLETION = True  # 查看所有Teams消息
```

## 🎯 Teams vs 邮件对比

| 功能 | Teams | 邮件 |
|------|-------|------|
| 实时性 | ⚡ 即时推送 | 📧 需要检查邮箱 |
| 可见性 | 👥 团队共享 | 👤 个人接收 |
| 互动性 | 💬 支持回复讨论 | 📝 需要单独回复 |
| 格式 | 🎨 简洁卡片格式 | 📄 详细HTML格式 |
| 移动端 | 📱 Teams移动应用 | 📧 邮件客户端 |

## ⚠️ 注意事项

1. **Webhook安全**: 不要将Webhook URL提交到代码仓库
2. **频道权限**: 确保Webhook所在频道对相关人员可见
3. **网络连接**: Teams通知需要互联网连接
4. **频率限制**: 避免过于频繁的通知，以免影响工作效率
5. **备用方案**: 建议同时保持邮件通知作为备用

## 🔍 故障排除

### Teams通知不发送
1. 检查 `teams_config.json` 中 `enabled: true`
2. 检查Webhook URL格式是否正确
3. 检查网络连接和防火墙设置
4. 查看控制台错误日志

### Teams消息格式异常
1. 检查配置文件JSON格式是否正确
2. 确认Webhook URL有效性
3. 检查Teams频道是否仍然存在

### 权限问题
1. 确认Webhook创建者有频道管理权限
2. 确认频道对目标用户可见
3. 检查组织的Teams连接器政策