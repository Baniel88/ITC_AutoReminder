# -*- coding: utf-8 -*-
"""
Teams消息发送模块
支持通过Webhook向Teams频道发送格式化消息
"""

import requests
import json
import traceback
from datetime import datetime
import os


def load_teams_config(config_path=None):
    """加载Teams配置文件"""
    if not config_path:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "teams_config.json")
    
    if not os.path.exists(config_path):
        # 如果配置文件不存在，返回默认配置
        return {
            "enabled": False,
            "webhooks": {},
            "default_webhook": None
        }
    
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def send_teams_message(title, content, webhook_name="default", urgent=False, teams_config=None):
    """
    发送消息到Teams频道
    
    参数:
    - title: 消息标题
    - content: 消息内容（支持简单的Markdown）
    - webhook_name: 使用的webhook名称（在配置文件中定义）
    - urgent: 是否为紧急消息（影响颜色和提醒）
    - teams_config: Teams配置（可选，不传则自动加载）
    
    返回: (成功标志, 消息)
    """
    try:
        # 加载配置
        if teams_config is None:
            teams_config = load_teams_config()
        
        if not teams_config.get("enabled", False):
            return False, "Teams消息功能已禁用"
        
        # 获取webhook URL
        webhooks = teams_config.get("webhooks", {})
        
        if webhook_name == "default":
            webhook_name = teams_config.get("default_webhook", "itc_notifications")
        
        if webhook_name not in webhooks:
            return False, f"未找到webhook配置: {webhook_name}"
        
        webhook_url = webhooks[webhook_name]
        
        # 构建Teams消息格式（Adaptive Cards）
        if urgent:
            theme_color = "FF0000"  # 红色
            activity_title = f"🚨 {title}"
            activity_subtitle = "紧急通知 - 请立即处理"
        else:
            theme_color = "0078D4"  # Teams蓝色
            activity_title = f"ℹ️ {title}"
            activity_subtitle = "系统通知"
        
        # 构建消息体
        card_content = {
            "@type": "MessageCard",
            "@context": "https://schema.org/extensions",
            "themeColor": theme_color,
            "summary": title,
            "sections": [
                {
                    "activityTitle": activity_title,
                    "activitySubtitle": activity_subtitle,
                    "activityImage": "https://teamsnodesample.azurewebsites.net/static/img/image5.png",
                    "facts": [
                        {
                            "name": "发送时间",
                            "value": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        },
                        {
                            "name": "系统",
                            "value": "ITC报表自动处理系统"
                        }
                    ],
                    "markdown": True,
                    "text": content
                }
            ]
        }
        
        # 发送消息
        headers = {
            'Content-Type': 'application/json'
        }
        
        response = requests.post(
            webhook_url, 
            data=json.dumps(card_content), 
            headers=headers,
            timeout=30
        )
        
        if response.status_code == 200:
            return True, "Teams消息发送成功"
        else:
            return False, f"Teams消息发送失败: HTTP {response.status_code}"
            
    except Exception as e:
        error_msg = f"发送Teams消息时出错: {str(e)}"
        print(error_msg)
        print(traceback.format_exc())
        return False, error_msg


def send_itc_processing_notification(log_summary, teams_config=None):
    """
    发送ITC处理结果到Teams
    
    参数:
    - log_summary: 处理结果摘要（与邮件系统相同的数据结构）
    - teams_config: Teams配置（可选）
    
    返回: (成功标志, 消息)
    """
    try:
        # 判断紧急程度
        is_urgent = log_summary.get('has_urgent_issues', False)
        
        # 构建标题
        if log_summary.get('has_urgent_issues'):
            title = "ITC报表发现紧急问题"
            emoji = "🚨"
        elif log_summary.get('has_pending_issues'):
            title = "ITC报表发现待处理任务"
            emoji = "⚠️"
        else:
            title = "ITC报表检查完成 - 无紧急问题"
            emoji = "✅"
        
        # 构建消息内容
        content = f"""
**{emoji} ITC报表处理结果**

**数据统计:**
- 📊 总记录数: **{log_summary.get('total_records', 'N/A')}**
- ⏰ 紧急待审核(≤2天): **{log_summary.get('urgent_pending', 0)}**
- ⚠️ 常规待审核(≤10天): **{log_summary.get('normal_pending', 0)}**
- 🔄 Revoked状态: **{log_summary.get('revoked_count', 0)}**
- ✅ 排除的长期任务: {log_summary.get('excluded_long_term', 0)}

{log_summary.get('action_required', '')}

---
*由ITC自动处理系统发送*
"""
        
        # 发送到Teams
        return send_teams_message(
            title=title,
            content=content,
            webhook_name="itc_notifications",
            urgent=is_urgent,
            teams_config=teams_config
        )
        
    except Exception as e:
        error_msg = f"构建ITC Teams通知时出错: {str(e)}"
        print(error_msg)
        return False, error_msg


def create_teams_config_template():
    """创建Teams配置文件模板"""
    config_template = {
        "enabled": True,
        "default_webhook": "itc_notifications",
        "webhooks": {
            "itc_notifications": "https://pgone.webhook.office.com/webhookb2/b320358b-da36-47e8-9007-21fecd43e383@3596192b-fdf5-4e2c-a6fa-acb706c963d8/IncomingWebhook/ef9e46bea68647d1853d7653cb713c4c/b25908c2-19b3-42a9-b373-975bcf564b5b/V2uJRtxbhqR61xeTNhSeT5DgBSLfR1elVDS8LjQunHb5k1",
            "urgent_alerts": "https://pg.webhook.office.com/webhookb2/YOUR_URGENT_WEBHOOK_URL_HERE",
            "general_notifications": "https://pg.webhook.office.com/webhookb2/YOUR_GENERAL_WEBHOOK_URL_HERE"
        },
        "team_members": {
            "itc_team": [
                "user1@pg.com",
                "user2@pg.com"
            ],
            "managers": [
                "manager1@pg.com",
                "manager2@pg.com"
            ]
        },
        "notification_rules": {
            "urgent_issues": {
                "webhook": "urgent_alerts",
                "mention_team": "itc_team"
            },
            "normal_issues": {
                "webhook": "itc_notifications",
                "mention_team": None
            },
            "completion_notice": {
                "webhook": "general_notifications",
                "mention_team": None
            }
        }
    }
    
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "teams_config.json")
    
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config_template, f, indent=4, ensure_ascii=False)
    
    return config_path


if __name__ == "__main__":
    # 测试代码
    print("=== Teams消息发送模块测试 ===")
    
    # 创建配置文件模板
    config_path = create_teams_config_template()
    print(f"已创建配置文件模板: {config_path}")
    print("请编辑此文件，填入正确的Webhook URL")
    
    # 测试消息发送（使用示例数据）
    test_summary = {
        'total_records': 150,
        'urgent_pending': 3,
        'normal_pending': 8,
        'revoked_count': 2,
        'excluded_long_term': 5,
        'has_urgent_issues': True,
        'has_pending_issues': True,
        'action_required': '**建议立即处理3个紧急项目**'
    }
    
    success, message = send_itc_processing_notification(test_summary)
    print(f"测试结果: {success} - {message}")