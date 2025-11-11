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


def load_email_config(config_path=None):
    """加载邮件配置文件，获取DC人员信息"""
    if not config_path:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_config.json")
    
    if not os.path.exists(config_path):
        return {}
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"读取邮件配置文件出错: {str(e)}")
        return {}


def get_dc_contacts_for_revoked(revoked_categories, email_config=None):
    """
    根据revoked请求的Category获取相应DC的联系人
    
    参数:
    - revoked_categories: revoked请求中涉及的Category列表
    - email_config: 邮件配置（可选，不传则自动加载）
    
    返回: DC联系人邮箱列表
    """
    if email_config is None:
        email_config = load_email_config()
    
    # 获取Revoked状态任务提醒的cc1配置
    revoked_cc1 = email_config.get("reports", {}).get("Revoked状态任务提醒", {}).get("cc1", {})
    
    contacts = []
    matched_dcs = []
    
    for category in revoked_categories:
        if not category or category.strip() == '':
            continue
            
        category_clean = str(category).strip()
        
        # 尝试精确匹配
        if category_clean in revoked_cc1:
            dc_contacts = revoked_cc1[category_clean]
            contacts.extend(dc_contacts)
            matched_dcs.append(category_clean)
            continue
        
        # 尝试模糊匹配（查找包含关系）
        for dc_key, dc_contacts in revoked_cc1.items():
            if (dc_key.lower() in category_clean.lower()) or (category_clean.lower() in dc_key.lower()):
                contacts.extend(dc_contacts)
                matched_dcs.append(dc_key)
                break
    
    # 去重
    unique_contacts = list(set(contacts))
    
    print(f"调试：Revoked Categories: {revoked_categories}")
    print(f"调试：匹配到的DC: {matched_dcs}")
    print(f"调试：对应的联系人: {unique_contacts}")
    
    return unique_contacts, matched_dcs


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
        revoked_count = log_summary.get('revoked_count', 0)
        
        # 基础消息内容
        content = f"""
**{emoji} ITC报表处理结果**

**数据统计:**
- 📊 总记录数: **{log_summary.get('total_records', 'N/A')}**
- ⏰ 紧急待审核(≤2天): **{log_summary.get('urgent_pending', 0)}**
- ⚠️ 常规待审核(≤10天): **{log_summary.get('normal_pending', 0)}**
- 🔄 Revoked状态: **{revoked_count}**
- ✅ 排除的长期任务: {log_summary.get('excluded_long_term', 0)}

{log_summary.get('action_required', '')}"""

        # 如果有revoked状态，添加@提醒相关site同事
        if revoked_count > 0:
            # 获取revoked Categories信息
            revoked_categories = log_summary.get('revoked_categories', [])
            
            # 从email_config.json中获取对应DC的联系人
            email_config = load_email_config()
            dc_contacts, matched_dcs = get_dc_contacts_for_revoked(revoked_categories, email_config)
            
            # 构建@提醒内容
            mention_text = ""
            if dc_contacts:
                mention_text = f"\n**📢 请相关DC同事关注:** {', '.join(dc_contacts)}"
                if matched_dcs:
                    mention_text += f"\n**📍 涉及DC:** {', '.join(matched_dcs)}"
            
            content += f"""

**🔄 Revoked状态特别提醒:**
发现 **{revoked_count}** 条revoked请求需要处理{mention_text}
- 对于ExitForm相关：SSO的应用/加入域的系统，可以在1年内在系统里面移除并确认，否则24小时移除
- 对于RoleChange相关：请尽快和user联系，尽快处理，在30天内在ITC确认"""

        content += """

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
            "itc_notifications": "https://pgone.webhook.office.com/webhookb2/b320358b-da36-47e8-9007-21fecd43e383@3596192b-fdf5-4e2c-a6fa-acb706c963d8/IncomingWebhook/12c8a632ac8f47f7b3814b1b4af0a640/b25908c2-19b3-42a9-b373-975bcf564b5b/V2PkxSjGEfEZ8rYrcLeNThHD38dFKT9mMxnSueWvdbgqw1",
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


def send_revoked_status_notification(revoked_items, email_config=None, teams_config=None):
    """
    发送Revoked状态任务的Teams通知
    
    参数:
    - revoked_items: Revoked状态的请求列表，每项包含: category, site_id, request_info等
    - email_config: 邮件配置（可选）
    - teams_config: Teams配置（可选）
    
    返回: (成功标志, 消息)
    """
    try:
        if email_config is None:
            email_config = load_email_config()
        
        if not revoked_items:
            return True, "无Revoked状态的请求"
        
        # 提取Categories
        revoked_categories = [item.get('category', 'Unknown') for item in revoked_items]
        
        # 获取对应DC的联系人
        dc_contacts, matched_dcs = get_dc_contacts_for_revoked(revoked_categories, email_config)
        
        # 构建项目列表
        items_text = ""
        for idx, item in enumerate(revoked_items, 1):
            category = item.get('category', 'N/A')
            site_id = item.get('site_id', 'N/A')
            request_id = item.get('request_id', 'N/A')
            items_text += f"\n  {idx}. **Category:** {category} | **SiteID:** {site_id} | **Request ID:** {request_id}"
        
        # 构建@提醒文本
        mention_text = ""
        if dc_contacts:
            # 在Teams中，通过<at>user</at>格式@提醒用户
            mention_users = " ".join([f"@{contact.split('@')[0]}" for contact in dc_contacts])
            mention_text = f"\n\n**📢 @相关Site同事:** {mention_users}"
        
        # 构建消息内容
        content = f"""
**🔴 Revoked状态任务警告**

发现 **{len(revoked_items)}** 条处于 **Revoked** 状态的请求，需要立即处理！

**待处理的请求:**
{items_text}

**🔧 处理步骤:**
1. 📱 使用 **Chrome浏览器** 打开ITC系统（其他浏览器可能存在兼容问题）
2. 🔗 访问链接: https://itc-tool.pg.com/NewRequest?siteId=193#
3. 📋 点击 **"MyTasks"** 或 **"MyActions"** 标签页
4. ✅ 完成相关任务处理

**⚠️ 处理提醒:**
- 对于 **ExitForm** 相关: SSO的应用/加入域的系统，可以在 **1年内** 在系统里移除并确认，否则 **24小时内** 移除
- 对于 **RoleChange** 相关: 请尽快和user联系，尽快处理，在 **30天内** 在ITC确认

**📍 涉及DC:** {', '.join(matched_dcs) if matched_dcs else 'N/A'}{mention_text}

---
*由ITC自动处理系统发送 - Revoked状态专项通知*
"""
        
        # 发送到Teams
        success, message = send_teams_message(
            title="🔴 Revoked状态任务 - 紧急处理",
            content=content,
            webhook_name="itc_notifications",
            urgent=True,
            teams_config=teams_config
        )
        
        return success, message
        
    except Exception as e:
        error_msg = f"发送Revoked状态Teams通知时出错: {str(e)}"
        print(error_msg)
        print(traceback.format_exc())
        return False, error_msg


def send_pending_review_status_notification(pending_items, email_config=None, teams_config=None):
    """
    发送Pending Review状态任务的Teams通知
    
    参数:
    - pending_items: Pending Review状态的请求列表，每项包含: category, site_id, request_info, days_pending等
    - email_config: 邮件配置（可选）
    - teams_config: Teams配置（可选）
    
    返回: (成功标志, 消息)
    """
    try:
        if email_config is None:
            email_config = load_email_config()
        
        if not pending_items:
            return True, "无Pending Review状态的请求"
        
        # 提取Categories
        pending_categories = [item.get('category', 'Unknown') for item in pending_items]
        
        # 获取对应DC的联系人（使用pending review相关配置）
        pending_cc1 = email_config.get("reports", {}).get("Pending Review状态任务提醒", {}).get("cc1", {})
        
        contacts = []
        matched_dcs = []
        
        for category in pending_categories:
            if not category or category.strip() == '':
                continue
            
            category_clean = str(category).strip()
            
            # 尝试精确匹配
            if category_clean in pending_cc1:
                dc_contacts = pending_cc1[category_clean]
                contacts.extend(dc_contacts)
                matched_dcs.append(category_clean)
                continue
            
            # 尝试模糊匹配
            for dc_key, dc_contacts in pending_cc1.items():
                if (dc_key.lower() in category_clean.lower()) or (category_clean.lower() in dc_key.lower()):
                    contacts.extend(dc_contacts)
                    matched_dcs.append(dc_key)
                    break
        
        # 去重
        dc_contacts = list(set(contacts))
        
        # 分类统计（按待审核天数）
        urgent_pending = [item for item in pending_items if item.get('days_pending', 999) <= 2]
        normal_pending = [item for item in pending_items if 2 < item.get('days_pending', 999) <= 10]
        
        # 构建项目列表
        items_text = ""
        for idx, item in enumerate(pending_items, 1):
            category = item.get('category', 'N/A')
            site_id = item.get('site_id', 'N/A')
            request_id = item.get('request_id', 'N/A')
            days_pending = item.get('days_pending', 'N/A')
            items_text += f"\n  {idx}. **Category:** {category} | **SiteID:** {site_id} | **Request ID:** {request_id} | ⏱️ {days_pending}天"
        
        # 构建@提醒文本
        mention_text = ""
        if dc_contacts:
            mention_users = " ".join([f"@{contact.split('@')[0]}" for contact in dc_contacts])
            mention_text = f"\n\n**📢 @相关Site同事:** {mention_users}"
        
        # 构建消息内容
        priority_info = ""
        if urgent_pending:
            priority_info = f"\n⏰ **紧急 (≤2天):** {len(urgent_pending)} 条\n"
        if normal_pending:
            priority_info += f"⚠️ **常规 (3-10天):** {len(normal_pending)} 条"
        
        content = f"""
**⏳ Pending Review状态任务提醒**

发现 **{len(pending_items)}** 条处于 **Pending Review** 状态的请求，需要关注！
{priority_info}

**待审核的请求:**
{items_text}

**🔧 处理步骤:**
1. 📱 使用 **Chrome浏览器** 打开ITC系统（其他浏览器可能存在兼容问题）
2. 🔗 访问链接: https://itc-tool.pg.com/NewRequest?siteId=193#
3. 📋 点击 **"MyTasks"** 或 **"MyActions"** 标签页进行审核
4. ✅ 完成相关审核或处理操作

**💡 审核建议:**
- 请及时审核待处理的请求
- 对于超过10天的请求，请优先处理
- 如有疑问，请联系相关的Request发起人或IT支持

**📍 涉及DC:** {', '.join(matched_dcs) if matched_dcs else 'N/A'}{mention_text}

---
*由ITC自动处理系统发送 - Pending Review状态提醒*
"""
        
        # 发送到Teams
        success, message = send_teams_message(
            title="⏳ Pending Review状态任务 - 审核提醒",
            content=content,
            webhook_name="itc_notifications",
            urgent=len(urgent_pending) > 0,
            teams_config=teams_config
        )
        
        return success, message
        
    except Exception as e:
        error_msg = f"发送Pending Review状态Teams通知时出错: {str(e)}"
        print(error_msg)
        print(traceback.format_exc())
        return False, error_msg


def send_separated_status_notifications(log_summary, teams_config=None):
    """
    分开发送Revoked和Pending Review状态的Teams通知
    
    参数:
    - log_summary: 处理结果摘要，包含:
        - revoked_items: Revoked状态的请求列表
        - pending_review_items: Pending Review状态的请求列表
    - teams_config: Teams配置（可选）
    
    返回: (总体成功标志, 消息列表)
    """
    results = []
    all_success = True
    email_config = load_email_config()
    
    # 发送Revoked状态通知
    revoked_items = log_summary.get('revoked_items', [])
    if revoked_items:
        success, message = send_revoked_status_notification(
            revoked_items, 
            email_config=email_config,
            teams_config=teams_config
        )
        results.append({
            'type': 'Revoked',
            'success': success,
            'message': message,
            'count': len(revoked_items)
        })
        all_success = all_success and success
    
    # 发送Pending Review状态通知
    pending_items = log_summary.get('pending_review_items', [])
    if pending_items:
        success, message = send_pending_review_status_notification(
            pending_items,
            email_config=email_config,
            teams_config=teams_config
        )
        results.append({
            'type': 'Pending Review',
            'success': success,
            'message': message,
            'count': len(pending_items)
        })
        all_success = all_success and success
    
    return all_success, results


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