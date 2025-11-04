# -*- coding: utf-8 -*-
"""
邮件自动发送测试脚本
测试EMAIL_AUTO_SEND配置是否正确生效
"""

import sys
import os
from datetime import datetime

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_email_auto_send():
    """测试邮件自动发送功能"""
    print("🧪 测试邮件自动发送功能")
    print("=" * 50)
    
    try:
        # 导入配置获取函数
        from BatRun_ITCreport_downloader_rev1 import get_notification_settings
        
        # 获取当前配置
        email_enabled, auto_send, send_completion_email, teams_enabled, teams_send_completion = get_notification_settings()
        
        print(f"📧 当前邮件配置:")
        print(f"   邮件启用: {email_enabled}")
        print(f"   自动发送: {auto_send}")
        print(f"   完成通知: {send_completion_email}")
        print(f"   Teams启用: {teams_enabled}")
        print(f"   Teams完成通知: {teams_send_completion}")
        
        if not email_enabled:
            print("❌ 邮件功能已禁用，无法测试")
            return False
        
        if not auto_send:
            print("⚠️ 邮件配置为预览模式，需要手动发送")
            print("💡 要启用自动发送，请设置:")
            print("   主文件: EMAIL_AUTO_SEND = True")
            print("   或配置文件: AUTO_SEND_EMAIL = true")
            return False
        
        # 测试邮件发送
        print(f"\n🚀 测试自动发送邮件...")
        
        from email_sender import send_email
        
        test_subject = f"ITC邮件自动发送测试 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        test_content = f"""
        <div style="background-color: #e8f5e9; border-left: 4px solid #4caf50; padding: 15px; margin: 10px 0;">
            <h2>📧 邮件自动发送测试</h2>
            <p><strong>测试时间：</strong>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            <p><strong>配置状态：</strong></p>
            <ul>
                <li>邮件启用: ✅ {email_enabled}</li>
                <li>自动发送: ✅ {auto_send}</li>
                <li>完成通知: {'✅' if send_completion_email else '❌'} {send_completion_email}</li>
            </ul>
            <p><strong>测试结果：</strong></p>
            <p>🎉 如果您收到此邮件，说明自动发送功能正常工作！</p>
            <hr>
            <p style="font-size: 12px; color: #666;"><em>此邮件由ITC邮件自动发送测试工具生成</em></p>
        </div>
        """
        
        to_addrs = ["liang.wq.1@pg.com"]
        cc_addrs = ["liang.wq.1@pg.com"]
        
        # 明确指定 auto_send=True 进行测试
        send_email(
            subject=test_subject,
            html_content=test_content,
            to_addrs=to_addrs,
            cc_addrs=cc_addrs,
            auto_send=True  # 强制自动发送测试
        )
        
        print("✅ 测试邮件已发送")
        print("📱 请检查您的邮箱和已发送邮件文件夹")
        print("💡 如果收到邮件，说明自动发送功能正常")
        
        return True
        
    except ImportError as e:
        print(f"❌ 导入模块失败: {e}")
        return False
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def test_config_priority():
    """测试配置文件优先级"""
    print(f"\n🔍 检查配置文件优先级...")
    
    # 检查主文件配置
    try:
        with open("BatRun_ITCreport_downloader_rev1.py", 'r', encoding='utf-8') as f:
            content = f.read()
            
        import re
        auto_send_match = re.search(r'EMAIL_AUTO_SEND\s*=\s*(True|False)', content)
        if auto_send_match:
            main_auto_send = auto_send_match.group(1)
            print(f"📄 主文件配置: EMAIL_AUTO_SEND = {main_auto_send}")
        
    except Exception as e:
        print(f"⚠️ 读取主文件配置失败: {e}")
    
    # 检查配置文件
    try:
        import json
        with open("email_config.json", 'r', encoding='utf-8') as f:
            config = json.load(f)
            
        auto_send_config = config.get('system_config', {}).get('AUTO_SEND_EMAIL')
        print(f"⚙️ 配置文件设置: AUTO_SEND_EMAIL = {auto_send_config}")
        
        if auto_send_config is not None:
            print("💡 配置文件将覆盖主文件设置")
            
    except Exception as e:
        print(f"⚠️ 读取配置文件失败: {e}")

def main():
    """主测试函数"""
    print("🔧 ITC邮件自动发送测试工具")
    print("=" * 60)
    
    # 检查配置优先级
    test_config_priority()
    
    # 测试自动发送
    test_result = test_email_auto_send()
    
    print(f"\n📊 测试总结:")
    print("=" * 60)
    
    if test_result:
        print("🎉 邮件自动发送功能测试完成")
        print("📧 请检查邮箱确认是否收到测试邮件")
        print("✅ 如果收到邮件且在已发送文件夹中看到，说明配置正确")
    else:
        print("❌ 邮件自动发送功能测试失败")
        print("🔧 请检查配置并修复问题")
    
    print(f"\n💡 问题解决建议:")
    print(f"   1. 确保 email_config.json 中 AUTO_SEND_EMAIL = true")
    print(f"   2. 确保 Outlook 已登录并正常工作")
    print(f"   3. 检查网络连接和防火墙设置")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")