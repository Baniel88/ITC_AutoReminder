# -*- coding: utf-8 -*-
"""
快速验证Teams完成通知配置
"""

import sys
import os
from datetime import datetime

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_teams_completion_notification():
    """测试Teams完成通知配置"""
    print("🔧 验证Teams完成通知配置")
    print("=" * 50)
    
    try:
        # 检查配置更新
        from BatRun_ITCreport_downloader_rev1 import get_notification_settings
        
        email_enabled, auto_send, send_completion_email, teams_enabled, teams_send_completion = get_notification_settings()
        
        print(f"📋 当前通知配置:")
        print(f"   📧 邮件启用: {email_enabled}")
        print(f"   📧 邮件自动发送: {auto_send}")
        print(f"   📧 邮件完成通知: {send_completion_email}")
        print(f"   📱 Teams启用: {teams_enabled}")
        print(f"   📱 Teams完成通知: {teams_send_completion}")
        
        if teams_enabled and teams_send_completion:
            print(f"\n✅ 配置正确！现在将发送所有Teams通知")
            print(f"   - 有revoked请求时：✅ 发送")
            print(f"   - 没有问题时：✅ 也会发送完成通知")
            
            # 测试无问题场景的Teams通知
            print(f"\n🧪 测试无问题场景的Teams通知...")
            
            from teams_sender import send_itc_processing_notification
            
            # 模拟无问题的完成通知
            test_log_summary = {
                'success': True,
                'normal_pending': 0,
                'revoked_count': 0,
                'excluded_long_term': 10,
                'total_records': 10,
                'has_urgent_issues': False,
                'has_pending_issues': False,
                'action_required': "<p style='color: #4caf50;'><strong>✅ 状态良好：无需处理的紧急任务</strong></p>",
                'has_errors': False,
                'processing_start_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'processing_end_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'execution_time': '00:01:30'
            }
            
            # 根据新配置判断是否发送
            should_send = (
                test_log_summary['has_urgent_issues'] or 
                test_log_summary['has_pending_issues'] or 
                teams_send_completion  # 现在应该是True
            )
            
            print(f"   应该发送通知: {should_send}")
            
            if should_send:
                success, message = send_itc_processing_notification(test_log_summary)
                
                if success:
                    print("🎉 完成通知测试发送成功！")
                    print("📱 请检查Teams频道确认收到完成通知")
                else:
                    print(f"❌ 完成通知发送失败: {message}")
            else:
                print("❌ 配置可能没有生效")
                
        else:
            if not teams_enabled:
                print(f"\n⚠️ Teams功能已禁用")
            if not teams_send_completion:
                print(f"\n⚠️ Teams完成通知仍然禁用")
                print(f"🔧 请检查配置文件覆盖设置")
        
        return teams_enabled and teams_send_completion
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def main():
    """主函数"""
    print("🔧 ITC Teams完成通知配置验证")
    print("=" * 60)
    
    result = test_teams_completion_notification()
    
    print(f"\n📊 验证结果:")
    print("=" * 60)
    
    if result:
        print("✅ Teams完成通知配置成功")
        print("📱 现在即使没有revoked请求也会收到完成通知")
        print("🎯 下次运行ITC程序时会发送Teams通知")
    else:
        print("❌ Teams完成通知配置有问题")
        print("🔧 请检查配置文件和程序设置")
    
    print(f"\n💡 使用建议:")
    print("1. 运行完整的ITC程序测试Teams通知")
    print("2. 检查Teams频道确认收到通知")
    print("3. 如有问题，检查teams_config.json配置")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")