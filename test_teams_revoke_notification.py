# -*- coding: utf-8 -*-
"""
测试Teams revoke通知功能
模拟发现revoked请求的场景
"""

import sys
import os
from datetime import datetime

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_teams_revoke_notification():
    """测试Teams revoke通知功能"""
    print("🧪 测试Teams Revoke通知功能")
    print("=" * 50)
    
    try:
        # 导入Teams发送模块
        from teams_sender import send_itc_processing_notification
        
        # 模拟发现revoked请求的日志摘要
        test_log_summary = {
            'success': True,
            'normal_pending': 2,  # 2个正常待处理
            'revoked_count': 3,   # 3个revoked请求
            'excluded_long_term': 1,
            'total_records': 6,
            'has_urgent_issues': False,
            'has_pending_issues': True,  # 有待处理问题
            'action_required': "<p style='color: #f57c00;'><strong>⚠️ 需要关注：发现5个待处理任务</strong></p>",
            'has_errors': False,
            'processing_start_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'processing_end_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'execution_time': '00:02:15'
        }
        
        print(f"📊 模拟数据摘要:")
        print(f"   ✅ 处理成功: {test_log_summary['success']}")
        print(f"   ⏳ 正常待处理: {test_log_summary['normal_pending']}")
        print(f"   🔄 Revoked状态: {test_log_summary['revoked_count']}")
        print(f"   📋 总记录数: {test_log_summary['total_records']}")
        print(f"   ⚠️ 有待处理问题: {test_log_summary['has_pending_issues']}")
        
        print(f"\n🚀 发送Teams Revoke通知...")
        
        # 发送Teams通知
        success, message = send_itc_processing_notification(test_log_summary)
        
        if success:
            print("✅ Teams Revoke通知发送成功！")
            print("📱 请检查您的Teams频道确认收到通知")
            print("💡 通知应该包含3个Revoked请求的信息")
        else:
            print(f"❌ Teams通知发送失败: {message}")
            print("🔧 可能的问题:")
            print("   1. Teams Webhook URL配置错误")
            print("   2. 网络连接问题")
            print("   3. Teams服务暂时不可用")
        
        return success
        
    except ImportError as e:
        print(f"❌ 导入Teams模块失败: {e}")
        print("💡 请确保teams_sender.py文件存在")
        return False
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def test_no_revoke_scenario():
    """测试没有revoked请求的场景"""
    print(f"\n🧪 测试无Revoke请求场景")
    print("=" * 50)
    
    try:
        from teams_sender import send_itc_processing_notification
        
        # 模拟没有revoked请求的日志摘要
        test_log_summary = {
            'success': True,
            'normal_pending': 0,
            'revoked_count': 0,    # 没有revoked请求
            'excluded_long_term': 5,
            'total_records': 5,
            'has_urgent_issues': False,
            'has_pending_issues': False,  # 没有待处理问题
            'action_required': "<p style='color: #4caf50;'><strong>✅ 状态良好：无需处理的紧急任务</strong></p>",
            'has_errors': False,
            'processing_start_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'processing_end_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'execution_time': '00:01:45'
        }
        
        print(f"📊 模拟数据摘要:")
        print(f"   ✅ 处理成功: {test_log_summary['success']}")
        print(f"   ⏳ 正常待处理: {test_log_summary['normal_pending']}")
        print(f"   🔄 Revoked状态: {test_log_summary['revoked_count']}")
        print(f"   📋 总记录数: {test_log_summary['total_records']}")
        print(f"   ⚠️ 有待处理问题: {test_log_summary['has_pending_issues']}")
        
        # 检查当前Teams配置
        from BatRun_ITCreport_downloader_rev1 import get_notification_settings
        _, _, _, teams_enabled, teams_send_completion = get_notification_settings()
        
        print(f"\n⚙️ 当前Teams配置:")
        print(f"   Teams启用: {teams_enabled}")
        print(f"   发送完成通知: {teams_send_completion}")
        
        # 根据当前逻辑判断是否会发送通知
        should_send = (
            test_log_summary['has_urgent_issues'] or 
            test_log_summary['has_pending_issues'] or 
            teams_send_completion
        )
        
        print(f"   应该发送通知: {should_send}")
        
        if not should_send:
            print("💡 根据当前配置，没有revoked请求时不会发送Teams通知")
            print("🔧 要接收所有完成通知，请设置 TEAMS_SEND_COMPLETION = True")
            return True
        else:
            print(f"\n🚀 发送Teams通知...")
            success, message = send_itc_processing_notification(test_log_summary)
            
            if success:
                print("✅ Teams通知发送成功！")
            else:
                print(f"❌ Teams通知发送失败: {message}")
                
            return success
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def suggest_configuration_changes():
    """建议配置更改"""
    print(f"\n💡 配置建议")
    print("=" * 50)
    
    print("🔧 要确保始终收到Teams通知，可以:")
    print("1. 启用完成通知：TEAMS_SEND_COMPLETION = True")
    print("2. 或者修改通知逻辑，包含更多触发条件")
    print("3. 设置专门的revoked检查阈值")
    
    print(f"\n📋 当前通知触发条件:")
    print("✅ has_urgent_issues (紧急问题)")
    print("✅ has_pending_issues (待处理问题: normal_pending > 0 或 revoked_count > 0)")
    print("✅ teams_send_completion (完成通知开关)")

def main():
    """主测试函数"""
    print("🔧 ITC Teams Revoke通知测试工具")
    print("=" * 60)
    
    # 测试有revoked请求的场景
    revoke_test_result = test_teams_revoke_notification()
    
    # 测试没有revoked请求的场景
    no_revoke_test_result = test_no_revoke_scenario()
    
    # 提供配置建议
    suggest_configuration_changes()
    
    print(f"\n📊 测试总结:")
    print("=" * 60)
    
    if revoke_test_result:
        print("🎉 Revoke通知测试成功")
        print("📱 请检查Teams频道确认收到revoked提醒")
    else:
        print("❌ Revoke通知测试失败")
    
    if no_revoke_test_result:
        print("✅ 无Revoke场景测试完成")
    else:
        print("❌ 无Revoke场景测试失败")
    
    print(f"\n🔍 问题排查:")
    print("1. 检查teams_config.json中的webhook URL是否正确")
    print("2. 确认Teams频道权限和配置")
    print("3. 验证网络连接和防火墙设置")
    print("4. 考虑调整TEAMS_SEND_COMPLETION设置")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")