# -*- coding: utf-8 -*-
"""
Teams消息发送测试脚本
快速测试Teams通知功能
"""

from teams_sender import send_teams_message, send_itc_processing_notification, load_teams_config
import json

def test_basic_message():
    """测试基础消息发送"""
    print("🧪 测试基础Teams消息发送...")
    
    success, message = send_teams_message(
        title="ITC系统测试通知",
        content="""
这是一条来自ITC报表自动处理系统的测试消息。

**测试信息:**
- 🕒 测试时间: 现在
- 🔧 功能: Teams集成测试
- ✅ 状态: Webhook连接正常

如果您收到此消息，说明Teams集成配置成功！
        """,
        urgent=False
    )
    
    if success:
        print("✅ 基础消息发送成功！")
    else:
        print(f"❌ 基础消息发送失败: {message}")
    
    return success

def test_urgent_message():
    """测试紧急消息发送"""
    print("\n🚨 测试紧急消息发送...")
    
    success, message = send_teams_message(
        title="紧急测试通知",
        content="""
**这是一个紧急消息测试**

🚨 **紧急情况模拟:**
- 发现 3 个紧急待审核项目
- 超时天数: ≤2天
- 需要立即处理

这条消息应该显示为红色紧急样式。
        """,
        urgent=True
    )
    
    if success:
        print("✅ 紧急消息发送成功！")
    else:
        print(f"❌ 紧急消息发送失败: {message}")
    
    return success

def test_itc_notifications():
    """测试ITC处理结果通知"""
    print("\n📊 测试ITC处理结果通知...")
    
    # 测试有问题的情况
    print("   测试情况1: 发现紧急问题")
    test_summary_urgent = {
        'total_records': 150,
        'urgent_pending': 3,
        'normal_pending': 7,
        'revoked_count': 2,
        'excluded_long_term': 12,
        'has_urgent_issues': True,
        'has_pending_issues': True,
        'action_required': '**建议立即处理3个紧急项目**'
    }
    
    success1, msg1 = send_itc_processing_notification(test_summary_urgent)
    print(f"      结果: {'✅' if success1 else '❌'} {msg1}")
    
    # 测试常规问题情况
    print("   测试情况2: 常规待处理项目")
    test_summary_normal = {
        'total_records': 120,
        'urgent_pending': 0,
        'normal_pending': 5,
        'revoked_count': 1,
        'excluded_long_term': 8,
        'has_urgent_issues': False,
        'has_pending_issues': True,
        'action_required': '**建议关注5个常规待处理项目**'
    }
    
    success2, msg2 = send_itc_processing_notification(test_summary_normal)
    print(f"      结果: {'✅' if success2 else '❌'} {msg2}")
    
    # 测试无问题情况
    print("   测试情况3: 无需关注")
    test_summary_ok = {
        'total_records': 95,
        'urgent_pending': 0,
        'normal_pending': 0,
        'revoked_count': 0,
        'excluded_long_term': 18,
        'has_urgent_issues': False,
        'has_pending_issues': False,
        'action_required': '**当前无需特别关注的项目**'
    }
    
    success3, msg3 = send_itc_processing_notification(test_summary_ok)
    print(f"      结果: {'✅' if success3 else '❌'} {msg3}")
    
    return success1 and success2 and success3

def show_config_info():
    """显示当前配置信息"""
    print("📋 当前Teams配置信息:")
    print("=" * 60)
    
    try:
        config = load_teams_config()
        
        print(f"启用状态: {'✅ 已启用' if config.get('enabled', False) else '❌ 已禁用'}")
        print(f"默认Webhook: {config.get('default_webhook', 'None')}")
        
        webhooks = config.get('webhooks', {})
        print(f"\nWebhook配置:")
        for name, url in webhooks.items():
            if "YOUR_" in url:
                status = "❌ 未配置"
            else:
                status = "✅ 已配置"
            print(f"  {name}: {status}")
            print(f"    URL: {url[:50]}..." if len(url) > 50 else f"    URL: {url}")
        
        team_members = config.get('team_members', {})
        print(f"\n团队成员:")
        for team, members in team_members.items():
            print(f"  {team}: {', '.join(members)}")
            
    except Exception as e:
        print(f"❌ 读取配置失败: {str(e)}")

def main():
    """主测试函数"""
    print("🚀 Teams集成功能测试")
    print("=" * 60)
    
    # 显示配置
    show_config_info()
    
    print(f"\n🧪 开始测试...")
    
    # 测试序列
    tests = [
        ("基础消息", test_basic_message),
        ("紧急消息", test_urgent_message),
        ("ITC通知", test_itc_notifications)
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ {test_name}测试出错: {str(e)}")
            results.append((test_name, False))
    
    # 测试结果总结
    print(f"\n� 测试结果总结:")
    print("=" * 60)
    
    all_passed = True
    for test_name, result in results:
        status = "✅ 通过" if result else "❌ 失败"
        print(f"{test_name}: {status}")
        if not result:
            all_passed = False
    
    if all_passed:
        print(f"\n🎉 所有测试通过！Teams集成功能正常")
        print(f"📱 请检查Teams频道确认收到所有测试消息")
    else:
        print(f"\n⚠️ 部分测试失败，请检查配置和网络连接")
    
    print(f"\n💡 使用提示:")
    print(f"   - 紧急消息显示为红色样式")
    print(f"   - 常规消息显示为蓝色样式")
    print(f"   - ITC处理结果会根据问题严重程度自动调整样式")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")