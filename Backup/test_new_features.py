#!/usr/bin/env python3
"""
测试新功能：Chrome窗口重用、邮件控制选项
"""
import os
import sys

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def test_email_config():
    """测试邮件配置功能"""
    print("🧪 测试邮件配置功能...")
    
    try:
        from BatRun_ITCreport_downloader_rev1 import get_email_settings
        
        # 测试邮件配置
        email_enabled, auto_send, send_completion_email = get_email_settings()
        
        print(f"✅ 邮件配置测试通过:")
        print(f"   📧 邮件启用: {email_enabled}")
        print(f"   🚀 自动发送: {auto_send}")
        print(f"   📬 完成通知: {send_completion_email}")
        
        return True
        
    except Exception as e:
        print(f"❌ 邮件配置测试失败: {str(e)}")
        return False

def test_chrome_reuse():
    """测试Chrome重用功能"""
    print("\n🧪 测试Chrome重用功能...")
    
    try:
        from BatRun_ITCreport_downloader_rev1 import (
            check_existing_chrome_debug, 
            REUSE_EXISTING_CHROME, 
            CLOSE_CHROME_ON_EXIT
        )
        
        print(f"✅ Chrome重用配置:")
        print(f"   ♻️ 重用现有Chrome: {REUSE_EXISTING_CHROME}")
        print(f"   🔚 退出时关闭: {CLOSE_CHROME_ON_EXIT}")
        
        # 检查是否有现有的Chrome调试会话
        existing_port = check_existing_chrome_debug()
        if existing_port:
            print(f"   🔍 发现已存在的Chrome调试会话 (端口: {existing_port})")
        else:
            print(f"   📭 未发现现有的Chrome调试会话")
            
        return True
        
    except Exception as e:
        print(f"❌ Chrome重用测试失败: {str(e)}")
        return False

def test_port_allocation():
    """测试端口分配功能"""
    print("\n🧪 测试端口分配功能...")
    
    try:
        from BatRun_ITCreport_downloader_rev1 import (
            allocate_debug_port, 
            ITC_PORT_RANGE,
            is_port_available
        )
        
        print(f"✅ 端口配置:")
        print(f"   📦 ITC端口范围: {ITC_PORT_RANGE['start']}-{ITC_PORT_RANGE['end']}")
        
        # 测试端口分配
        available_port = allocate_debug_port()
        if available_port:
            print(f"   🎯 分配到端口: {available_port}")
            print(f"   🔍 端口可用性: {is_port_available(available_port)}")
        else:
            print(f"   ⚠️ 当前没有可用端口")
            
        return True
        
    except Exception as e:
        print(f"❌ 端口分配测试失败: {str(e)}")
        return False

def main():
    """主测试函数"""
    print("🚀 开始测试新功能...")
    print("=" * 50)
    
    test_results = []
    
    # 测试各个功能
    test_results.append(test_email_config())
    test_results.append(test_chrome_reuse())
    test_results.append(test_port_allocation())
    
    # 汇总结果
    print("\n" + "=" * 50)
    passed_tests = sum(test_results)
    total_tests = len(test_results)
    
    if passed_tests == total_tests:
        print(f"🎉 所有测试通过! ({passed_tests}/{total_tests})")
        print("✨ 新功能已准备就绪！")
    else:
        print(f"⚠️ 部分测试失败: {passed_tests}/{total_tests}")
        print("🔧 请检查失败的功能")

if __name__ == "__main__":
    main()