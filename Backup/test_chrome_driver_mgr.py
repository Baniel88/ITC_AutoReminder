# -*- coding: utf-8 -*-
"""
Chrome Driver 管理器测试脚本
用于测试Chrome_Driver_mgr.py模块是否正常工作
"""

import os
import sys
from datetime import datetime

def test_chrome_driver_manager():
    """测试Chrome Driver管理器"""
    print("="*60)
    print("Chrome Driver 管理器测试")
    print(f"测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    try:
        # 导入Chrome Driver管理器
        from Chrome_Driver_mgr import ChromeDriverManager, check_chrome_environment
        print("✅ 成功导入Chrome Driver管理器模块")
        
        # 测试环境检查
        print("\n📋 开始环境检查...")
        env_ok = check_chrome_environment()
        
        if env_ok:
            print("\n🎉 环境检查通过！")
            
            # 测试管理器实例
            print("\n🔧 测试管理器功能...")
            manager = ChromeDriverManager()
            
            # 获取Chrome路径
            chrome_path = manager.get_chrome_path()
            print(f"Chrome路径: {chrome_path}")
            
            # 获取Chrome版本
            chrome_version = manager.get_chrome_full_version()
            print(f"Chrome版本: {chrome_version}")
            
            # 获取ChromeDriver路径
            driver_path = manager.get_chromedriver_path()
            print(f"ChromeDriver路径: {driver_path}")
            
            # 获取Driver版本
            driver_version = manager.get_driver_version()
            print(f"ChromeDriver版本: {driver_version}")
            
            # 检查兼容性
            compatible = manager.is_driver_compatible()
            print(f"版本兼容性: {'✅ 兼容' if compatible else '❌ 不兼容'}")
            
        else:
            print("\n❌ 环境检查未通过")
            print("请检查Chrome和ChromeDriver安装")
            
    except ImportError as e:
        print(f"❌ 导入失败: {str(e)}")
        print("请确保Chrome_Driver_mgr.py文件在同一目录下")
        return False
    except Exception as e:
        print(f"❌ 测试过程中出错: {str(e)}")
        return False
    
    print("\n" + "="*60)
    print("测试完成")
    print("="*60)
    return env_ok


def test_integration_with_original_script():
    """测试与原始脚本的集成"""
    print("\n" + "="*60)
    print("集成测试 - 验证与原始脚本的兼容性")
    print("="*60)
    
    try:
        # 模拟原始脚本的导入方式
        from Chrome_Driver_mgr import get_chromedriver_path, get_chrome_path
        
        print("✅ 成功导入便捷函数")
        
        # 测试便捷函数
        def simple_log(msg):
            print(f"[测试日志] {msg}")
        
        print("\n🔧 测试便捷函数...")
        driver_path = get_chromedriver_path(log_callback=simple_log)
        chrome_path = get_chrome_path(log_callback=simple_log)
        
        print(f"\n📋 测试结果:")
        print(f"ChromeDriver路径: {driver_path}")
        print(f"Chrome路径: {chrome_path}")
        
        success = driver_path is not None and chrome_path is not None
        print(f"集成测试: {'✅ 通过' if success else '❌ 失败'}")
        return success
        
    except Exception as e:
        print(f"❌ 集成测试失败: {str(e)}")
        return False


if __name__ == "__main__":
    print("Chrome Driver 管理器 - 完整测试套件")
    print(f"当前工作目录: {os.getcwd()}")
    print(f"脚本目录: {os.path.dirname(os.path.abspath(__file__))}")
    
    # 运行基本测试
    test1_passed = test_chrome_driver_manager()
    
    # 运行集成测试
    test2_passed = test_integration_with_original_script()
    
    # 输出总结
    print("\n" + "="*60)
    print("测试总结")
    print("="*60)
    print(f"基本功能测试: {'✅ 通过' if test1_passed else '❌ 失败'}")
    print(f"集成兼容测试: {'✅ 通过' if test2_passed else '❌ 失败'}")
    
    if test1_passed and test2_passed:
        print("\n🎉 所有测试通过！Chrome Driver管理器可以正常使用。")
        print("💡 现在可以在其他项目中导入并使用这个模块了。")
    else:
        print("\n❌ 部分测试失败，请检查问题并修复。")
    
    print("\n按任意键退出...")
    input()