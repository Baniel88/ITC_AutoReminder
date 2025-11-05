# -*- coding: utf-8 -*-
"""
端口冲突检测和Chrome会话管理测试
"""

import socket
import requests
import json
from chrome_port_config import EVPChromeConfig, ITCChromeConfig, get_project_config

def check_port_status(port):
    """检查端口状态"""
    try:
        # 检查端口是否被占用
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(2)
            if sock.connect_ex(("127.0.0.1", port)) == 0:
                # 端口被占用，检查是否是Chrome调试端口
                try:
                    response = requests.get(f"http://127.0.0.1:{port}/json", timeout=3)
                    if response.status_code == 200:
                        data = response.json()
                        return "Chrome调试端口", len(data)
                    else:
                        return "其他服务", 0
                except:
                    return "未知服务", 0
            else:
                return "空闲", 0
    except Exception as e:
        return f"检查错误: {str(e)}", 0

def scan_all_ports():
    """扫描所有相关端口"""
    print("🔍 Chrome端口占用情况扫描")
    print("=" * 80)
    
    # EVP端口范围
    print(f"\n📦 EVP Scorecard 端口范围 ({EVPChromeConfig.PORT_START}-{EVPChromeConfig.PORT_END}):")
    evp_chrome_count = 0
    for port in range(EVPChromeConfig.PORT_START, EVPChromeConfig.PORT_END + 1):
        status, tabs = check_port_status(port)
        if "Chrome" in status:
            evp_chrome_count += 1
            print(f"   端口 {port}: {status} ({tabs} 标签页) ⚠️")
        elif status != "空闲":
            print(f"   端口 {port}: {status} ❌")
        else:
            print(f"   端口 {port}: {status} ✅")
    
    # ITC端口范围
    print(f"\n📦 ITC Scorecard 端口范围 ({ITCChromeConfig.PORT_START}-{ITCChromeConfig.PORT_END}):")
    itc_chrome_count = 0
    for port in range(ITCChromeConfig.PORT_START, ITCChromeConfig.PORT_END + 1):
        status, tabs = check_port_status(port)
        if "Chrome" in status:
            itc_chrome_count += 1
            print(f"   端口 {port}: {status} ({tabs} 标签页) ⚠️")
        elif status != "空闲":
            print(f"   端口 {port}: {status} ❌")
        else:
            print(f"   端口 {port}: {status} ✅")
    
    # 总结
    print(f"\n📊 扫描结果总结:")
    print(f"   EVP Chrome会话: {evp_chrome_count} 个")
    print(f"   ITC Chrome会话: {itc_chrome_count} 个")
    
    if evp_chrome_count > 0 and itc_chrome_count > 0:
        print("   ✅ 端口隔离正常，两个项目可以并行运行")
    elif evp_chrome_count > 0:
        print("   ℹ️ 只有EVP项目在运行")
    elif itc_chrome_count > 0:
        print("   ℹ️ 只有ITC项目在运行")
    else:
        print("   ℹ️ 没有检测到Chrome调试会话")

def check_chrome_tabs(port):
    """检查Chrome标签页内容"""
    try:
        response = requests.get(f"http://127.0.0.1:{port}/json", timeout=3)
        if response.status_code == 200:
            tabs = response.json()
            print(f"\n🌐 端口 {port} 的Chrome标签页:")
            for i, tab in enumerate(tabs, 1):
                title = tab.get('title', 'No Title')[:50]
                url = tab.get('url', 'No URL')[:80]
                print(f"   标签 {i}: {title}")
                print(f"           URL: {url}")
                
                # 检查是否是ITC页面但使用了EVP端口
                if 'itc-tool.pg.com' in url and EVPChromeConfig.PORT_START <= port <= EVPChromeConfig.PORT_END:
                    print(f"           ⚠️ 警告: ITC页面运行在EVP端口范围内!")
                elif 'evp' in url.lower() and ITCChromeConfig.PORT_START <= port <= ITCChromeConfig.PORT_END:
                    print(f"           ⚠️ 警告: EVP页面运行在ITC端口范围内!")
            return True
    except Exception as e:
        print(f"   获取标签页信息失败: {str(e)}")
        return False

def main():
    print("🛠️ Chrome端口冲突检测工具")
    print("=" * 60)
    
    # 显示配置信息
    print(f"\n📋 端口配置信息:")
    print(f"   EVP Scorecard: {EVPChromeConfig.PORT_START}-{EVPChromeConfig.PORT_END}")
    print(f"   ITC Scorecard: {ITCChromeConfig.PORT_START}-{ITCChromeConfig.PORT_END}")
    
    # 扫描端口
    scan_all_ports()
    
    # 检查有问题的Chrome会话
    print(f"\n🔍 详细Chrome会话检查:")
    problem_found = False
    
    # 检查EVP端口范围
    for port in range(EVPChromeConfig.PORT_START, EVPChromeConfig.PORT_END + 1):
        status, _ = check_port_status(port)
        if "Chrome" in status:
            if check_chrome_tabs(port):
                problem_found = True
    
    # 检查ITC端口范围
    for port in range(ITCChromeConfig.PORT_START, ITCChromeConfig.PORT_END + 1):
        status, _ = check_port_status(port)
        if "Chrome" in status:
            if check_chrome_tabs(port):
                problem_found = True
    
    if not problem_found:
        print("   ✅ 未发现Chrome会话")
    
    print(f"\n💡 建议操作:")
    print(f"   1. 如发现端口冲突，请关闭冲突的Chrome会话")
    print(f"   2. 确保EVP项目只使用{EVPChromeConfig.PORT_START}-{EVPChromeConfig.PORT_END}端口")
    print(f"   3. 确保ITC项目只使用{ITCChromeConfig.PORT_START}-{ITCChromeConfig.PORT_END}端口")
    print(f"   4. 使用REUSE_EXISTING_CHROME=True避免重复打开新窗口")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")