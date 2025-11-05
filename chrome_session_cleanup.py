# -*- coding: utf-8 -*-
"""
Chrome会话清理工具
用于清理冲突的Chrome调试会话，确保项目间端口隔离
"""

import socket
import requests
import subprocess
import sys
from chrome_port_config import EVPChromeConfig, ITCChromeConfig

def get_chrome_sessions():
    """获取所有Chrome调试会话"""
    sessions = []
    
    # 检查所有可能的端口范围
    all_ports = list(range(EVPChromeConfig.PORT_START, EVPChromeConfig.PORT_END + 1)) + \
                list(range(ITCChromeConfig.PORT_START, ITCChromeConfig.PORT_END + 1))
    
    for port in all_ports:
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.settimeout(2)
                if sock.connect_ex(("127.0.0.1", port)) == 0:
                    # 检查是否是Chrome调试端口
                    try:
                        response = requests.get(f"http://127.0.0.1:{port}/json", timeout=3)
                        if response.status_code == 200:
                            tabs = response.json()
                            
                            # 分析标签页内容
                            project = "Unknown"
                            urls = [tab.get('url', '') for tab in tabs]
                            
                            if any('itc-tool.pg.com' in url for url in urls):
                                if EVPChromeConfig.PORT_START <= port <= EVPChromeConfig.PORT_END:
                                    project = "ITC-on-EVP-Port ⚠️"  # 问题会话
                                else:
                                    project = "ITC"
                            elif any('evp' in url.lower() for url in urls):
                                if ITCChromeConfig.PORT_START <= port <= ITCChromeConfig.PORT_END:
                                    project = "EVP-on-ITC-Port ⚠️"  # 问题会话
                                else:
                                    project = "EVP"
                            else:
                                # 根据端口范围推断
                                if EVPChromeConfig.PORT_START <= port <= EVPChromeConfig.PORT_END:
                                    project = "EVP-Range"
                                else:
                                    project = "ITC-Range"
                            
                            sessions.append({
                                'port': port,
                                'project': project,
                                'tab_count': len(tabs),
                                'urls': urls[:3],  # 只保留前3个URL
                                'is_conflict': '⚠️' in project
                            })
                    except Exception:
                        # 端口被占用但不是Chrome
                        sessions.append({
                            'port': port,
                            'project': 'Non-Chrome',
                            'tab_count': 0,
                            'urls': [],
                            'is_conflict': False
                        })
        except Exception:
            continue
    
    return sessions

def display_sessions(sessions):
    """显示Chrome会话信息"""
    print("🌐 当前Chrome调试会话:")
    print("=" * 80)
    
    if not sessions:
        print("   ✅ 没有发现Chrome调试会话")
        return
    
    conflict_sessions = []
    normal_sessions = []
    
    for session in sessions:
        if session['is_conflict']:
            conflict_sessions.append(session)
        else:
            normal_sessions.append(session)
    
    # 显示冲突会话
    if conflict_sessions:
        print("⚠️ 发现端口冲突的会话:")
        for session in conflict_sessions:
            print(f"   端口 {session['port']}: {session['project']} ({session['tab_count']} 标签页)")
            for url in session['urls']:
                if url:
                    print(f"      📄 {url[:70]}...")
    
    # 显示正常会话
    if normal_sessions:
        print(f"\n✅ 正常的Chrome会话:")
        for session in normal_sessions:
            if session['project'] != 'Non-Chrome':
                print(f"   端口 {session['port']}: {session['project']} ({session['tab_count']} 标签页)")
    
    return conflict_sessions

def close_chrome_session(port):
    """关闭指定端口的Chrome会话"""
    try:
        # 尝试通过API关闭
        response = requests.get(f"http://127.0.0.1:{port}/json", timeout=3)
        if response.status_code == 200:
            tabs = response.json()
            for tab in tabs:
                tab_id = tab.get('id')
                if tab_id:
                    try:
                        requests.post(f"http://127.0.0.1:{port}/json/close/{tab_id}", timeout=2)
                    except:
                        pass
        
        print(f"   ✅ 已关闭端口 {port} 的Chrome会话")
        return True
        
    except Exception as e:
        print(f"   ❌ 关闭端口 {port} 失败: {str(e)}")
        return False

def interactive_cleanup():
    """交互式清理Chrome会话"""
    print("🧹 Chrome会话清理工具")
    print("=" * 60)
    
    sessions = get_chrome_sessions()
    conflict_sessions = display_sessions(sessions)
    
    if not conflict_sessions:
        print("\n🎉 没有发现端口冲突的Chrome会话!")
        return
    
    print(f"\n🛠️ 清理选项:")
    print(f"   1. 自动清理所有冲突会话")
    print(f"   2. 手动选择要清理的会话")
    print(f"   3. 只显示信息，不清理")
    print(f"   0. 退出")
    
    try:
        choice = input("\n请选择操作 (1-3, 0): ").strip()
        
        if choice == "1":
            print(f"\n🚀 自动清理所有冲突会话...")
            for session in conflict_sessions:
                print(f"   正在关闭端口 {session['port']} ({session['project']})...")
                close_chrome_session(session['port'])
        
        elif choice == "2":
            print(f"\n🎯 手动选择清理:")
            for i, session in enumerate(conflict_sessions, 1):
                print(f"   {i}. 端口 {session['port']}: {session['project']}")
            
            selections = input(f"\n请输入要清理的会话编号 (用逗号分隔，如: 1,3): ").strip()
            if selections:
                try:
                    indices = [int(x.strip()) - 1 for x in selections.split(',')]
                    for idx in indices:
                        if 0 <= idx < len(conflict_sessions):
                            session = conflict_sessions[idx]
                            print(f"   正在关闭端口 {session['port']}...")
                            close_chrome_session(session['port'])
                except ValueError:
                    print("❌ 输入格式错误")
        
        elif choice == "3":
            print("ℹ️ 仅显示信息，未进行清理")
        
        else:
            print("👋 退出清理工具")
            
    except KeyboardInterrupt:
        print("\n\n👋 用户取消操作")
    except Exception as e:
        print(f"\n❌ 操作出错: {str(e)}")

def main():
    print("🔧 Chrome会话管理工具")
    print("用于解决EVP和ITC项目间的端口冲突问题")
    print("=" * 60)
    
    # 显示端口配置
    print(f"\n📋 端口配置:")
    print(f"   EVP Scorecard: {EVPChromeConfig.PORT_START}-{EVPChromeConfig.PORT_END}")
    print(f"   ITC Scorecard: {ITCChromeConfig.PORT_START}-{ITCChromeConfig.PORT_END}")
    
    interactive_cleanup()
    
    print(f"\n💡 使用建议:")
    print(f"   1. 定期运行此工具检查端口冲突")
    print(f"   2. 启用Chrome重用功能 (REUSE_EXISTING_CHROME=True)")
    print(f"   3. 确保每个项目只使用分配的端口范围")

if __name__ == "__main__":
    main()
    input("\n按回车键退出...")