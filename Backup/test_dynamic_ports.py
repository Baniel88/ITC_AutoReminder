# -*- coding: utf-8 -*-
"""
测试动态端口分配功能
"""

import sys
import os
import socket

# 添加当前目录到路径，以便导入主脚本模块
script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, script_dir)

# 从主脚本导入必要的配置和函数
try:
    from Backup.ITCreport_downloader import (
        ITC_PORT_RANGE,
        is_port_available,
        allocate_debug_port,
        get_chrome_user_data_dir,
        log_message
    )
    print("✅ 成功导入主脚本模块")
except ImportError as e:
    print(f"❌ 导入失败: {e}")
    sys.exit(1)

def test_port_availability():
    """测试端口可用性检查"""
    print("\n🔍 测试端口可用性检查...")
    
    # 测试一些已知可能占用的端口
    test_ports = [80, 443, 9222, 9233, 9999]
    
    for port in test_ports:
        available = is_port_available(port)
        status = "可用" if available else "被占用"
        print(f"   端口 {port}: {status}")

def test_dynamic_port_allocation():
    """测试动态端口分配"""
    print(f"\n🎯 测试动态端口分配...")
    print(f"   ITC端口范围: {ITC_PORT_RANGE['start']}-{ITC_PORT_RANGE['end']}")
    
    allocated_ports = []
    occupied_sockets = []
    
    # 尝试分配多个端口
    for i in range(3):
        port = allocate_debug_port()
        if port:
            allocated_ports.append(port)
            print(f"   第{i+1}次分配: 端口 {port}")
            
            # 获取对应的用户数据目录
            user_data_dir = get_chrome_user_data_dir(port)
            print(f"   对应用户数据目录: {user_data_dir}")
            
            # 立即占用这个端口，以便下次分配时跳过
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.bind(('localhost', port))
                sock.listen(1)
                occupied_sockets.append(sock)
                print(f"   已占用端口 {port}，下次分配应跳过")
            except Exception as e:
                print(f"   无法占用端口 {port}: {e}")
        else:
            print(f"   第{i+1}次分配: 失败")
    
    # 释放所有占用的端口
    for sock in occupied_sockets:
        sock.close()
    
    return allocated_ports

def test_port_occupation_simulation():
    """模拟端口占用测试"""
    print(f"\n🚧 测试端口占用情况...")
    
    # 尝试占用一个ITC范围内的端口
    test_port = ITC_PORT_RANGE['start']
    
    try:
        # 创建一个socket来占用端口
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.bind(('localhost', test_port))
        sock.listen(1)
        print(f"   成功占用端口 {test_port}")
        
        # 测试端口检查功能
        available = is_port_available(test_port)
        print(f"   端口可用性检查结果: {'可用' if available else '被占用'} ✅")
        
        # 测试动态分配是否会跳过被占用的端口
        allocated_port = allocate_debug_port()
        if allocated_port and allocated_port != test_port:
            print(f"   动态分配跳过被占用端口，分配了: {allocated_port} ✅")
        else:
            print(f"   动态分配结果异常: {allocated_port} ❌")
        
        # 释放端口
        sock.close()
        print(f"   已释放端口 {test_port}")
        
    except Exception as e:
        print(f"   端口占用测试失败: {e}")

def test_range_validation():
    """测试端口范围验证"""
    print(f"\n📊 测试端口范围...")
    
    # 检查ITC端口范围的合理性
    start_port = ITC_PORT_RANGE['start']
    end_port = ITC_PORT_RANGE['end']
    range_size = end_port - start_port + 1
    
    print(f"   起始端口: {start_port}")
    print(f"   结束端口: {end_port}")
    print(f"   端口范围大小: {range_size}")
    
    if range_size >= 5:
        print("   ✅ 端口范围大小合理（>=5）")
    else:
        print("   ⚠️ 端口范围较小，可能不够用")
    
    # 检查是否与EVP范围冲突（假设EVP使用9222-9232）
    evp_start, evp_end = 9222, 9232
    if start_port > evp_end or end_port < evp_start:
        print(f"   ✅ 与EVP端口范围({evp_start}-{evp_end})无冲突")
    else:
        print(f"   ⚠️ 与EVP端口范围({evp_start}-{evp_end})有重叠")

if __name__ == "__main__":
    print("🧪 动态端口分配功能测试")
    print("=" * 50)
    
    # 运行各项测试
    test_port_availability()
    test_dynamic_port_allocation()
    test_port_occupation_simulation()
    test_range_validation()
    
    print("\n" + "=" * 50)
    print("🎉 测试完成！")