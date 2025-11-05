# -*- coding: utf-8 -*-
"""
Chrome端口配置管理 - 统一多项目端口分配
用于避免多项目间的端口冲突，实现项目级别的隔离
"""

class EVPChromeConfig:
    """EVP Scorecard 项目的Chrome配置"""
    # EVP项目专用端口范围: 9222-9232 (11个端口)
    PORT_START = 9222
    PORT_END = 9232
    DEFAULT_PORT = 9222
    USER_DATA_DIR_PREFIX = "ChromeProfile_EVP"
    PROJECT_NAME = "EVP_Scorecard"


class ITCChromeConfig:
    """ITC Scorecard 项目的Chrome配置"""
    # ITC项目专用端口范围: 9233-9242 (10个端口)
    PORT_START = 9233
    PORT_END = 9242
    DEFAULT_PORT = 9233
    USER_DATA_DIR_PREFIX = "ChromeProfile_ITC"
    PROJECT_NAME = "ITC_Scorecard"


# 默认使用ITC配置（因为这是ITC项目）
DEFAULT_CONFIG = ITCChromeConfig


def get_project_config(project_name: str = "ITC_Scorecard"):
    """
    根据项目名称获取Chrome配置
    
    Args:
        project_name: 项目名称 (EVP_Scorecard, ITC_Scorecard)
        
    Returns:
        对应的配置类
    """
    if project_name == "EVP_Scorecard":
        return EVPChromeConfig
    else:
        return ITCChromeConfig


def is_port_in_range(port: int, project_name: str) -> bool:
    """
    检查端口是否在指定项目的范围内
    
    Args:
        port: 端口号
        project_name: 项目名称
        
    Returns:
        是否在范围内
    """
    config = get_project_config(project_name)
    return config.PORT_START <= port <= config.PORT_END


def get_available_ports(project_name: str) -> list:
    """
    获取项目可用的端口列表
    
    Args:
        project_name: 项目名称
        
    Returns:
        端口列表
    """
    config = get_project_config(project_name)
    return list(range(config.PORT_START, config.PORT_END + 1))


if __name__ == "__main__":
    print("Chrome端口配置管理 - 多项目统一管理")
    print("=" * 60)
    
    print(f"\n📦 EVP Scorecard 项目配置:")
    print(f"   端口范围: {EVPChromeConfig.PORT_START}-{EVPChromeConfig.PORT_END}")
    print(f"   默认端口: {EVPChromeConfig.DEFAULT_PORT}")
    print(f"   用户目录前缀: {EVPChromeConfig.USER_DATA_DIR_PREFIX}")
    
    print(f"\n📦 ITC Scorecard 项目配置:")
    print(f"   端口范围: {ITCChromeConfig.PORT_START}-{ITCChromeConfig.PORT_END}")
    print(f"   默认端口: {ITCChromeConfig.DEFAULT_PORT}")
    print(f"   用户目录前缀: {ITCChromeConfig.USER_DATA_DIR_PREFIX}")
    
    print(f"\n✅ 端口配置无冲突，可以安全运行多个项目")
    
    # 测试端口检查
    print(f"\n🔍 端口分配测试:")
    test_ports = [9222, 9232, 9233, 9242]
    for port in test_ports:
        evp_range = is_port_in_range(port, "EVP_Scorecard")
        itc_range = is_port_in_range(port, "ITC_Scorecard")
        print(f"   端口 {port}: EVP={evp_range}, ITC={itc_range}")