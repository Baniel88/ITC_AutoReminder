# -*- coding: utf-8 -*-
"""
多项目Chrome端口分配配置

为了避免多个项目同时使用Chrome时的端口冲突，
这里定义了各个项目的专用端口范围。

端口分配策略：
- 每个项目分配10个连续端口
- 如果项目需要更多实例，可以扩展端口范围
"""

# 项目端口分配表
PROJECT_PORT_ALLOCATION = {
    # EVP Scorecard 项目（已存在）
    "EVP_Scorecard": {
        "start_port": 9222,
        "end_port": 9232,
        "description": "EVP合规评分卡项目",
        "status": "已占用"
    },
    
    # ITC Scorecard 项目（当前项目）
    "ITC_Scorecard": {
        "start_port": 9233,
        "end_port": 9242,
        "description": "ITC运输评分卡项目",
        "status": "当前项目"
    },
    
    # 预留的其他项目端口
    "Project_3": {
        "start_port": 9243,
        "end_port": 9252,
        "description": "预留项目3",
        "status": "预留"
    },
    
    "Project_4": {
        "start_port": 9253,
        "end_port": 9262,
        "description": "预留项目4",
        "status": "预留"
    },
    
    "Project_5": {
        "start_port": 9263,
        "end_port": 9272,
        "description": "预留项目5",
        "status": "预留"
    }
}

# 当前项目配置
CURRENT_PROJECT = "ITC_Scorecard"
CURRENT_PROJECT_CONFIG = PROJECT_PORT_ALLOCATION[CURRENT_PROJECT]

# ITC项目专用配置
class ITCChromeConfig:
    """ITC项目Chrome配置类"""
    
    # 端口配置
    PORT_START = CURRENT_PROJECT_CONFIG["start_port"]
    PORT_END = CURRENT_PROJECT_CONFIG["end_port"]
    DEFAULT_PORT = PORT_START  # 默认使用起始端口
    
    # 目录配置
    USER_DATA_DIR_PREFIX = "ChromeProfile_ITC"
    PROJECT_NAME = "ITC_Scorecard"
    
    # Chrome选项
    DEFAULT_CHROME_OPTIONS = [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-web-security",
        "--disable-features=VizDisplayCompositor",
        "--disable-extensions",
        "--disable-plugins",
        "--start-maximized"
    ]
    
    # 调试选项
    DEBUG_OPTIONS = [
        "--enable-logging",
        "--log-level=0",
        "--disable-background-timer-throttling",
        "--disable-backgrounding-occluded-windows",
        "--disable-renderer-backgrounding"
    ]
    
    @classmethod
    def get_project_info(cls):
        """获取项目信息"""
        return {
            "project_name": cls.PROJECT_NAME,
            "port_range": f"{cls.PORT_START}-{cls.PORT_END}",
            "available_ports": cls.PORT_END - cls.PORT_START + 1,
            "user_data_prefix": cls.USER_DATA_DIR_PREFIX
        }
    
    @classmethod
    def get_chrome_command_template(cls, port, user_data_dir):
        """获取Chrome启动命令模板"""
        return [
            "chrome.exe",  # 将被实际路径替换
            f"--remote-debugging-port={port}",
            f"--user-data-dir={user_data_dir}",
            *cls.DEFAULT_CHROME_OPTIONS
        ]


def check_port_conflicts():
    """检查端口分配是否有冲突"""
    used_ports = set()
    conflicts = []
    
    for project_name, config in PROJECT_PORT_ALLOCATION.items():
        port_range = set(range(config["start_port"], config["end_port"] + 1))
        
        # 检查是否与已使用的端口重叠
        overlap = used_ports & port_range
        if overlap:
            conflicts.append({
                "project": project_name,
                "conflicted_ports": list(overlap)
            })
        
        used_ports.update(port_range)
    
    return conflicts


def get_available_port_for_project(project_name):
    """获取指定项目的可用端口"""
    if project_name not in PROJECT_PORT_ALLOCATION:
        return None
    
    config = PROJECT_PORT_ALLOCATION[project_name]
    
    # 尝试从起始端口开始查找可用端口
    import socket
    
    for port in range(config["start_port"], config["end_port"] + 1):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(('localhost', port))
                return port
        except OSError:
            continue
    
    return None


def print_port_allocation_table():
    """打印端口分配表"""
    print("=" * 80)
    print("Chrome远程调试端口分配表")
    print("=" * 80)
    print(f"{'项目名称':<20} {'端口范围':<15} {'状态':<10} {'描述'}")
    print("-" * 80)
    
    for project_name, config in PROJECT_PORT_ALLOCATION.items():
        port_range = f"{config['start_port']}-{config['end_port']}"
        print(f"{project_name:<20} {port_range:<15} {config['status']:<10} {config['description']}")
    
    print("=" * 80)
    print(f"总计端口数量: {sum(config['end_port'] - config['start_port'] + 1 for config in PROJECT_PORT_ALLOCATION.values())}")
    
    # 检查冲突
    conflicts = check_port_conflicts()
    if conflicts:
        print("\n⚠️ 发现端口冲突:")
        for conflict in conflicts:
            print(f"  项目 {conflict['project']}: 端口 {conflict['conflicted_ports']}")
    else:
        print("\n✅ 端口分配无冲突")


if __name__ == "__main__":
    # 显示端口分配信息
    print_port_allocation_table()
    
    # 显示当前项目信息
    print(f"\n当前项目: {CURRENT_PROJECT}")
    print(f"端口范围: {CURRENT_PROJECT_CONFIG['start_port']}-{CURRENT_PROJECT_CONFIG['end_port']}")
    
    # 检查当前项目的可用端口
    available_port = get_available_port_for_project(CURRENT_PROJECT)
    if available_port:
        print(f"建议使用端口: {available_port}")
    else:
        print("⚠️ 当前项目端口范围内没有可用端口")
    
    # 显示ITC项目配置
    print(f"\nITC项目配置:")
    info = ITCChromeConfig.get_project_info()
    for key, value in info.items():
        print(f"  {key}: {value}")