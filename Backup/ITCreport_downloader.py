# -*- coding: utf-8 -*-
"""
ITC报表自动下载器（简化登录检测，优化日期参数）

邮件发送控制说明：
1. 主文件配置（第75-76行）：
   - EMAIL_ENABLED: True/False 控制是否启用邮件功能
   - EMAIL_AUTO_SEND: True/False 控制是否自动发送邮件
   
2. 配置文件覆盖（email_config.json）：
   - "EMAIL_ENABLED": true/false 可覆盖主文件设置
   - "AUTO_SEND_EMAIL": true/false 可覆盖主文件设置
   
3. 推荐设置：
   - EMAIL_ENABLED = True, EMAIL_AUTO_SEND = False (预览后发送)
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import os
import shutil
import subprocess
import platform
import urllib.parse
import requests
import socket
import re
import zipfile
import sys
import socket
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import json
from bs4 import BeautifulSoup
import warnings


# -------------------------- 核心配置 --------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_PROCESSOR_SCRIPT_NAME = "pending_review_report.py"
REPORT_PROCESSOR_SCRIPT = os.path.join(SCRIPT_DIR, REPORT_PROCESSOR_SCRIPT_NAME)

# 报表下载参数（动态计算日期范围）
BASE_REPORT_URL = "https://itc-tool.pg.com/RequestReport/GetRequestExportReport"
ITC_LOGIN_URL = "https://itc-tool.pg.com/"  # ITC系统登录页面
TIME_RANGE_DAYS = 365
end_date = datetime.now()
start_date = end_date - timedelta(days=TIME_RANGE_DAYS)
REPORT_PARAMS = {
    "siteId": "193",
    "areaId": "-1",
    "systemId": "-1",
    "categoryId": "-1",
    "requestStatus": "8",
    "requestedForId": "",
    "dateRange": f"{start_date.strftime('%m/%d/%Y')} - {end_date.strftime('%m/%d/%Y')}",
    "accessTypeId": "-1"
}
REPORT_URL = f"{BASE_REPORT_URL}?{urllib.parse.urlencode(REPORT_PARAMS)}"

# 目录设置
ITC_REPORT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, "ITC report"))
RAW_DATA_DIR = os.path.join(ITC_REPORT_DIR, "RawData")
REMINDER_DIR = os.path.join(ITC_REPORT_DIR, "Reminder")
HTML_REPORT_DIR = os.path.join(ITC_REPORT_DIR, "HTML Reports")
LOG_DIR = os.path.join(ITC_REPORT_DIR, "Log")

# 动态端口配置
# ITC项目专用端口范围：9233-9242（避免与EVP项目9222-9232冲突）
ITC_PORT_RANGE = {
    "start": 9233,
    "end": 9242,
    "project_name": "ITC_Scorecard"
}

# 当前使用的调试端口（将动态分配）
DEBUG_PORT = None

# Chrome配置
CHROME_USER_DATA_DIR = None  # 将动态分配

# 运行参数
LOGIN_TIMEOUT = 300  # 登录超时时间（5分钟）
LOGIN_CHECK_INTERVAL = 5
DOWNLOAD_TIMEOUT = 600
POST_DOWNLOAD_WAIT = 10
SCRIPT_CALL_TIMEOUT = 600

# 邮件发送控制（主配置）
EMAIL_AUTO_SEND = False  # True=直接发送, False=预览后发送（推荐）
EMAIL_ENABLED = True     # True=启用邮件功能, False=完全禁用邮件

# 简化登录检测元素（移除可能失效的XPATH）
LOGGED_IN_ELEMENTS = [
    By.CSS_SELECTOR, "#frmRequestAccess",  # 保留可靠的CSS选择器
    By.LINK_TEXT, "退出登录"  # 保留退出登录链接检测
]
# ITC系统域名（用于URL验证）
ITC_DOMAIN = "itc-tool.pg.com"


# 全局变量：驱动可执行文件名
DRIVER_EXECUTABLE = "chromedriver.exe" if os.name == 'nt' else "chromedriver"

# 网络配置：解决SSL问题
ALLOW_INSECURE_SSL = True  # 当有自签名证书时启用


# -------------------------- 辅助函数 --------------------------
def get_email_settings():
    """
    获取邮件发送设置（支持主文件配置和配置文件覆盖）
    返回: (email_enabled, auto_send)
    """
    # 默认使用主文件中的配置
    email_enabled = EMAIL_ENABLED
    auto_send = EMAIL_AUTO_SEND
    
    # 尝试从配置文件读取覆盖设置
    try:
        config_path = os.path.join(SCRIPT_DIR, "email_config.json")
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                system_config = config.get('system_config', {})
                
                # 配置文件可以覆盖主文件设置
                if 'EMAIL_ENABLED' in system_config:
                    email_enabled = system_config['EMAIL_ENABLED']
                if 'AUTO_SEND_EMAIL' in system_config:
                    auto_send = system_config['AUTO_SEND_EMAIL']
                    
                log_message(f"📧 邮件配置: 启用={email_enabled}, 自动发送={auto_send} (来源: 配置文件)")
        else:
            log_message(f"📧 邮件配置: 启用={email_enabled}, 自动发送={auto_send} (来源: 主文件)")
            
    except Exception as e:
        log_message(f"⚠️ 读取邮件配置文件失败，使用主文件默认设置: {str(e)}")
        log_message(f"📧 邮件配置: 启用={email_enabled}, 自动发送={auto_send} (来源: 主文件默认)")
    
    return email_enabled, auto_send


def ensure_directory_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)
        log_message(f"已创建文件夹: {directory}")
    return directory


def log_message(message):
    ensure_directory_exists(LOG_DIR)
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_date = datetime.now().strftime('%Y%m%d')
    log_file = os.path.join(LOG_DIR, f"itc_downloader_{log_date}.log")
    
    print(f"[{timestamp}] {message}")
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {message}\n")


def pre_check_report_script():
    log_message("="*60)
    log_message("📌 开始预检查报表处理脚本")
    log_message(f"当前脚本目录: {SCRIPT_DIR}")
    log_message(f"待调用报表脚本: {REPORT_PROCESSOR_SCRIPT_NAME}")
    log_message(f"报表脚本完整路径: {REPORT_PROCESSOR_SCRIPT}")
    
    if os.path.exists(REPORT_PROCESSOR_SCRIPT) and os.access(REPORT_PROCESSOR_SCRIPT, os.R_OK):
        log_message("✅ 报表处理脚本存在且可读，预检查通过")
        return True
    else:
        log_message(f"❌ 报表处理脚本不存在或不可读: {REPORT_PROCESSOR_SCRIPT}")
        return False


def analyze_report_name(downloaded_path):
    """分析报表文件名，提取关键信息"""
    log_message("\n===== 开始分析报表名称 =====")
    
    try:
        # 获取文件名和扩展名
        file_name = os.path.basename(downloaded_path)
        file_base, file_ext = os.path.splitext(file_name)
        
        log_message(f"原始文件名: {file_name}")
        log_message(f"文件名(无扩展名): {file_base}")
        log_message(f"文件扩展名: {file_ext}")
        
        # 分析文件名结构
        report_info = {
            "full_name": file_name,
            "base_name": file_base,
            "extension": file_ext,
            "path": downloaded_path,
            "size": os.path.getsize(downloaded_path),
            "download_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        # 尝试从文件名中提取日期信息
        date_patterns = [
            r"(\d{4})(\d{2})(\d{2})",  # YYYYMMDD
            r"(\d{2})(\d{2})(\d{4})",  # MMDDYYYY
            r"(\d{2})-(\d{2})-(\d{4})",  # MM-DD-YYYY
            r"(\d{4})-(\d{2})-(\d{2})"   # YYYY-MM-DD
        ]
        
        for pattern in date_patterns:
            match = re.search(pattern, file_base)
            if match:
                report_info["date_pattern"] = pattern
                report_info["date_match"] = match.group(0)
                break
        
        # 分析文件名中的关键词
        keywords = []
        if "request" in file_base.lower():
            keywords.append("请求")
        if "report" in file_base.lower():
            keywords.append("报表")
        if "export" in file_base.lower():
            keywords.append("导出")
        if "itc" in file_base.lower():
            keywords.append("ITC系统")
        if "pending" in file_base.lower():
            keywords.append("待处理")
        if "review" in file_base.lower():
            keywords.append("审核")
        
        report_info["keywords"] = keywords
        
        log_message(f"提取的报表信息: {json.dumps(report_info, ensure_ascii=False, indent=2)}")
        log_message("===== 报表名称分析完成 =====")
        
        return report_info
        
    except Exception as e:
        log_message(f"❌ 分析报表名称时出错: {str(e)}")
        return None


def generate_html_report(report_info, report_params):
    """生成HTML格式的报表分析报告"""
    log_message("\n===== 开始生成HTML报告 =====")
    
    try:
        ensure_directory_exists(HTML_REPORT_DIR)
        
        # 创建报告文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        html_file_name = f"report_analysis_{timestamp}.html"
        html_file_path = os.path.join(HTML_REPORT_DIR, html_file_name)
        
        # 构建HTML内容
        html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ITC报表分析报告 - {timestamp}</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }}
        .container {{
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            padding: 30px;
        }}
        .header {{
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 2px solid #4CAF50;
        }}
        .header h1 {{
            color: #2c3e50;
            margin-bottom: 10px;
        }}
        .header .timestamp {{
            color: #7f8c8d;
            font-style: italic;
        }}
        .section {{
            margin-bottom: 30px;
        }}
        .section-title {{
            color: #2c3e50;
            border-left: 4px solid #4CAF50;
            padding-left: 15px;
            margin-bottom: 15px;
            font-size: 1.2em;
        }}
        .info-table {{
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
        }}
        .info-table th, .info-table td {{
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #ddd;
        }}
        .info-table th {{
            background-color: #f8f9fa;
            font-weight: bold;
            color: #2c3e50;
        }}
        .info-table tr:hover {{
            background-color: #f5f5f5;
        }}
        .keyword {{
            display: inline-block;
            background-color: #e8f5e9;
            color: #2e7d32;
            padding: 5px 10px;
            border-radius: 20px;
            margin-right: 8px;
            margin-bottom: 8px;
            font-size: 0.9em;
        }}
        .footer {{
            text-align: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid #ddd;
            color: #7f8c8d;
            font-size: 0.9em;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>ITC报表分析报告</h1>
            <p class="timestamp">生成时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
        </div>
        
        <div class="section">
            <h2 class="section-title">报表基本信息</h2>
            <table class="info-table">
                <tr>
                    <th>项目</th>
                    <th>值</th>
                </tr>
                <tr>
                    <td>报表文件名</td>
                    <td>{report_info.get('full_name', 'N/A')}</td>
                </tr>
                <tr>
                    <td>文件路径</td>
                    <td>{report_info.get('path', 'N/A')}</td>
                </tr>
                <tr>
                    <td>文件大小</td>
                    <td>{format_file_size(report_info.get('size', 0))}</td>
                </tr>
                <tr>
                    <td>下载时间</td>
                    <td>{report_info.get('download_time', 'N/A')}</td>
                </tr>
                <tr>
                    <td>文件类型</td>
                    <td>{report_info.get('extension', 'N/A')}</td>
                </tr>
                <tr>
                    <td>日期信息</td>
                    <td>{report_info.get('date_match', '未找到日期信息')}</td>
                </tr>
            </table>
        </div>
        
        <div class="section">
            <h2 class="section-title">报表关键词分析</h2>
            <div class="keywords">
                {''.join([f'<span class="keyword">{kw}</span>' for kw in report_info.get('keywords', [])])}
                {'' if report_info.get('keywords') else '<p>未提取到关键词</p>'}
            </div>
        </div>
        
        <div class="section">
            <h2 class="section-title">下载参数信息</h2>
            <table class="info-table">
                <tr>
                    <th>参数名</th>
                    <th>参数值</th>
                </tr>"""
        
        # 添加报表参数信息
        for param_name, param_value in report_params.items():
            html_content += f"""
                <tr>
                    <td>{param_name}</td>
                    <td>{param_value}</td>
                </tr>"""
        
        html_content += f"""
            </table>
        </div>
        
        <div class="section">
            <h2 class="section-title">报表分析结论</h2>
            <div class="analysis">
                <p>根据文件名分析，该报表可能是：</p>
                <ul>
                    <li>{get_report_type_analysis(report_info)}</li>
                    <li>数据时间范围：{report_params.get('dateRange', 'N/A')}</li>
                    <li>来源系统：ITC系统</li>
                </ul>
            </div>
        </div>
        
        <div class="footer">
            <p>ITC报表自动下载器生成 | 报告生成时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
        </div>
    </div>
</body>
</html>"""
        
        # 写入HTML文件
        with open(html_file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        log_message(f"✅ HTML报告已生成: {html_file_path}")
        log_message("===== HTML报告生成完成 =====")
        
        return html_file_path
        
    except Exception as e:
        log_message(f"❌ 生成HTML报告时出错: {str(e)}")
        return None


def format_file_size(size_bytes):
    """格式化文件大小"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.2f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


def get_report_type_analysis(report_info):
    """根据报表信息分析报表类型"""
    keywords = report_info.get('keywords', [])
    base_name = report_info.get('base_name', '').lower()
    
    if 'pending' in base_name or 'review' in base_name:
        return "待审核请求报表 - 包含需要审核的系统请求记录"
    elif 'request' in base_name:
        return "系统请求报表 - 包含系统操作请求记录"
    elif 'export' in base_name:
        return "数据导出报表 - 系统数据导出结果"
    else:
        return "ITC系统报表 - 具体类型需要进一步分析内容"


def send_email_notification(report_info, html_report_path=None):
    """发送邮件通知，包含报表分析信息"""
    log_message("\n===== 开始发送邮件通知 =====")
    
    try:
        # 检查是否有smtplib模块
        try:
            import smtplib
            from email.mime.text import MIMEText
            from email.mime.multipart import MIMEMultipart
            from email.mime.application import MIMEApplication
        except ImportError:
            log_message("⚠️ 无法导入邮件模块，跳过邮件发送")
            return False
        
        # 邮件配置（需要根据实际情况修改）
        SMTP_SERVER = "smtp.example.com"
        SMTP_PORT = 587
        SMTP_USER = "your_email@example.com"
        SMTP_PASSWORD = "your_password"
        RECIPIENTS = ["recipient1@example.com", "recipient2@example.com"]
        
        # 创建邮件
        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = ", ".join(RECIPIENTS)
        msg['Subject'] = f"ITC报表下载完成通知 - {report_info.get('full_name', '未知报表')}"
        
        # 邮件正文
        body = f"""
        <html>
        <body>
            <h2>ITC报表下载完成通知</h2>
            <p>报表下载和分析已完成，以下是详细信息：</p>
            
            <h3>报表基本信息</h3>
            <ul>
                <li><strong>报表名称：</strong>{report_info.get('full_name', 'N/A')}</li>
                <li><strong>文件大小：</strong>{format_file_size(report_info.get('size', 0))}</li>
                <li><strong>下载时间：</strong>{report_info.get('download_time', 'N/A')}</li>
                <li><strong>存储路径：</strong>{report_info.get('path', 'N/A')}</li>
            </ul>
            
            <h3>报表类型分析</h3>
            <p>{get_report_type_analysis(report_info)}</p>
            
            <h3>关键词标签</h3>
            <p>{', '.join(report_info.get('keywords', ['无']))}</p>
            
            <p>详细分析报告请查看附件。</p>
            
            <p>此致<br>ITC报表自动下载系统</p>
        </body>
        </html>
        """
        
        msg.attach(MIMEText(body, 'html', 'utf-8'))
        
        # 添加HTML报告附件
        if html_report_path and os.path.exists(html_report_path):
            with open(html_report_path, 'rb') as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(html_report_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(html_report_path)}"'
                msg.attach(part)
        
        # 发送邮件
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        
        log_message(f"✅ 邮件已发送至: {', '.join(RECIPIENTS)}")
        log_message("===== 邮件通知发送完成 =====")
        
        return True
        
    except Exception as e:
        log_message(f"❌ 发送邮件时出错: {str(e)}")
        log_message("⚠️ 请检查邮件配置是否正确")
        return False


def call_report_processor(downloaded_csv_path):
    log_message(f"\n===== 开始调用报表处理脚本: {REPORT_PROCESSOR_SCRIPT_NAME} =====")
    log_message(f"待处理CSV文件: {downloaded_csv_path}")
    
    if not os.path.exists(REPORT_PROCESSOR_SCRIPT) or not os.path.exists(downloaded_csv_path):
        log_message("❌ 脚本或CSV文件不存在，无法调用")
        return False
    
    try:
        cmd = [sys.executable, REPORT_PROCESSOR_SCRIPT, "--csv-path", downloaded_csv_path, "--log-dir", LOG_DIR]
        result = subprocess.run(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            text=True, timeout=SCRIPT_CALL_TIMEOUT, cwd=SCRIPT_DIR
        )
        
        if result.stdout:
            log_message(f"📤 报表脚本输出:\n{result.stdout}")
        if result.stderr:
            log_message(f"⚠️ 报表脚本错误:\n{result.stderr}")
        
        return result.returncode == 0
    except Exception as e:
        log_message(f"❌ 调用脚本异常: {str(e)}")
        return False
    finally:
        log_message(f"===== 报表处理脚本调用结束 =====")


# -------------------------- ChromeDriver管理 --------------------------
# 导入Chrome Driver管理器
try:
    from Chrome_Driver_mgr import ChromeDriverManager, get_chromedriver_path as get_driver_path, get_chrome_path as get_browser_path
    CHROME_DRIVER_MGR_AVAILABLE = True
    log_message("✅ 成功导入Chrome Driver管理器")
    
    # 初始化Chrome Driver管理器
    chrome_driver_manager = ChromeDriverManager(
        script_dir=SCRIPT_DIR,
        allow_insecure_ssl=ALLOW_INSECURE_SSL,
        log_callback=log_message
    )
except ImportError as e:
    log_message(f"⚠️ 无法导入Chrome Driver管理器: {str(e)}")
    log_message("将使用简化的Driver管理功能")
    CHROME_DRIVER_MGR_AVAILABLE = False
    chrome_driver_manager = None

# 动态端口管理函数
def is_port_available(port):
    """检查端口是否可用"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(2)  # 增加超时时间
            # 尝试绑定端口，如果成功则端口可用
            sock.bind(('localhost', port))
            return True
    except (socket.error, OSError) as e:
        # 端口被占用或其他错误
        return False
    except Exception as e:
        log_message(f"检查端口 {port} 时出错: {str(e)}")
        return False

def allocate_debug_port():
    """动态分配可用的调试端口"""
    for port in range(ITC_PORT_RANGE['start'], ITC_PORT_RANGE['end'] + 1):
        if is_port_available(port):
            log_message(f"✅ 分配端口: {port}")
            return port
    
    # 如果ITC范围内没有可用端口，记录警告
    log_message(f"⚠️ ITC端口范围 {ITC_PORT_RANGE['start']}-{ITC_PORT_RANGE['end']} 内无可用端口")
    return None

def get_chrome_user_data_dir(debug_port):
    """获取项目特定的Chrome用户数据目录"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    user_data_dir = os.path.join(script_dir, f"chrome_debug_profile_itc_{debug_port}")
    ensure_directory_exists(user_data_dir)
    return user_data_dir

def get_chromedriver_path():
    """获取ChromeDriver路径"""
    if CHROME_DRIVER_MGR_AVAILABLE and chrome_driver_manager:
        return chrome_driver_manager.get_chromedriver_path()
    else:
        # 备用简化版本
        script_dir = os.path.dirname(os.path.abspath(__file__))
        driver_path = os.path.join(script_dir, DRIVER_EXECUTABLE)
        if os.path.exists(driver_path):
            log_message(f"找到ChromeDriver: {driver_path}")
            return driver_path
        else:
            log_message(f"❌ 未找到ChromeDriver: {driver_path}")
            return None


# -------------------------- 浏览器管理 --------------------------
def get_chrome_path():
    """获取Chrome浏览器路径"""
    if CHROME_DRIVER_MGR_AVAILABLE and chrome_driver_manager:
        return chrome_driver_manager.get_chrome_path()
    else:
        # 备用简化版本
        system = platform.system()
        log_message(f"检测到操作系统: {system}")
        
        if system == "Windows":
            try:
                import winreg
                reg_paths = [
                    r"Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                    r"Software\Wow6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"
                ]
                for reg_path in reg_paths:
                    try:
                        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
                        value, _ = winreg.QueryValueEx(key, "")
                        winreg.CloseKey(key)
                        if value and os.path.exists(value):
                            log_message(f"从注册表找到Chrome路径: {value}")
                            return value
                    except:
                        continue
            except Exception as e:
                log_message(f"从注册表查找Chrome失败: {str(e)}")
            
            common_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            ]
            for path in common_paths:
                if os.path.exists(path):
                    log_message(f"找到Chrome路径: {path}")
                    return path
            return common_paths[0]
        
        elif system == "Darwin":
            chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
            return chrome_path if os.path.exists(chrome_path) else chrome_path
        else:
            chrome_path = "/usr/bin/google-chrome"
            return chrome_path if os.path.exists(chrome_path) else chrome_path

CHROME_PATH = get_chrome_path()


def start_chrome_debug_session():
    global DEBUG_PORT, CHROME_USER_DATA_DIR
    
    if not os.path.exists(CHROME_PATH):
        log_message(f"❌ Chrome路径无效: {CHROME_PATH}")
        return False

    # 动态分配端口
    allocated_port = allocate_debug_port()
    if allocated_port is None:
        log_message("❌ 无法分配可用端口")
        return False
    
    DEBUG_PORT = allocated_port
    CHROME_USER_DATA_DIR = get_chrome_user_data_dir(DEBUG_PORT)
    
    # 检查端口是否已被占用
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            if sock.connect_ex(("127.0.0.1", DEBUG_PORT)) == 0:
                log_message(f"✅ 检测到已运行的调试模式Chrome (端口: {DEBUG_PORT})")
                return True
    except Exception as e:
        log_message(f"⚠️ 检查端口时出错: {str(e)}")

    # 启动新的调试模式Chrome
    try:
        log_message(f"创建Chrome调试配置文件目录: {CHROME_USER_DATA_DIR}")
        
        chrome_args = [
            CHROME_PATH,
            f"--remote-debugging-port={DEBUG_PORT}",
            f"--user-data-dir={CHROME_USER_DATA_DIR}",
            "--no-first-run",
            "--no-default-browser-check",
            ITC_LOGIN_URL  # 启动时打开ITC系统登录页面
        ]
        
        subprocess.Popen(chrome_args, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        log_message("🔄 正在启动Chrome...")
        time.sleep(8)
        log_message("✅ Chrome已启动，请在浏览器中登录ITC系统")
        return True
    except Exception as e:
        log_message(f"❌ 启动Chrome失败: {str(e)}")
        return False


# -------------------------- 登录检测（优化版） --------------------------
def is_itc_logged_in(driver):
    try:
        driver.refresh()
        log_message("刷新页面以检测登录状态")
        time.sleep(2)
        
        # 1. 优先通过URL验证：如果当前URL包含ITC系统域名，直接认为已登录
        current_url = driver.current_url
        if ITC_DOMAIN in current_url:
            log_message(f"✅ 检测到ITC系统URL: {current_url}，确认已登录")
            return True
        
        # 2. 检查核心登录元素（简化后的元素列表）
        for i in range(0, len(LOGGED_IN_ELEMENTS), 2):
            by_type = LOGGED_IN_ELEMENTS[i]
            selector = LOGGED_IN_ELEMENTS[i+1]
            try:
                element = driver.find_element(by_type, selector)
                if element.is_displayed():
                    log_message(f"✅ 检测到登录元素: {by_type}={selector}")
                    return True
            except:
                log_message(f"未检测到登录元素: {by_type}={selector}")
                continue
        
        log_message(f"ℹ️ 未检测到登录状态（当前URL: {current_url}）")
        return False
    except Exception as e:
        log_message(f"⚠️ 登录检测异常: {str(e)}")
        return False


def wait_for_itc_login():
    chromedriver_path = get_chromedriver_path()
    if not chromedriver_path or not os.path.exists(chromedriver_path):
        log_message("❌ 无法获取有效ChromeDriver路径")
        return None
    
    driver = None
    try:
        options = Options()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
        
        prefs = {
            "download.default_directory": RAW_DATA_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False
        }
        options.add_experimental_option("prefs", prefs)
        log_message(f"设置Chrome下载目录为: {RAW_DATA_DIR}")
        
        if not os.environ.get('DISPLAY') and platform.system() != 'Windows':
            options.add_argument("--headless=new")
            options.add_argument("--no-sandbox")
            log_message("启用无头模式运行Chrome")
        
        driver = webdriver.Chrome(service=Service(chromedriver_path), options=options)
        log_message("✅ 成功连接到Chrome调试会话")
        
        start_time = time.time()
        while time.time() - start_time < LOGIN_TIMEOUT:
            if is_itc_logged_in(driver):
                return driver
            
            elapsed = int(time.time() - start_time)
            remaining = max(0, LOGIN_TIMEOUT - elapsed)
            log_message(f"⏳ 等待登录中...（已等待 {elapsed} 秒，剩余 {remaining} 秒）")
            time.sleep(LOGIN_CHECK_INTERVAL)
        
        log_message(f"❌ 登录超时（超过{LOGIN_TIMEOUT}秒）")
        driver.quit()
        return None
    except Exception as e:
        log_message(f"❌ 连接Chrome失败: {str(e)}")
        if driver:
            driver.quit()
        return None


# -------------------------- 报表下载与处理 --------------------------
def download_report(driver):
    log_message(f"\n===== 开始下载任务: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} =====")
    log_message(f"📅 日期范围: {start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}")
    log_message(f"🎯 目标目录: {RAW_DATA_DIR}")
    log_message(f"📤 下载请求URL: {REPORT_URL}")
    
    download_paths = [
        RAW_DATA_DIR,
        os.path.join(os.environ.get("USERPROFILE", ""), "Downloads"),
        os.path.join(os.environ.get("HOME", ""), "Downloads")
    ]
    initial_files = {}
    for path in download_paths:
        if os.path.exists(path):
            initial_files[path] = {
                f: os.path.getmtime(os.path.join(path, f)) 
                for f in os.listdir(path) 
                if os.path.isfile(os.path.join(path, f)) and not f.endswith((".crdownload", ".tmp"))
            }
            log_message(f"记录初始文件: {path}（共{len(initial_files[path])}个文件）")
    
    try:
        task_start = time.time()
        driver.get(REPORT_URL)  # 直接访问下载链接
        log_message("🔄 已向ITC系统发送报表下载请求")
        
        downloaded_path = None
        while time.time() - task_start < DOWNLOAD_TIMEOUT:
            for path in download_paths:
                if not os.path.exists(path):
                    continue
                
                for f in os.listdir(path):
                    file_path = os.path.join(path, f)
                    if (os.path.isfile(file_path) and 
                        not f.endswith((".crdownload", ".tmp")) and 
                        os.path.getsize(file_path) > 1024):
                        
                        is_new = (path not in initial_files or 
                                 f not in initial_files[path] or 
                                 os.path.getmtime(file_path) > task_start)
                        
                        if is_new:
                            downloaded_path = file_path
                            log_message(f"✅ 发现新下载文件: {downloaded_path}")
                            break
            
            if downloaded_path:
                break
            
            elapsed = int(time.time() - task_start)
            if elapsed % 15 == 0:
                log_message(f"   等待下载中...（已等待 {elapsed} 秒，超时时间 {DOWNLOAD_TIMEOUT} 秒）")
            time.sleep(3)
        
        if not downloaded_path:
            log_message("❌ 报表下载超时，未找到下载文件")
            return False, None
        
        if not downloaded_path.startswith(RAW_DATA_DIR):
            target_name = os.path.basename(downloaded_path)
            target_path = os.path.join(RAW_DATA_DIR, target_name)
            if os.path.exists(target_path):
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                name, ext = os.path.splitext(target_name)
                target_path = os.path.join(RAW_DATA_DIR, f"{name}_{timestamp}{ext}")
            
            shutil.move(downloaded_path, target_path)
            log_message(f"已将文件从 {downloaded_path} 移动到 {target_path}")
            downloaded_path = target_path
        
        file_size = os.path.getsize(downloaded_path) / (1024 * 1024)
        log_message(f"\n✅ 报表下载成功！")
        log_message(f"📄 文件名: {os.path.basename(downloaded_path)}")
        log_message(f"📁 保存路径: {downloaded_path}")
        log_message(f"📊 文件大小: {file_size:.2f} MB")
        return True, downloaded_path
        
    except Exception as e:
        log_message(f"❌ 下载过程异常: {str(e)}")
        return False, None
    finally:
        log_message(f"===== 下载任务结束 =====")


# -------------------------- 主程序 --------------------------
if __name__ == "__main__":
    start_time = datetime.now()
    log_message("="*60)
    log_message("          ITC报表下载器与处理系统（增强版）")
    log_message(f"          启动时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    log_message("="*60)
    
    if not pre_check_report_script():
        log_message("❌ 预检查失败，程序终止")
        sys.exit(1)
    
    log_message(f"📌 数据范围: 近{TIME_RANGE_DAYS}天（{start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')}）")
    log_message(f"📌 原始数据目录: {RAW_DATA_DIR}")
    log_message(f"📌 报告输出目录: {REMINDER_DIR}")
    log_message(f"📌 日志目录: {LOG_DIR}")
    log_message(f"📌 登录超时: {LOGIN_TIMEOUT // 60}分钟")
    log_message("="*60)
    
    try:
        # 1. 获取ChromeDriver
        chromedriver_path = get_chromedriver_path()
        if not chromedriver_path or not os.path.exists(chromedriver_path):
            log_message("❌ 程序终止: 无法获取ChromeDriver")
            log_message("💡 手动下载指引:")
            log_message("   1. 访问官方下载页: https://googlechromelabs.github.io/chrome-for-testing/")
            log_message("   2. 找到与Chrome版本匹配的ChromeDriver")
            log_message("   3. 下载对应系统版本（Windows 64位选择win64）")
            log_message(f"   4. 解压后将{DRIVER_EXECUTABLE}放在脚本目录: {SCRIPT_DIR}")
            sys.exit(1)
        
        # 2. 检查Chrome路径
        if not os.path.exists(CHROME_PATH):
            log_message(f"❌ 程序终止: Chrome未找到（路径: {CHROME_PATH}）")
            sys.exit(1)
        
        # 3. 启动Chrome调试模式
        if not start_chrome_debug_session():
            log_message("❌ 程序终止: 无法启动Chrome调试模式")
            sys.exit(1)
        
        # 4. 等待用户登录
        driver = wait_for_itc_login()
        if not driver:
            log_message("❌ 程序终止: 登录失败或超时")
            sys.exit(1)
        
        # 5. 下载报表
        download_success, downloaded_path = download_report(driver)
        driver.quit()
        
        # 6. 报表分析和处理
        if download_success and downloaded_path:
            log_message(f"\n⏳ 等待{POST_DOWNLOAD_WAIT}秒，确保文件写入完成...")
            time.sleep(POST_DOWNLOAD_WAIT)
            
            # 6.1 分析报表名称
            report_info = analyze_report_name(downloaded_path)
            
            # 6.2 生成HTML报告
            html_report_path = None
            if report_info:
                html_report_path = generate_html_report(report_info, REPORT_PARAMS)
            
            # 6.3 调用报表处理脚本
            log_message("\n===== 开始处理下载的报表数据 =====")
            if call_report_processor(downloaded_path):
                log_message("✅ 报表处理完成！")
            else:
                log_message("⚠️ 报表处理脚本调用失败，请手动运行脚本处理:")
                log_message(f"   脚本路径: {REPORT_PROCESSOR_SCRIPT}")
                log_message(f"   数据路径: {downloaded_path}")
            
            # 6.4 发送邮件通知
            if report_info:
                email_enabled, auto_send = get_email_settings()
                
                if not email_enabled:
                    log_message("📧 邮件功能已禁用")
                else:
                    try:
                        # 使用Outlook发送邮件通知
                        from email_sender import send_email
                        
                        subject = f"ITC报表处理完成通知 - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                        
                        html_content = f"""
                        <h2>ITC报表处理完成</h2>
                        <p><strong>处理时间：</strong>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                        <p><strong>报表文件：</strong>{report_info.get('full_name', 'N/A')}</p>
                        <p><strong>文件大小：</strong>{report_info.get('size', 'N/A')} bytes</p>
                        <p><strong>下载时间：</strong>{report_info.get('download_time', 'N/A')}</p>
                        <hr>
                        <p>请查看详细的处理报告以了解具体情况。</p>
                        <p><em>此邮件由ITC自动处理系统发送</em></p>
                        """
                        
                        # 发送给您自己
                        to_addrs = ["liang.wq.1@pg.com"]
                        
                        send_email(subject, html_content, to_addrs, auto_send=auto_send)
                        
                        if auto_send:
                            log_message("✅ 邮件已直接发送")
                        else:
                            log_message("✅ 邮件预览窗口已打开，请确认后发送")
                        
                    except Exception as email_error:
                        log_message(f"⚠️ 邮件发送失败: {str(email_error)}")
                        log_message("💡 请确保Outlook客户端已登录并可用")
        else:
            log_message("❌ 报表下载失败，不执行数据处理")
    
    except Exception as e:
        log_message(f"❌ 程序运行异常: {str(e)}")
        sys.exit(1)
    
    # 程序结束
    end_time = datetime.now()
    total_duration = (end_time - start_time).total_seconds()
    log_message("\n" + "="*60)
    log_message(f"✅ 所有操作完成！")
    log_message(f"📅 结束时间: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    log_message(f"⏱️ 总耗时: {total_duration:.2f}秒（约{total_duration//60:.0f}分钟{total_duration%60:.0f}秒）")
    if 'DEBUG_PORT' in globals():
        log_message(f"🔧 本次使用的Chrome调试端口: {DEBUG_PORT}")
        log_message(f"📁 本次使用的Chrome用户数据目录: {CHROME_USER_DATA_DIR}")
    log_message("="*60)
    time.sleep(3)