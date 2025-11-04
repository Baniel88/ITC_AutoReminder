# -*- coding: utf-8 -*-
"""
Chrome Driver 管理器
用于自动下载和管理ChromeDriver，支持版本匹配和多平台
"""

import os
import platform
import subprocess
import re
import zipfile
import shutil
import requests
import warnings
import uuid
import socket
import tempfile
from datetime import datetime
from bs4 import BeautifulSoup

# 导入项目端口配置
try:
    from chrome_port_config import ITCChromeConfig
    PORT_CONFIG_AVAILABLE = True
except ImportError:
    PORT_CONFIG_AVAILABLE = False
    # 如果配置文件不可用，使用默认配置
    class ITCChromeConfig:
        PORT_START = 9233
        PORT_END = 9242
        DEFAULT_PORT = 9233
        USER_DATA_DIR_PREFIX = "ChromeProfile_ITC"
        PROJECT_NAME = "ITC_Scorecard"


class ChromeDriverManager:
    """Chrome Driver 管理器类"""
    
    def __init__(self, script_dir=None, allow_insecure_ssl=True, log_callback=None, instance_id=None, remote_debug_port=None):
        """
        初始化Chrome Driver管理器
        
        Args:
            script_dir: 脚本目录，如果不指定则使用当前文件所在目录
            allow_insecure_ssl: 是否允许不安全的SSL连接
            log_callback: 日志回调函数，用于输出日志信息
            instance_id: 实例唯一标识，用于多实例隔离
            remote_debug_port: 远程调试端口，如果不指定则自动分配
        """
        self.script_dir = script_dir or os.path.dirname(os.path.abspath(__file__))
        self.allow_insecure_ssl = allow_insecure_ssl
        self.log_callback = log_callback or print
        
        # 实例标识，用于多实例隔离
        self.instance_id = instance_id or str(uuid.uuid4())[:8]
        
        # 项目专用端口分配策略（避免与其他项目冲突）
        # ITC_Scorecard 项目: 9233-9242 (10个端口)
        # EVP_Scorecard 项目: 9222-9232 (已被占用)
        if remote_debug_port:
            self.remote_debug_port = remote_debug_port
        else:
            # ITC项目使用专用端口范围
            self.remote_debug_port = self.find_free_port(
                start_port=ITCChromeConfig.PORT_START, 
                max_attempts=ITCChromeConfig.PORT_END - ITCChromeConfig.PORT_START + 1
            )
        
        # 根据操作系统确定驱动文件名
        self.driver_executable = "chromedriver.exe" if os.name == 'nt' else "chromedriver"
        self.driver_path = os.path.join(self.script_dir, self.driver_executable)
        
        # 为每个实例分配独立的用户数据目录（项目级别隔离）
        self.user_data_dir = os.path.join(
            self.script_dir, 
            f"{ITCChromeConfig.USER_DATA_DIR_PREFIX}_{self.instance_id}"
        )
        
        self.log(f"Chrome Driver 管理器初始化完成 [实例ID: {self.instance_id}]")
        self.log(f"工作目录: {self.script_dir}")
        self.log(f"驱动文件名: {self.driver_executable}")
        self.log(f"远程调试端口: {self.remote_debug_port}")
        self.log(f"用户数据目录: {self.user_data_dir}")
        self.log(f"项目标识: ITC_Scorecard")
    
    def log(self, message):
        """记录日志信息"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        formatted_message = f"[{timestamp}] [ChromeDriverMgr-{self.instance_id}] {message}"
        self.log_callback(formatted_message)

    def find_free_port(self, start_port=9222, max_attempts=100):
        """查找可用的端口号"""
        for port in range(start_port, start_port + max_attempts):
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.bind(('localhost', port))
                    return port
            except OSError:
                continue
        raise RuntimeError(f"无法在端口范围 {start_port}-{start_port + max_attempts} 内找到可用端口")
    
    def get_chrome_options(self, additional_options=None):
        """
        获取Chrome选项配置，支持多实例隔离
        
        Args:
            additional_options: 额外的Chrome选项列表
            
        Returns:
            dict: 包含Chrome选项的字典
        """
        # 为当前实例分配独立的调试端口
        debug_port = self.find_free_port()
        
        # 确保用户数据目录存在
        os.makedirs(self.user_data_dir, exist_ok=True)
        
        # 基础选项配置
        options = [
            f"--remote-debugging-port={debug_port}",
            f"--user-data-dir={self.user_data_dir}",
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-web-security",
            "--disable-features=VizDisplayCompositor",
            "--disable-extensions",
            "--disable-plugins",
            "--disable-images",
            "--disable-javascript",
            "--disable-gpu",
            "--disable-dev-tools",
            f"--crash-dumps-dir={os.path.join(tempfile.gettempdir(), f'chrome_crashes_{self.instance_id}')}",
            f"--log-file={os.path.join(tempfile.gettempdir(), f'chrome_log_{self.instance_id}.log')}"
        ]
        
        # 添加额外选项
        if additional_options:
            options.extend(additional_options)
        
        self.log(f"配置Chrome选项 [调试端口: {debug_port}]")
        self.log(f"用户数据目录: {self.user_data_dir}")
        
        return {
            'options': options,
            'debug_port': debug_port,
            'user_data_dir': self.user_data_dir
        }
    
    def cleanup_instance_data(self):
        """清理当前实例的临时数据"""
        try:
            if os.path.exists(self.user_data_dir):
                shutil.rmtree(self.user_data_dir, ignore_errors=True)
                self.log(f"已清理实例数据目录: {self.user_data_dir}")
            
            # 清理日志文件
            log_file = os.path.join(tempfile.gettempdir(), f'chrome_log_{self.instance_id}.log')
            if os.path.exists(log_file):
                os.remove(log_file)
            
            # 清理崩溃转储目录
            crash_dir = os.path.join(tempfile.gettempdir(), f'chrome_crashes_{self.instance_id}')
            if os.path.exists(crash_dir):
                shutil.rmtree(crash_dir, ignore_errors=True)
                
        except Exception as e:
            self.log(f"清理实例数据时发生错误: {str(e)}")
    
    def create_remote_debug_chrome(self, headless=False, additional_options=None):
        """
        创建远程调试模式的Chrome实例（类似EVP项目的方式）
        
        Args:
            headless: 是否使用无头模式
            additional_options: 额外的Chrome选项
            
        Returns:
            dict: 包含Chrome进程信息和连接方法的字典
        """
        try:
            import subprocess
            import time
            
            # 确保用户数据目录存在
            os.makedirs(self.user_data_dir, exist_ok=True)
            
            # 查找Chrome可执行文件路径
            chrome_path = self.get_chrome_path()
            if not chrome_path:
                self.log("❌ 无法找到Chrome浏览器")
                return None
            
            # 构建Chrome启动命令
            chrome_cmd = [
                chrome_path,
                f"--remote-debugging-port={self.remote_debug_port}",
                f"--user-data-dir={self.user_data_dir}",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-web-security",
                "--disable-features=VizDisplayCompositor",
                "--start-maximized"
            ]
            
            if headless:
                chrome_cmd.append("--headless")
            
            # 添加额外选项
            if additional_options:
                chrome_cmd.extend(additional_options)
            
            self.log(f"🚀 启动Chrome远程调试实例...")
            self.log(f"   端口: {self.remote_debug_port}")
            self.log(f"   用户数据目录: {self.user_data_dir}")
            self.log(f"   命令: {' '.join(chrome_cmd)}")
            
            # 启动Chrome进程
            chrome_process = subprocess.Popen(
                chrome_cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                shell=False
            )
            
            # 等待Chrome启动
            time.sleep(3)
            
            # 检查进程是否正常启动
            if chrome_process.poll() is not None:
                stdout, stderr = chrome_process.communicate()
                self.log(f"❌ Chrome启动失败")
                self.log(f"   stdout: {stdout.decode('utf-8', errors='ignore')}")
                self.log(f"   stderr: {stderr.decode('utf-8', errors='ignore')}")
                return None
            
            self.log(f"✅ Chrome远程调试实例启动成功 [PID: {chrome_process.pid}]")
            
            return {
                'process': chrome_process,
                'debug_port': self.remote_debug_port,
                'user_data_dir': self.user_data_dir,
                'instance_id': self.instance_id,
                'chrome_path': chrome_path,
                'connect_method': self.connect_to_remote_chrome,
                'cleanup_method': lambda: self.cleanup_remote_chrome(chrome_process)
            }
            
        except Exception as e:
            self.log(f"❌ 创建远程调试Chrome失败: {str(e)}")
            return None
    
    def connect_to_remote_chrome(self, additional_options=None):
        """
        连接到远程调试Chrome实例（类似EVP项目的连接方式）
        
        Args:
            additional_options: 额外的Chrome选项
            
        Returns:
            selenium WebDriver 实例
        """
        try:
            from selenium import webdriver
            from selenium.webdriver.chrome.service import Service
            from selenium.webdriver.chrome.options import Options
        except ImportError:
            self.log("❌ 未安装selenium，请安装: pip install selenium")
            return None
        
        try:
            # 确保ChromeDriver可用
            driver_path = self.get_chromedriver_path()
            if not driver_path:
                self.log("❌ 无法获取ChromeDriver路径")
                return None
            
            # 配置Chrome选项以连接到远程调试实例
            options = Options()
            options.add_experimental_option("debuggerAddress", f"127.0.0.1:{self.remote_debug_port}")
            
            # 添加额外选项
            if additional_options:
                for option in additional_options:
                    options.add_argument(option)
            
            # 创建Service对象
            service = Service(driver_path)
            
            self.log(f"🔗 连接到远程调试Chrome实例...")
            self.log(f"   调试地址: 127.0.0.1:{self.remote_debug_port}")
            
            # 创建WebDriver实例
            driver = webdriver.Chrome(service=service, options=options)
            
            self.log(f"✅ 成功连接到远程Chrome实例")
            
            return driver
            
        except Exception as e:
            self.log(f"❌ 连接远程Chrome失败: {str(e)}")
            self.log(f"")
            self.log(f"💡 请确保Chrome远程调试实例已启动：")
            self.log(f"   使用 create_remote_debug_chrome() 方法先启动Chrome")
            self.log(f"   或手动启动命令：")
            chrome_path = self.get_chrome_path()
            if chrome_path:
                self.log(f'   & "{chrome_path}" --remote-debugging-port={self.remote_debug_port} --user-data-dir="{self.user_data_dir}"')
            return None
    
    def cleanup_remote_chrome(self, chrome_process):
        """清理远程调试Chrome进程"""
        try:
            if chrome_process and chrome_process.poll() is None:
                chrome_process.terminate()
                
                # 等待进程结束
                try:
                    chrome_process.wait(timeout=5)
                    self.log("✅ Chrome进程已正常关闭")
                except subprocess.TimeoutExpired:
                    # 强制杀死进程
                    chrome_process.kill()
                    chrome_process.wait()
                    self.log("⚠️ Chrome进程已强制关闭")
            
            # 清理用户数据目录
            if os.path.exists(self.user_data_dir):
                try:
                    import time
                    time.sleep(1)  # 等待文件释放
                    shutil.rmtree(self.user_data_dir, ignore_errors=True)
                    self.log(f"🧹 已清理用户数据目录: {self.user_data_dir}")
                except Exception as e:
                    self.log(f"⚠️ 清理用户数据目录时出错: {str(e)}")
        
        except Exception as e:
            self.log(f"❌ 清理Chrome进程时出错: {str(e)}")
    
    def create_evp_style_chrome_instance(self, headless=False, additional_options=None):
        """
        创建EVP项目风格的Chrome实例（完整的启动+连接流程）
        
        Args:
            headless: 是否使用无头模式
            additional_options: 额外的Chrome选项
            
        Returns:
            dict: 包含driver和管理信息的字典
        """
        # 1. 启动远程调试Chrome
        chrome_info = self.create_remote_debug_chrome(headless, additional_options)
        if not chrome_info:
            return None
        
        # 2. 连接到Chrome实例
        driver = self.connect_to_remote_chrome()
        if not driver:
            # 如果连接失败，清理Chrome进程
            chrome_info['cleanup_method']()
            return None
        
        return {
            'driver': driver,
            'chrome_process': chrome_info['process'],
            'debug_port': chrome_info['debug_port'],
            'user_data_dir': chrome_info['user_data_dir'],
            'instance_id': self.instance_id,
            'cleanup_method': lambda: self.cleanup_evp_style_instance(driver, chrome_info)
        }
    
    def cleanup_evp_style_instance(self, driver, chrome_info):
        """清理EVP风格的Chrome实例"""
        try:
            # 关闭WebDriver
            if driver:
                driver.quit()
                self.log("🧹 WebDriver已关闭")
        except:
            pass
        
        try:
            # 清理Chrome进程
            chrome_info['cleanup_method']()
        except:
            pass
    
    def get_chrome_full_version(self):
        """获取Chrome完整版本号"""
        try:
            if os.name == 'nt':
                import winreg
                # 读取Chrome版本注册表
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
                full_version, _ = winreg.QueryValueEx(key, "version")
                winreg.CloseKey(key)
            else:
                # Linux/Mac获取完整版本
                chrome_cmd = ["google-chrome" if platform.system() != "Darwin" 
                             else "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "--version"]
                result = subprocess.run(chrome_cmd, capture_output=True, text=True)
                version_match = re.search(r"Chrome (\d+\.\d+\.\d+\.\d+)", result.stdout)
                full_version = version_match.group(1) if version_match else None
            
            if full_version and re.match(r"\d+\.\d+\.\d+\.\d+", full_version):
                self.log(f"获取到Chrome完整版本: {full_version}")
                return full_version
            else:
                self.log(f"⚠️ 获取的Chrome版本格式异常: {full_version}")
                return None
        except Exception as e:
            self.log(f"❌ 获取Chrome完整版本失败: {str(e)}")
            return None
    
    def get_stable_chromedriver_version(self):
        """从官方页面获取稳定版ChromeDriver的精确版本号"""
        try:
            url = "https://googlechromelabs.github.io/chrome-for-testing/"
            # 处理SSL证书问题
            with warnings.catch_warnings():
                if self.allow_insecure_ssl:
                    import urllib3
                    warnings.simplefilter("ignore", category=urllib3.exceptions.InsecureRequestWarning)
                    response = requests.get(url, timeout=10, verify=False)  # 禁用SSL验证
                else:
                    response = requests.get(url, timeout=10, verify=True)
            response.raise_for_status()
            
            # 解析页面内容
            soup = BeautifulSoup(response.text, 'html.parser')
            stable_section = soup.find('section', id='stable')
            if not stable_section:
                self.log("❌ 未找到Stable版本信息")
                return None
                
            # 提取版本号
            version_p = stable_section.find('p')
            if not version_p:
                self.log("❌ 未找到版本信息段落")
                return None
            
            version_text = version_p.get_text() if hasattr(version_p, 'get_text') else str(version_p)
            version_match = re.search(r"Version: (\d+\.\d+\.\d+\.\d+)", version_text)
            if version_match:
                stable_version = version_match.group(1)
                self.log(f"从官方页面获取到稳定版ChromeDriver版本: {stable_version}")
                return stable_version
            else:
                self.log("❌ 无法从页面提取版本号")
                return None
        except Exception as e:
            self.log(f"❌ 获取稳定版版本号失败: {str(e)}")
            # 提供备选方案
            chrome_version = self.get_chrome_full_version()
            if chrome_version:
                fallback_version = ".".join(chrome_version.split('.')[:3]) + ".76"
                self.log(f"⚠️ 使用备选版本号: {fallback_version}")
                return fallback_version
            return None
    
    def download_chromedriver(self):
        """增强版ChromeDriver下载，兼容版本号差异"""
        driver_url = None  # 初始化变量，避免未定义错误
        
        chrome_full_version = self.get_chrome_full_version()
        if not chrome_full_version:
            self.log("❌ 无法获取Chrome完整版本，无法下载ChromeDriver")
            return False
        
        # 获取官方稳定版精确版本号
        stable_version = self.get_stable_chromedriver_version()
        if not stable_version:
            self.log("❌ 无法获取官方稳定版ChromeDriver版本")
            return False
        
        # 宽松版本匹配：只要求前3段一致
        chrome_main_version = ".".join(chrome_full_version.split('.')[:3])
        driver_main_version = ".".join(stable_version.split('.')[:3])
        
        if chrome_main_version != driver_main_version:
            self.log(f"⚠️ Chrome主版本({chrome_main_version})与Driver主版本({driver_main_version})不兼容")
            return False
        else:
            self.log(f"✅ 版本兼容性检查通过 - Chrome: {chrome_full_version}, Driver: {stable_version}")
        
        try:
            # 确定系统对应的平台参数
            system = platform.system()
            if system == "Windows":
                platform_name = "win64" if platform.machine().endswith('64') else "win32"
                driver_zip_name = f"chromedriver-{platform_name}.zip"
            elif system == "Darwin":
                platform_name = "mac-arm64" if platform.machine() == "arm64" else "mac-x64"
                driver_zip_name = f"chromedriver-{platform_name}.zip"
            else:  # Linux
                platform_name = "linux64"
                driver_zip_name = f"chromedriver-{platform_name}.zip"
            
            # 构造下载URL
            driver_url = f"https://storage.googleapis.com/chrome-for-testing-public/{stable_version}/{platform_name}/{driver_zip_name}"
            self.log(f"✅ 构造的ChromeDriver下载URL: {driver_url}")
            
            # 移除旧驱动
            if os.path.exists(self.driver_path):
                os.remove(self.driver_path)
                self.log(f"已移除旧版本ChromeDriver: {self.driver_path}")
            
            # 下载驱动压缩包
            driver_zip_path = os.path.join(self.script_dir, "chromedriver.zip")
            self.log(f"开始下载ChromeDriver（保存路径: {driver_zip_path}）")
            
            with warnings.catch_warnings():
                if self.allow_insecure_ssl:
                    import urllib3
                    warnings.simplefilter("ignore", category=urllib3.exceptions.InsecureRequestWarning)
                
                with requests.get(driver_url, stream=True, timeout=30, verify=not self.allow_insecure_ssl) as r:
                    r.raise_for_status()
                    total_size = int(r.headers.get('content-length', 0))
                    downloaded_size = 0
                    
                    with open(driver_zip_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=8192):
                            if chunk:
                                downloaded_size += len(chunk)
                                f.write(chunk)
                                if total_size > 0:
                                    progress = (downloaded_size / total_size) * 100
                                    if int(progress) % 20 == 0 and progress > 0:
                                        self.log(f"ChromeDriver下载进度: {progress:.1f}%")
            
            # 解压处理
            with zipfile.ZipFile(driver_zip_path, 'r') as zip_ref:
                zip_ref.extractall(self.script_dir)
                extracted_dir = os.path.join(self.script_dir, f"chromedriver-{platform_name}")
                extracted_driver_path = os.path.join(extracted_dir, self.driver_executable)
                
                if os.path.exists(extracted_driver_path):
                    shutil.move(extracted_driver_path, self.driver_path)
                    self.log(f"已将ChromeDriver移动到: {self.driver_path}")
                else:
                    raise Exception(f"解压后未找到驱动文件: {extracted_driver_path}")
                
                shutil.rmtree(extracted_dir)
            
            # 清理与权限设置
            os.remove(driver_zip_path)
            if system != "Windows":
                os.chmod(self.driver_path, 0o755)
                self.log("已为ChromeDriver设置执行权限")
            
            # 验证驱动版本
            result = subprocess.run([self.driver_path, "--version"], capture_output=True, text=True)
            version_match = re.search(r"ChromeDriver (\d+\.\d+\.\d+\.\d+)", result.stdout)
            if version_match:
                driver_version = version_match.group(1)
                self.log(f"✅ ChromeDriver下载成功！版本: {driver_version}")
            else:
                self.log("✅ ChromeDriver下载成功！")
            return True
            
        except requests.exceptions.HTTPError as e:
            self.log(f"❌ 下载URL错误: {str(e)}")
            if driver_url:
                self.log(f"💡 请手动检查URL是否有效: {driver_url}")
        except Exception as e:
            self.log(f"❌ 下载ChromeDriver失败: {str(e)}")
        finally:
            # 清理临时文件
            temp_zip = os.path.join(self.script_dir, "chromedriver.zip")
            if os.path.exists(temp_zip):
                os.remove(temp_zip)
        
        return False
    
    def is_driver_compatible(self):
        """检查当前驱动是否与Chrome版本兼容"""
        try:
            if not os.path.exists(self.driver_path):
                return False
            
            # 获取驱动版本
            driver_result = subprocess.run([self.driver_path, "--version"], capture_output=True, text=True)
            version_match = re.search(r"ChromeDriver (\d+\.\d+\.\d+\.\d+)", driver_result.stdout)
            if not version_match:
                self.log("❌ 无法获取驱动版本信息")
                return False
            driver_version = version_match.group(1)
            
            # 获取Chrome版本
            chrome_version = self.get_chrome_full_version()
            if not chrome_version:
                return False
            
            # 只检查前3段版本号
            driver_main = ".".join(driver_version.split('.')[:3])
            chrome_main = ".".join(chrome_version.split('.')[:3])
            
            compatible = driver_main == chrome_main
            if compatible:
                self.log(f"✅ 驱动版本兼容 - Chrome: {chrome_version}, Driver: {driver_version}")
            else:
                self.log(f"⚠️ 驱动版本不兼容 - Chrome: {chrome_version}, Driver: {driver_version}")
            
            return compatible
        except Exception as e:
            self.log(f"❌ 检查驱动兼容性时出错: {str(e)}")
            return False
    
    def get_chromedriver_path(self):
        """获取ChromeDriver路径，如果不存在或不兼容则自动下载"""
        # 检查已存在的驱动是否兼容
        if os.path.exists(self.driver_path) and self.is_driver_compatible():
            self.log(f"找到兼容的ChromeDriver: {self.driver_path}")
            return self.driver_path
        
        # 不存在或不兼容，尝试下载
        self.log("⚠️ 未找到兼容的ChromeDriver，开始下载流程")
        if self.download_chromedriver():
            return self.driver_path
        else:
            self.log("❌ ChromeDriver下载失败")
            return None
    
    def get_chrome_path(self):
        """获取Chrome浏览器可执行文件路径"""
        system = platform.system()
        self.log(f"检测到操作系统: {system}")
        
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
                            self.log(f"从注册表找到Chrome路径: {value}")
                            return value
                    except:
                        continue
            except Exception as e:
                self.log(f"从注册表查找Chrome失败: {str(e)}")
            
            common_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            ]
            for path in common_paths:
                if os.path.exists(path):
                    self.log(f"找到Chrome路径: {path}")
                    return path
            return common_paths[0]  # 返回默认路径，即使不存在
        
        elif system == "Darwin":
            chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
            if os.path.exists(chrome_path):
                self.log(f"找到Chrome路径: {chrome_path}")
            return chrome_path
        else:  # Linux
            chrome_path = "/usr/bin/google-chrome"
            if os.path.exists(chrome_path):
                self.log(f"找到Chrome路径: {chrome_path}")
            return chrome_path
    
    def get_driver_version(self):
        """获取当前ChromeDriver版本"""
        try:
            if not os.path.exists(self.driver_path):
                return None
            
            result = subprocess.run([self.driver_path, "--version"], capture_output=True, text=True)
            version_match = re.search(r"ChromeDriver (\d+\.\d+\.\d+\.\d+)", result.stdout)
            return version_match.group(1) if version_match else None
        except Exception as e:
            self.log(f"❌ 获取驱动版本失败: {str(e)}")
            return None
    
    def check_environment(self):
        """检查Chrome和ChromeDriver环境"""
        self.log("开始检查Chrome和ChromeDriver环境...")
        
        # 检查Chrome
        chrome_path = self.get_chrome_path()
        chrome_exists = os.path.exists(chrome_path)
        chrome_version = self.get_chrome_full_version()
        
        # 检查ChromeDriver
        driver_path = self.get_chromedriver_path()
        driver_exists = driver_path and os.path.exists(driver_path)
        driver_version = self.get_driver_version()
        
        # 输出检查结果
        self.log("="*50)
        self.log("环境检查结果:")
        self.log(f"Chrome浏览器: {'✅ 已安装' if chrome_exists else '❌ 未找到'}")
        self.log(f"Chrome路径: {chrome_path}")
        self.log(f"Chrome版本: {chrome_version or '未知'}")
        self.log("-" * 30)
        self.log(f"ChromeDriver: {'✅ 已准备' if driver_exists else '❌ 未找到'}")
        self.log(f"Driver路径: {driver_path or '无'}")
        self.log(f"Driver版本: {driver_version or '未知'}")
        self.log("-" * 30)
        
        if chrome_exists and driver_exists:
            compatible = self.is_driver_compatible()
            self.log(f"版本兼容性: {'✅ 兼容' if compatible else '❌ 不兼容'}")
            if compatible:
                self.log("🎉 环境检查通过，可以正常使用！")
                return True
            else:
                self.log("⚠️ 版本不兼容，建议重新下载ChromeDriver")
        else:
            self.log("❌ 环境不完整，请检查Chrome和ChromeDriver安装")
        
        self.log("="*50)
        return chrome_exists and driver_exists and self.is_driver_compatible()


# 便捷函数，用于向后兼容
def get_chromedriver_path(script_dir=None, log_callback=None):
    """
    获取ChromeDriver路径的便捷函数
    
    Args:
        script_dir: 脚本目录
        log_callback: 日志回调函数
    
    Returns:
        str: ChromeDriver路径，如果失败返回None
    """
    manager = ChromeDriverManager(script_dir, log_callback=log_callback)
    return manager.get_chromedriver_path()


def get_chrome_path(log_callback=None):
    """
    获取Chrome路径的便捷函数
    
    Args:
        log_callback: 日志回调函数
    
    Returns:
        str: Chrome路径
    """
    manager = ChromeDriverManager(log_callback=log_callback)
    return manager.get_chrome_path()


def check_chrome_environment(script_dir=None, log_callback=None):
    """
    检查Chrome环境的便捷函数
    
    Args:
        script_dir: 脚本目录
        log_callback: 日志回调函数
    
    Returns:
        bool: 环境是否正常
    """
    manager = ChromeDriverManager(script_dir, log_callback=log_callback)
    return manager.check_environment()


# 主程序入口（用于独立测试）
if __name__ == "__main__":
    # 创建管理器实例
    manager = ChromeDriverManager()
    
    # 检查环境
    print("Chrome Driver 管理器 - 环境检查")
    print("="*50)
    
    if manager.check_environment():
        print("\n✅ 所有检查通过，Chrome环境正常！")
    else:
        print("\n❌ 环境检查未通过，请检查问题并重试")
        
        # 提供手动解决方案
        print("\n💡 手动解决方案:")
        print("1. 确保已安装Google Chrome浏览器")
        print("2. 访问: https://googlechromelabs.github.io/chrome-for-testing/")
        print("3. 下载与Chrome版本匹配的ChromeDriver")
        print(f"4. 将驱动文件放在: {manager.script_dir}")
    
    print("\n测试完成！")


def create_isolated_chrome_driver(instance_id=None, script_dir=None, additional_options=None, log_callback=None, use_evp_style=True):
    """
    创建隔离的Chrome WebDriver实例的便捷函数
    
    Args:
        instance_id: 实例唯一标识
        script_dir: 脚本目录
        additional_options: 额外的Chrome选项
        log_callback: 日志回调函数
        use_evp_style: 是否使用EVP项目风格（远程调试模式）
        
    Returns:
        dict: 包含driver和配置信息的字典
    """
    manager = ChromeDriverManager(
        script_dir=script_dir, 
        instance_id=instance_id,
        log_callback=log_callback
    )
    
    if use_evp_style:
        return manager.create_evp_style_chrome_instance(additional_options=additional_options)
    else:
        # 保留原有的直接创建方式
        return manager.get_chrome_options(additional_options)


def cleanup_all_chrome_instances():
    """清理所有Chrome实例的临时数据"""
    import glob
    
    temp_dir = tempfile.gettempdir()
    
    # 清理用户数据目录
    profile_dirs = glob.glob(os.path.join(temp_dir, "chrome_profile_*"))
    for profile_dir in profile_dirs:
        try:
            shutil.rmtree(profile_dir, ignore_errors=True)
            print(f"已清理: {profile_dir}")
        except:
            pass
    
    # 清理日志文件
    log_files = glob.glob(os.path.join(temp_dir, "chrome_log_*.log"))
    for log_file in log_files:
        try:
            os.remove(log_file)
            print(f"已清理日志: {log_file}")
        except:
            pass
    
    # 清理崩溃转储目录
    crash_dirs = glob.glob(os.path.join(temp_dir, "chrome_crashes_*"))
    for crash_dir in crash_dirs:
        try:
            shutil.rmtree(crash_dir, ignore_errors=True)
            print(f"已清理崩溃转储: {crash_dir}")
        except:
            pass
    
    print("Chrome实例清理完成！")


class ChromeDriverPool:
    """Chrome Driver 连接池，管理多个隔离的Chrome实例"""
    
    def __init__(self, max_instances=3, script_dir=None, log_callback=None):
        """
        初始化Chrome Driver连接池
        
        Args:
            max_instances: 最大实例数量
            script_dir: 脚本目录
            log_callback: 日志回调函数
        """
        self.max_instances = max_instances
        self.script_dir = script_dir
        self.log_callback = log_callback or print
        self.instances = []
        self.available_instances = []
        
    def get_driver(self, additional_options=None):
        """获取一个可用的Chrome driver实例"""
        # 如果有可用实例，直接返回
        if self.available_instances:
            return self.available_instances.pop()
        
        # 如果未达到最大实例数，创建新实例
        if len(self.instances) < self.max_instances:
            driver_info = create_isolated_chrome_driver(
                script_dir=self.script_dir,
                additional_options=additional_options,
                log_callback=self.log_callback
            )
            
            if driver_info:
                self.instances.append(driver_info)
                return driver_info
        
        # 无可用实例且已达最大数量
        raise RuntimeError(f"已达到最大实例数量 ({self.max_instances})，请释放一些实例")
    
    def release_driver(self, driver_info):
        """释放一个driver实例回到池中"""
        if driver_info in self.instances:
            self.available_instances.append(driver_info)
    
    def close_all(self):
        """关闭所有实例"""
        for driver_info in self.instances:
            try:
                driver_info['driver'].quit()
                driver_info['cleanup_method']()
            except:
                pass
        
        self.instances.clear()
        self.available_instances.clear()
        self.log_callback("所有Chrome实例已关闭")


# 使用示例和测试代码
if __name__ == "__main__":
    print("Chrome Driver 管理器 - 多实例支持测试")
    print("=" * 50)
    
    # 创建两个隔离的实例
    print("\n1. 创建隔离实例测试...")
    
    instance1 = create_isolated_chrome_driver(instance_id="test1")
    instance2 = create_isolated_chrome_driver(instance_id="test2")
    
    if instance1 and instance2:
        print(f"✅ 实例1创建成功 [端口: {instance1['debug_port']}]")
        print(f"✅ 实例2创建成功 [端口: {instance2['debug_port']}]")
        
        # 测试同时访问不同网站
        try:
            instance1['driver'].get("https://www.google.com")
            instance2['driver'].get("https://www.baidu.com")
            
            print(f"实例1标题: {instance1['driver'].title}")
            print(f"实例2标题: {instance2['driver'].title}")
            
        finally:
            # 清理实例
            instance1['driver'].quit()
            instance2['driver'].quit()
            instance1['cleanup_method']()
            instance2['cleanup_method']()
    
    print("\n2. 连接池测试...")
    pool = ChromeDriverPool(max_instances=2)
    
    try:
        # 获取两个实例
        driver1 = pool.get_driver()
        driver2 = pool.get_driver()
        
        if driver1 and driver2:
            print("✅ 连接池测试成功")
        
    finally:
        pool.close_all()
    
    print("\n测试完成！")