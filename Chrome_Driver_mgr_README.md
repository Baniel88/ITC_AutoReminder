# Chrome Driver 管理器使用说明

## 概述

`Chrome_Driver_mgr.py` 是一个独立的 Chrome Driver 管理模块，从原始的 ITC 报表下载器中提取出来，可以被其他项目复用。

## 主要功能

- 🔍 **自动检测Chrome版本**：获取本机安装的Chrome浏览器版本
- 📥 **自动下载ChromeDriver**：从官方源下载匹配的ChromeDriver版本
- ✅ **版本兼容性检查**：确保Chrome和ChromeDriver版本兼容
- 🎯 **多平台支持**：支持Windows、macOS和Linux
- 📋 **环境检查**：一键检查Chrome和ChromeDriver环境状态

## 快速开始

### 1. 基本使用（类方式）

```python
from Chrome_Driver_mgr import ChromeDriverManager

# 创建管理器实例
manager = ChromeDriverManager()

# 检查环境
if manager.check_environment():
    print("Chrome环境正常！")
    
    # 获取ChromeDriver路径（自动下载如果需要）
    driver_path = manager.get_chromedriver_path()
    
    # 获取Chrome路径
    chrome_path = manager.get_chrome_path()
    
    print(f"ChromeDriver: {driver_path}")
    print(f"Chrome: {chrome_path}")
```

### 2. 便捷函数使用

```python
from Chrome_Driver_mgr import get_chromedriver_path, get_chrome_path, check_chrome_environment

# 快速环境检查
if check_chrome_environment():
    # 获取路径
    driver_path = get_chromedriver_path()
    chrome_path = get_chrome_path()
```

### 3. 在原始项目中集成

```python
# 在你的原始脚本中
try:
    from Chrome_Driver_mgr import ChromeDriverManager
    
    # 创建管理器（传入日志回调函数）
    driver_manager = ChromeDriverManager(
        script_dir=os.path.dirname(__file__),
        log_callback=your_log_function
    )
    
    # 获取ChromeDriver路径
    chromedriver_path = driver_manager.get_chromedriver_path()
    
except ImportError:
    # 如果模块不可用，使用备用方案
    chromedriver_path = "fallback_path"
```

## 高级配置

### 自定义初始化参数

```python
manager = ChromeDriverManager(
    script_dir="/custom/path",           # 自定义脚本目录
    allow_insecure_ssl=True,             # 允许不安全的SSL连接
    log_callback=your_log_function       # 自定义日志函数
)
```

### 自定义日志函数

```python
def my_logger(message):
    with open("chrome_driver.log", "a") as f:
        f.write(f"{datetime.now()}: {message}\n")

manager = ChromeDriverManager(log_callback=my_logger)
```

## API 参考

### ChromeDriverManager 类

#### 初始化参数

- `script_dir` (str, 可选): 脚本目录，默认为当前文件目录
- `allow_insecure_ssl` (bool, 可选): 是否允许不安全的SSL，默认True
- `log_callback` (function, 可选): 日志回调函数，默认print

#### 主要方法

- `get_chromedriver_path()`: 获取ChromeDriver路径，不存在时自动下载
- `get_chrome_path()`: 获取Chrome浏览器路径
- `check_environment()`: 检查Chrome环境完整性
- `get_chrome_full_version()`: 获取Chrome完整版本号
- `get_driver_version()`: 获取ChromeDriver版本
- `is_driver_compatible()`: 检查版本兼容性
- `download_chromedriver()`: 手动下载ChromeDriver

### 便捷函数

- `get_chromedriver_path(script_dir, log_callback)`: 快速获取ChromeDriver路径
- `get_chrome_path(log_callback)`: 快速获取Chrome路径
- `check_chrome_environment(script_dir, log_callback)`: 快速环境检查

## 文件结构

```txt
项目目录/
├── Chrome_Driver_mgr.py          # 主模块文件
├── test_chrome_driver_mgr.py     # 测试脚本
├── chromedriver.exe              # 下载的ChromeDriver（自动生成）
└── 你的项目文件.py
```

## 故障排除

### 1. 导入失败

确保 `Chrome_Driver_mgr.py` 在Python路径中或与你的脚本在同一目录。

### 2. Chrome未找到

- Windows: 确保Chrome安装在标准路径
- macOS: 确保Chrome在Applications文件夹
- Linux: 确保google-chrome在PATH中

### 3. 下载失败

- 检查网络连接
- 如果有防火墙，确保允许访问 googleapis.com
- 可以设置 `allow_insecure_ssl=True` 解决SSL问题

### 4. 版本不兼容

- 更新Chrome到最新版本
- 删除旧的chromedriver文件让程序重新下载

## 在其他项目中使用

### 1. 复制文件

将 `Chrome_Driver_mgr.py` 复制到你的新项目目录中。

### 2. 简单集成示例

```python
# my_selenium_project.py
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from Chrome_Driver_mgr import ChromeDriverManager

def setup_chrome_driver():
    """设置Chrome驱动"""
    manager = ChromeDriverManager()
    
    if not manager.check_environment():
        print("Chrome环境不完整，正在尝试修复...")
    
    driver_path = manager.get_chromedriver_path()
    if not driver_path:
        raise Exception("无法获取ChromeDriver")
    
    return driver_path

def main():
    # 获取ChromeDriver路径
    driver_path = setup_chrome_driver()
    
    # 创建Chrome服务
    service = Service(driver_path)
    
    # 启动Chrome
    driver = webdriver.Chrome(service=service)
    
    try:
        # 你的自动化代码
        driver.get("https://www.google.com")
        print(f"页面标题: {driver.title}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
```

## 更新日志

### v1.0.0

- 从ITC报表下载器中提取Chrome Driver管理功能
- 支持自动版本检测和下载
- 提供类和函数两种使用方式
- 添加完整的错误处理和日志功能

## 许可证

此模块作为开源工具提供，可在你的项目中自由使用和修改。

## 支持

如果遇到问题，请检查：

1. Chrome浏览器是否正确安装
2. 网络连接是否正常
3. 防火墙设置是否阻止下载
4. Python环境是否包含所需依赖（requests, BeautifulSoup4等）

---

**提示**: 运行 `test_chrome_driver_mgr.py` 可以快速验证模块是否正常工作。