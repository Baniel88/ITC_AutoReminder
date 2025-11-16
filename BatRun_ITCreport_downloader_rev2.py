# -*- coding: utf-8 -*-
"""
ITC报表自动下载器 (优化版)
新增:
1. 可配置下载日期范围/强制下载/复用窗口(小时)参数放置前面
2. 若在复用窗口(< REUSE_WINDOW_HOURS)内已有最近成功CSV则直接复用, 跳过浏览器与再次下载
3. 增强调用 pending_review_report.py 的日志归类 (OUT/ERR 保留, 增加开始/结束标记与汇总)
4. 保持原整体结构和风格, 仅新增最少逻辑
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
import json
import sys
import re
from datetime import datetime, timedelta

# -------------------------- 可配置参数 --------------------------
# 下载报表日期范围(天)
TIME_RANGE_DAYS = 365
# 复用窗口(小时): 最近一次成功的CSV文件若在该时间内生成则直接复用
REUSE_WINDOW_HOURS = 8
# 是否强制忽略复用直接重新下载
FORCE_NEW_DOWNLOAD = False
# 允许的最小CSV文件大小(KB) (防止复用损坏文件)
MIN_CSV_SIZE_KB = 50
# 是否在成功复用时仍调用浏览器验证登录(通常False即可加快速度)
VERIFY_LOGIN_WHEN_REUSE = False

# -------------------------- 固定配置 --------------------------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_PROCESSOR_SCRIPT_NAME = "pending_review_report.py"
REPORT_PROCESSOR_SCRIPT = os.path.join(SCRIPT_DIR, REPORT_PROCESSOR_SCRIPT_NAME)

BASE_REPORT_URL = "https://itc-tool.pg.com/RequestReport/GetRequestExportReport"
ITC_LOGIN_URL = "https://itc-tool.pg.com/"
end_date = datetime.now()
start_date = end_date - timedelta(days=TIME_RANGE_DAYS)
REPORT_PARAMS = {
    "siteId": "193",
    "areaId": "-1",
    "systemId": "-1",
    "categoryId": "-1",
    "requestStatus": "8",          # 8=Pending Review + etc(按ITC接口定义)
    "requestedForId": "",
    "dateRange": f"{start_date.strftime('%m/%d/%Y')} - {end_date.strftime('%m/%d/%Y')}",
    "accessTypeId": "-1"
}
REPORT_URL = f"{BASE_REPORT_URL}?{urllib.parse.urlencode(REPORT_PARAMS)}"

ITC_REPORT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, "ITC report"))
RAW_DATA_DIR = os.path.join(ITC_REPORT_DIR, "RawData")
LOG_DIR = os.path.join(ITC_REPORT_DIR, "Log")

LOGIN_TIMEOUT = 300
LOGIN_CHECK_INTERVAL = 5
DOWNLOAD_TIMEOUT = 600
POST_DOWNLOAD_WAIT = 8
SCRIPT_CALL_TIMEOUT = 900

REUSE_EXISTING_CHROME = True
CLOSE_CHROME_ON_EXIT = False

LOGGED_IN_ELEMENTS = [
    By.CSS_SELECTOR, "#frmRequestAccess",
    By.LINK_TEXT, "退出登录"
]
ITC_DOMAIN = "itc-tool.pg.com"

DRIVER_EXECUTABLE = "chromedriver.exe" if os.name == 'nt' else "chromedriver"
ALLOW_INSECURE_SSL = True

# -------------------------- 日志 --------------------------
def ensure_directory_exists(d):
    if not os.path.exists(d):
        os.makedirs(d, exist_ok=True)
    return d

def log_message(msg):
    ensure_directory_exists(LOG_DIR)
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_file = os.path.join(LOG_DIR, f"itc_downloader_{datetime.now().strftime('%Y%m%d')}.log")
    line = f"[{ts}] {msg}"
    print(line)
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(line + "\n")

def pre_check_report_script():
    log_message("预检查 pending_review_report.py")
    if os.path.exists(REPORT_PROCESSOR_SCRIPT) and os.access(REPORT_PROCESSOR_SCRIPT, os.R_OK):
        log_message("报表处理脚本存在且可读")
        return True
    log_message("报表处理脚本缺失或不可读")
    return False

# -------------------------- ChromeDriver 管理 --------------------------
try:
    from Chrome_Driver_mgr import ChromeDriverManager
    CHROME_DRIVER_MGR_AVAILABLE = True
    chrome_driver_manager = ChromeDriverManager(
        script_dir=SCRIPT_DIR,
        allow_insecure_ssl=ALLOW_INSECURE_SSL,
        log_callback=log_message
    )
    log_message("已加载 Chrome_Driver_mgr")
except ImportError:
    CHROME_DRIVER_MGR_AVAILABLE = False
    chrome_driver_manager = None
    log_message("未找到 Chrome_Driver_mgr，使用简化模式")

def get_chromedriver_path():
    if CHROME_DRIVER_MGR_AVAILABLE and chrome_driver_manager:
        return chrome_driver_manager.get_chromedriver_path()
    p = os.path.join(SCRIPT_DIR, DRIVER_EXECUTABLE)
    if os.path.exists(p):
        log_message(f"找到 ChromeDriver: {p}")
        return p
    log_message(f"未找到 ChromeDriver: {p}")
    return None

def get_chrome_path():
    if CHROME_DRIVER_MGR_AVAILABLE and chrome_driver_manager:
        return chrome_driver_manager.get_chrome_path()
    system = platform.system()
    candidates = []
    if system == "Windows":
        candidates = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ]
    elif system == "Darwin":
        candidates = ["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"]
    else:
        candidates = ["/usr/bin/google-chrome", "/usr/bin/google-chrome-stable"]
    for c in candidates:
        if os.path.exists(c):
            return c
    return candidates[0]

CHROME_PATH = get_chrome_path()

# -------------------------- 调试端口管理 --------------------------
try:
    from chrome_port_config import ITCChromeConfig
    PORT_START = ITCChromeConfig.PORT_START
    PORT_END = ITCChromeConfig.PORT_END
    USER_DATA_PREFIX = ITCChromeConfig.USER_DATA_DIR_PREFIX
    PROJECT_NAME = ITCChromeConfig.PROJECT_NAME
    log_message(f"使用端口范围: {PORT_START}-{PORT_END}")
except ImportError:
    PORT_START = 9233
    PORT_END = 9242
    USER_DATA_PREFIX = "chrome_debug_profile_itc"
    PROJECT_NAME = "ITC_Scorecard"
    log_message("使用备用端口配置: 9233-9242")

DEBUG_PORT = None
CHROME_USER_DATA_DIR = None

def is_port_available(port):
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(1)
            s.bind(('127.0.0.1', port))
            return True
    except Exception:
        return False

def check_existing_chrome_debug():
    found = []
    for p in range(PORT_START, PORT_END + 1):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(0.5)
                if s.connect_ex(('127.0.0.1', p)) == 0:
                    try:
                        r = requests.get(f"http://127.0.0.1:{p}/json", timeout=2)
                        if r.status_code == 200:
                            found.append(p)
                    except:
                        pass
        except:
            pass
    return found[0] if found else None

def allocate_port():
    for p in range(PORT_START, PORT_END + 1):
        if is_port_available(p):
            return p
    return None

def get_user_data_dir(port):
    d = os.path.join(SCRIPT_DIR, f"{USER_DATA_PREFIX}_{port}")
    ensure_directory_exists(d)
    return d

def start_chrome_debug_session():
    global DEBUG_PORT, CHROME_USER_DATA_DIR
    if not os.path.exists(CHROME_PATH):
        log_message(f"Chrome 不存在: {CHROME_PATH}")
        return False

    if REUSE_EXISTING_CHROME:
        existing = check_existing_chrome_debug()
        if existing:
            DEBUG_PORT = existing
            CHROME_USER_DATA_DIR = get_user_data_dir(existing)
            log_message(f"重用 Chrome 调试端口: {existing}")
            return True

    p = allocate_port()
    if not p:
        log_message("没有可用端口")
        return False
    DEBUG_PORT = p
    CHROME_USER_DATA_DIR = get_user_data_dir(p)
    args = [
        CHROME_PATH,
        f"--remote-debugging-port={p}",
        f"--user-data-dir={CHROME_USER_DATA_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        ITC_LOGIN_URL
    ]
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    log_message(f"启动 Chrome 调试会话 端口={p} 目录={CHROME_USER_DATA_DIR}")
    time.sleep(6)
    return True

# -------------------------- 登录检测 --------------------------
def is_itc_logged_in(driver):
    try:
        current_url = driver.current_url
        if ITC_DOMAIN in current_url:
            return True
        for i in range(0, len(LOGGED_IN_ELEMENTS), 2):
            by_t = LOGGED_IN_ELEMENTS[i]
            sel = LOGGED_IN_ELEMENTS[i+1]
            try:
                el = driver.find_element(by_t, sel)
                if el.is_displayed():
                    return True
            except:
                continue
        return False
    except Exception:
        return False

def wait_for_itc_login():
    chromedriver_path = get_chromedriver_path()
    if not chromedriver_path:
        return None
    driver = None
    try:
        opt = Options()
        opt.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
        prefs = {
            "download.default_directory": RAW_DATA_DIR,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": False
        }
        opt.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(service=Service(chromedriver_path), options=opt)
        start = time.time()
        while time.time() - start < LOGIN_TIMEOUT:
            if is_itc_logged_in(driver):
                log_message("登录检测通过")
                return driver
            waited = int(time.time() - start)
            if waited % 15 == 0:
                log_message(f"等待登录中 {waited}s / {LOGIN_TIMEOUT}s")
            time.sleep(LOGIN_CHECK_INTERVAL)
        log_message("登录超时")
        driver.quit()
        return None
    except Exception as e:
        log_message(f"连接调试 Chrome 失败: {e}")
        if driver:
            driver.quit()
        return None

# -------------------------- 报表下载 --------------------------
def download_report(driver):
    log_message("开始下载报表")
    log_message(f"访问 URL: {REPORT_URL}")
    ensure_directory_exists(RAW_DATA_DIR)
    watch_dirs = [RAW_DATA_DIR]
    user_dl = os.path.join(os.environ.get("USERPROFILE", ""), "Downloads")
    if os.path.exists(user_dl):
        watch_dirs.append(user_dl)
    initial = {}
    for d in watch_dirs:
        initial[d] = {f: os.path.getmtime(os.path.join(d, f))
                      for f in os.listdir(d)
                      if os.path.isfile(os.path.join(d, f)) and not f.endswith((".crdownload", ".tmp"))}

    try:
        t0 = time.time()
        driver.get(REPORT_URL)
        log_message("下载请求已发出，轮询文件系统...")
        found = None
        while time.time() - t0 < DOWNLOAD_TIMEOUT:
            for d in watch_dirs:
                for f in os.listdir(d):
                    fp = os.path.join(d, f)
                    if (os.path.isfile(fp) and
                        not f.endswith((".crdownload", ".tmp")) and
                        os.path.getsize(fp) > 1024):
                        is_new = (f not in initial.get(d, {})) or (os.path.getmtime(fp) > t0)
                        if is_new:
                            found = fp
                            break
                if found:
                    break
            if found:
                break
            time.sleep(3)

        if not found:
            log_message("下载超时, 未发现新文件")
            return False, None

        if not found.startswith(RAW_DATA_DIR):
            target = os.path.join(RAW_DATA_DIR, os.path.basename(found))
            if os.path.exists(target):
                stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                name, ext = os.path.splitext(os.path.basename(found))
                target = os.path.join(RAW_DATA_DIR, f"{name}_{stamp}{ext}")
            shutil.move(found, target)
            found = target
        log_message(f"下载完成: {found} 大小 {os.path.getsize(found)/1024/1024:.2f} MB")
        return True, found
    except Exception as e:
        log_message(f"下载异常: {e}")
        return False, None

# -------------------------- 复用逻辑 --------------------------
def find_recent_csv(reuse_hours, min_size_kb):
    if not os.path.isdir(RAW_DATA_DIR):
        return None
    now = time.time()
    cutoff = now - reuse_hours * 3600
    candidates = []
    for f in os.listdir(RAW_DATA_DIR):
        if not f.lower().endswith(".csv"):
            continue
        full = os.path.join(RAW_DATA_DIR, f)
        try:
            stat = os.path.getmtime(full)
            size_kb = os.path.getsize(full) / 1024.0
            if stat >= cutoff and size_kb >= min_size_kb:
                candidates.append((stat, full, size_kb))
        except:
            continue
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]

# -------------------------- 调用处理脚本 --------------------------
def call_report_processor(csv_path):
    log_message("调用处理脚本开始")
    if not os.path.exists(REPORT_PROCESSOR_SCRIPT):
        log_message("处理脚本不存在")
        return False
    if not os.path.exists(csv_path):
        log_message("CSV 不存在")
        return False
    cmd = [sys.executable, REPORT_PROCESSOR_SCRIPT, "--csv-path", csv_path]
    log_message("执行命令: " + " ".join(cmd))
    try:
        r = subprocess.run(
            cmd,
            cwd=SCRIPT_DIR,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=SCRIPT_CALL_TIMEOUT
        )
        log_message("处理脚本输出开始 >>>")
        if r.stdout:
            for ln in r.stdout.splitlines():
                log_message(f"[OUT] {ln}")
        if r.stderr:
            for ln in r.stderr.splitlines():
                log_message(f"[ERR] {ln}")
        log_message("处理脚本输出结束 <<<")
        log_message(f"处理脚本返回码: {r.returncode}")
        return r.returncode == 0
    except subprocess.TimeoutExpired:
        log_message("处理脚本执行超时")
        return False
    except Exception as e:
        log_message(f"调用处理脚本异常: {e}")
        return False

# -------------------------- 主流程 --------------------------
if __name__ == "__main__":
    start = datetime.now()
    log_message("="*70)
    log_message("ITC 报表自动下载器 (优化版)")
    log_message(f"启动时间: {start.strftime('%Y-%m-%d %H:%M:%S')}")
    log_message(f"配置: TIME_RANGE_DAYS={TIME_RANGE_DAYS} REUSE_WINDOW_HOURS={REUSE_WINDOW_HOURS} FORCE_NEW_DOWNLOAD={FORCE_NEW_DOWNLOAD}")
    log_message("="*70)

    if not pre_check_report_script():
        sys.exit(1)

    ensure_directory_exists(ITC_REPORT_DIR)
    ensure_directory_exists(RAW_DATA_DIR)
    ensure_directory_exists(LOG_DIR)

    # 复用判定
    reused_csv = None
    if not FORCE_NEW_DOWNLOAD:
        reused_csv = find_recent_csv(REUSE_WINDOW_HOURS, MIN_CSV_SIZE_KB)
        if reused_csv:
            log_message(f"发现可复用CSV: {reused_csv} (修改时间<={REUSE_WINDOW_HOURS}h)")
        else:
            log_message(f"复用窗口内未找到合格CSV (窗口={REUSE_WINDOW_HOURS}h, 最小大小={MIN_CSV_SIZE_KB}KB)")

    proc_ok = False
    performed_download = False

    if reused_csv and not FORCE_NEW_DOWNLOAD:
        log_message("选择复用模式 -> 跳过下载阶段")
        # 可选验证登录
        if VERIFY_LOGIN_WHEN_REUSE:
            log_message("复用模式下验证登录已开启，尝试连接 Chrome")
            chromedriver_path = get_chromedriver_path()
            if chromedriver_path and start_chrome_debug_session():
                driver = wait_for_itc_login()
                if driver:
                    try:
                        driver.quit()
                    except:
                        pass
        proc_ok = call_report_processor(reused_csv)
    else:
        chromedriver_path = get_chromedriver_path()
        if not chromedriver_path:
            log_message("无法获取 Chromedriver, 退出")
            sys.exit(1)

        if not os.path.exists(CHROME_PATH):
            log_message(f"Chrome 不存在: {CHROME_PATH}")
            sys.exit(1)

        if not start_chrome_debug_session():
            log_message("启动调试会话失败")
            sys.exit(1)

        driver = wait_for_itc_login()
        if not driver:
            log_message("登录失败或超时")
            sys.exit(1)

        ok, csv_file = download_report(driver)
        performed_download = ok
        try:
            driver.quit()
        except:
            pass

        if ok and csv_file:
            log_message(f"等待 {POST_DOWNLOAD_WAIT}s 确保文件稳定")
            time.sleep(POST_DOWNLOAD_WAIT)
            proc_ok = call_report_processor(csv_file)
        else:
            log_message("下载失败，不调用处理脚本")
            proc_ok = False

    # 可选关闭 Chrome
    if CLOSE_CHROME_ON_EXIT and DEBUG_PORT is not None:
        try:
            if platform.system() == "Windows":
                subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'], timeout=10)
            else:
                subprocess.run(['pkill', '-f', f'remote-debugging-port={DEBUG_PORT}'], timeout=10)
            log_message(f"已尝试关闭 Chrome (端口 {DEBUG_PORT})")
        except Exception as e:
            log_message(f"关闭 Chrome 失败: {e}")
    else:
        if DEBUG_PORT is not None:
            log_message(f"Chrome 调试会话保持运行 (端口 {DEBUG_PORT})")

    end = datetime.now()
    dur = (end - start).total_seconds()
    log_message(f"结束时间: {end.strftime('%Y-%m-%d %H:%M:%S')}")
    log_message(f"总耗时: {dur:.2f} 秒")
    log_message(f"流程摘要: performed_download={performed_download} reused_csv={'YES' if reused_csv else 'NO'} proc_ok={proc_ok}")
    log_message("="*70)
    sys.exit(0 if proc_ok else 2)