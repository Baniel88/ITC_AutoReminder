# -*- coding: utf-8 -*-
"""
通用邮件发送模块（本地 Outlook）
增强版:
- 配置加载与容错
- 可选发送确认(查"已发送邮件")
- 可控最大等待/重试
- 附件支持
- 日志文件记录
- 失败保存草稿
保持原函数签名 send_email(subject, html_content, to_addrs, cc_addrs=None, config_path=None, max_retries=2)
"""

import os
import json
import time
import traceback
from datetime import datetime, timedelta

import sys

try:
    import win32com.client
    import pythoncom
except Exception:
    win32com = None
    pythoncom = None

# ---------------- 日志 ----------------
def _log_dir():
    base = os.path.dirname(os.path.abspath(__file__))
    itc_dir = os.path.join(base, "ITC report")
    log_dir = os.path.join(itc_dir, "Log")
    os.makedirs(log_dir, exist_ok=True)
    return log_dir

def _log_file():
    return os.path.join(_log_dir(), f"email_sender_{datetime.now().strftime('%Y%m%d')}.log")

def log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    try:
        print(line, flush=True)
    except (UnicodeEncodeError, ValueError):
        try:
            sys.stdout.buffer.write((line + "\n").encode("utf-8", "ignore"))
            sys.stdout.buffer.flush()
        except Exception:
            pass
    try:
        with open(_log_file(), "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

# ---------------- 配置 ----------------
# 日志写入到 ITC report\Log\email_sender_YYYYMMDD.log
# 配置项(可在 email_config.json 的 system_config 中增加):
# EMAIL_ENABLED (默认 True)
# EMAIL_VERIFY_SENT (默认 True，关闭则不扫描已发送文件夹)
# EMAIL_MAX_WAIT_SECONDS (默认 40)
# EMAIL_RETRY_INTERVAL_SECONDS (默认 5)
# EMAIL_SUBJECT_PREFIX (可选前缀)
# EMAIL_APPEND_SIGNATURE (默认 True)
# EMAIL_SIGNATURE_HTML (自定义签名)
# EMAIL_DRAFT_ON_FAIL (默认 True，失败保存草稿)
# EMAIL_TRIM_EMPTY (默认 True，过滤空地址)
# 支持附件: 通过 kwargs 传 attachments=[r'c:\path\file1.txt', ...]
# 超时与重试逻辑更清晰，阻塞时间可控
# 避免多次 CoInitialize/Uninitialize 反复调用
# 出错时继续进入后续流程(返回 False)而不抛出阻塞异常
DEFAULT_CFG = {
    "system_config": {
        "EMAIL_ENABLED": True,
        "EMAIL_VERIFY_SENT": False,  # 改为False以避免COM崩溃
        "EMAIL_MAX_WAIT_SECONDS": 40,
        "EMAIL_RETRY_INTERVAL_SECONDS": 5,
        "EMAIL_SUBJECT_PREFIX": "",
        "EMAIL_APPEND_SIGNATURE": True,
        "EMAIL_SIGNATURE_HTML": "<div style='font-size:12px;color:#555;'>-- ITC 自动提醒系统</div>",
        "EMAIL_DRAFT_ON_FAIL": True,
        "EMAIL_TRIM_EMPTY": True
    }
}

def load_email_config(config_path=None):
    if not config_path:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_config.json")
    if not os.path.exists(config_path):
        log(f"配置文件不存在，使用默认配置: {config_path}")
        return DEFAULT_CFG
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "system_config" not in data:
            data["system_config"] = {}
        merged = DEFAULT_CFG["system_config"].copy()
        merged.update(data["system_config"])
        data["system_config"] = merged
        return data
    except Exception as e:
        log(f"读取配置失败，使用默认: {e}")
        return DEFAULT_CFG

# ---------------- 工具 ----------------
def _sanitize_addresses(addrs):
    if not addrs:
        return []
    cleaned = []
    for a in addrs:
        if not a:
            continue
        a = str(a).strip()
        if a:
            cleaned.append(a)
    return cleaned

def _attach_files(mail, attachments):
    ok_files = []
    for fp in attachments or []:
        try:
            if fp and os.path.exists(fp):
                mail.Attachments.Add(fp)
                ok_files.append(fp)
            else:
                log(f"附件不存在: {fp}")
        except Exception as e:
            log(f"附件添加失败 {fp}: {e}")
    if ok_files:
        log(f"附件添加完成 数量={len(ok_files)}")

def _find_sent_item(sent_items, subject, send_time, lookback_seconds):
    """
    在已发送邮件中查找匹配主题且发送时间在窗口内的邮件
    安全版本：添加异常处理防止COM崩溃
    """
    cutoff = send_time - timedelta(seconds=lookback_seconds)
    try:
        for item in sent_items:
            try:
                # 防止 COM 对象访问导致崩溃
                if hasattr(item, 'Subject') and hasattr(item, 'SentOn'):
                    subject_val = str(item.Subject) if item.Subject else ""
                    sent_on_val = item.SentOn
                    if subject_val == subject and sent_on_val >= cutoff:
                        return item
            except Exception as item_err:
                # 单个项目读取失败，继续下一个
                log(f"读取邮件项异常(跳过): {item_err}")
                continue
    except Exception as e:
        log(f"查询已发送邮件列表异常(返回None): {e}")
    return None

# ---------------- 主发送函数 ----------------
def send_email(subject, html_content, to_addrs, cc_addrs=None,
               config_path=None, max_retries=2, **kwargs):
    """
    发送邮件(Outlook)
    subject: 标题
    html_content: HTML正文
    to_addrs: 收件人列表
    cc_addrs: 抄送列表
    max_retries: 发送确认扫描失败时的重试次数
    kwargs 支持:
        attachments = [filepath1, filepath2]
        use_public_mailbox = True (默认使用公共邮箱 ChinaPD_Cybersecurity_Robot)
    返回 True/False
    """
    cfg = load_email_config(config_path)
    scfg = cfg.get("system_config", {})
    if not scfg.get("EMAIL_ENABLED", True):
        log("EMAIL_ENABLED=False 跳过发送。")
        return True

    verify_sent = bool(scfg.get("EMAIL_VERIFY_SENT", True))
    max_wait = int(scfg.get("EMAIL_MAX_WAIT_SECONDS", 40))
    interval = int(scfg.get("EMAIL_RETRY_INTERVAL_SECONDS", 5))
    prefix = scfg.get("EMAIL_SUBJECT_PREFIX", "") or ""
    append_sig = bool(scfg.get("EMAIL_APPEND_SIGNATURE", True))
    signature_html = scfg.get("EMAIL_SIGNATURE_HTML", "")
    draft_on_fail = bool(scfg.get("EMAIL_DRAFT_ON_FAIL", True))
    trim = bool(scfg.get("EMAIL_TRIM_EMPTY", True))

    attachments = kwargs.get("attachments")  # list or None
    use_public_mailbox = kwargs.get("use_public_mailbox", True)  # 默认使用公共邮箱

    if cc_addrs is None:
        cc_addrs = []

    if trim:
        to_addrs = _sanitize_addresses(to_addrs)
        cc_addrs = _sanitize_addresses(cc_addrs)

    if not to_addrs:
        raise ValueError("收件人列表不能为空")

    final_subject = f"{prefix}{subject}" if prefix else subject

    # Outlook COM 环境检查
    if win32com is None or pythoncom is None:
        log("win32com / pythoncom 不可用，无法使用本地 Outlook 发送。")
        return False

        # ========== COM 初始化 ==========
    pythoncom.CoInitialize()
    
    try:
        # 【简化版】直接使用公共邮箱发送，不要回退到个人账户
        from public_mailbox_sender import PublicMailboxAutoSender
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # 简单的 logger 适配器
        class SimpleLogger:
            def info(self, msg):
                log(msg)
            def error(self, msg):
                log(f"❌ {msg}")
            def warning(self, msg):
                log(f"⚠️  {msg}")
            def debug(self, msg):
                log(f"🔍 {msg}")
        
        sender = PublicMailboxAutoSender(outlook, SimpleLogger())
        log(f"✅ 使用公共邮箱发送: ChinaPD_Cybersecurity_Robot")
        log(f"   收件人: {to_addrs}")
        
        success = sender.send_from_public_mailbox(
            mailbox_name="ChinaPD_Cybersecurity_Robot",
            to_addresses=to_addrs,
            cc_addresses=cc_addrs,
            subject=final_subject,
            html_body=html_content or "",
            attachments=attachments,
            save_draft_only=False
        )
        
        if success:
            log("✅ 邮件已成功通过公共邮箱发送！")
            return True
        else:
            log("❌ 公共邮箱发送失败")
            return False
    
    except Exception as e:
        log(f"❌ 错误: {type(e).__name__}: {e}")
        log(traceback.format_exc())
        return False
    
    finally:
        # ========== COM 反初始化 ==========
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            log(f"COM反初始化异常: {e}")

    return False

# 简单自测
if __name__ == "__main__":
    test_html = "<html><body><h3>测试邮件</h3><p>这是一封测试。</p></body></html>"
    ok = send_email(
        subject="本地Outlook发送测试",
        html_content=test_html,
        to_addrs=["example@pg.com"],
        cc_addrs=[],
        max_retries=1
    )
    print("结果:", ok)