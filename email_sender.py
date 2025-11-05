# -*- coding: utf-8 -*-
"""
通用邮件发送模块（适配Outlook）
支持发送HTML格式邮件，可直接指定收件人和抄送人
"""

import win32com.client
import pythoncom
import time
import traceback
from datetime import datetime, timedelta
import json
import os


def load_email_config(config_path=None):
    """加载邮件配置文件（保留此函数，如需其他配置可继续使用）"""
    if not config_path:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_config.json")
    
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def send_email(subject, html_content, to_addrs, cc_addrs=None, config_path=None, max_retries=2, auto_send=False):
    """
    发送邮件主函数（更新版）
    subject: 邮件标题
    html_content: HTML格式的邮件内容
    to_addrs: 收件人列表（必填，如 ["a@pg.com", "b@pg.com"]）
    cc_addrs: 抄送人列表（可选，默认空列表）
    config_path: 配置文件路径（可选，如需其他配置可传入）
    max_retries: 最大重试次数
    auto_send: True=直接发送, False=打开预览窗口让用户确认后发送
    """
    # 处理默认参数
    if cc_addrs is None:
        cc_addrs = []
    
    # 验证收件人
    if not to_addrs:
        raise ValueError("收件人列表不能为空！")
    
    retry_count = 0
    while retry_count <= max_retries:
        try:
            pythoncom.CoInitialize()
            
            # 创建Outlook应用实例
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            
            # 获取当前默认账户
            current_account = outlook.Session.Accounts.Item(1)
            print(f"使用账户: {current_account.SmtpAddress} (尝试 {retry_count + 1}/{max_retries + 1})")
            
            # 创建邮件对象
            mail = outlook.CreateItem(0)
            mail.Subject = subject
            mail.HTMLBody = html_content
            mail.To = ";".join(to_addrs)  # 直接使用传入的收件人
            mail.CC = ";".join(cc_addrs)  # 直接使用传入的抄送人
            
            # 保存为草稿
            mail.Save()
            print(f"邮件已保存为草稿，收件人: {';'.join(to_addrs)}, 抄送: {';'.join(cc_addrs)}")
            
            # 记录发送时间
            send_time = datetime.now()
            print(f"发送时间: {send_time.strftime('%Y-%m-%d %H:%M:%S')}")
            
            if auto_send:
                # 直接发送邮件
                mail.Send()
                print("邮件已直接发送")
                
                # 等待发送确认
                time.sleep(5)
                
                # 检查已发送邮件
                sent_folder = ns.GetDefaultFolder(5)  # 5 = 已发送邮件
                sent_items = sent_folder.Items
                sent_items.Sort("[SentOn]", True)
                
                # 等待超时设置
                timeout = time.time() + 60
                found = False
                
                while time.time() < timeout and not found:
                    for item in sent_items:
                        try:
                            if (item.Subject == subject and 
                                item.SentOn >= send_time - timedelta(seconds=60)):
                                print(f"确认发送成功！邮件时间: {item.SentOn.strftime('%Y-%m-%d %H:%M:%S')}")
                                found = True
                                break
                        except Exception:
                            continue
                    
                    if not found:
                        time.sleep(2)
                
                if not found:
                    print("警告：未能确认邮件发送状态，但邮件可能已发送")
            else:
                # 显示邮件预览窗口，让用户确认后发送
                mail.Display(True)  # True表示模态窗口，用户必须处理完邮件才能继续
                print("邮件预览窗口已打开，请确认后手动发送")
                found = True  # 预览模式下认为操作成功
            
            if found:
                return True
            else:
                print(f"第 {retry_count + 1} 次尝试：未在已发送邮件中找到")
                retry_count += 1
                if retry_count <= max_retries:
                    print(f"等待5秒后重试...")
                    time.sleep(5)
            
        except Exception as e:
            print(f"第 {retry_count + 1} 次尝试失败: {str(e)}")
            print("错误详情:")
            print(traceback.format_exc())
            retry_count += 1
            if retry_count <= max_retries:
                print(f"等待10秒后重试...")
                time.sleep(10)
        finally:
            pythoncom.CoUninitialize()
    
    # 所有尝试失败，保存到草稿
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.HTMLBody = html_content
        mail.To = ";".join(to_addrs)
        mail.CC = ";".join(cc_addrs)
        mail.Save()
        print("所有发送尝试失败，邮件已保存到草稿箱，请手动发送")
    except Exception as e:
        print(f"保存草稿失败: {str(e)}")
    finally:
        pythoncom.CoUninitialize()
        
    return False


if __name__ == "__main__":
    # 测试代码（使用新参数调用）
    test_html = """
    <html>
        <body>
            <h1>测试邮件</h1>
            <p>这是一封来自EVP合规系统的测试邮件</p>
        </body>
    </html>
    """
    try:
        # 测试时直接传入收件人和抄送人
        send_email(
            subject="EVP Compliance 情况和任务提醒",
            html_content=test_html,
            to_addrs=["liang.wq.1@pg.com"],  # 收件人
            cc_addrs=["liang.wq.1@pg.com"]   # 抄送人
        )
    except Exception as e:
        print(f"测试发送失败: {e}")
