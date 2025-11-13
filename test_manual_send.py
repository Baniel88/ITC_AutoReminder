# 创建测试脚本：test_manual_send.py
# -*- coding: utf-8 -*-

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from email_sender import send_email
from datetime import datetime

# 测试手动发送
send_email(
    subject=f"手动发送测试 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
    html_content="<h1>这是一个手动发送测试</h1><p>点击Outlook中的发送按钮来发送此邮件</p>",
    to_addrs=["liang.wq.1@pg.com"],
    cc_addrs=["liang.wq.1@pg.com"],
    auto_send=False  # 预览模式
)

print("✅ 测试完成")