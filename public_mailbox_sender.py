#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
【增强版】公共邮箱发送模块
可直接集成到 EVP_RawDataAnalysis_Sender.py

使用方式:
    from public_mailbox_sender import PublicMailboxAutoSender
    
    sender = PublicMailboxAutoSender(outlook, logger)
    sender.send_from_public_mailbox(
        mailbox_name="ChinaPD_Cybersecurity_Robot",
        to_addresses=["recipient@pg.com"],
        subject="Report",
        html_body="<html>...</html>",
        cc_addresses=["cc@pg.com"],
        attachments=[]
    )
"""

import logging
import os
import json
from typing import List, Optional

class PublicMailboxAutoSender:
    """公共邮箱自动发送器"""
    
    def __init__(self, outlook, logger):
        """
        初始化发送器
        
        Args:
            outlook: Outlook.Application 对象
            logger: logging.Logger 对象
        """
        self.outlook = outlook
        self.logger = logger
        self.namespace = outlook.GetNamespace("MAPI")
    
    def send_from_public_mailbox(self, mailbox_name: str, 
                                 to_addresses: List[str],
                                 subject: str,
                                 html_body: str,
                                 cc_addresses: Optional[List[str]] = None,
                                 attachments: Optional[List[str]] = None,
                                 save_draft_only: bool = False) -> bool:
        """
        从公共邮箱发送邮件（完全自动化，无admin警告）
        
        重要改进：使用从个人账号创建邮件，然后通过MAPI属性设置From为公共邮箱
        这样避免了mail.Send()在公共邮箱Drafts中不工作的问题
        
        Args:
            mailbox_name: 公共邮箱名称（如: ChinaPD_Cybersecurity_Robot）
            to_addresses: 收件人列表
            subject: 邮件主题
            html_body: HTML格式的邮件正文
            cc_addresses: 抄送人列表
            attachments: 附件路径列表
            save_draft_only: 仅保存为草稿，不发送
        
        Returns:
            bool: 是否成功
        """
        try:
            self.logger.info(f"\n【从公共邮箱发送】{mailbox_name}")
            
            # 关键改进：从个人账号创建邮件，而不是从公共邮箱的Drafts
            # 这样mail.Send()才能正常工作
            
            # 步骤1: 找到个人账号Store
            personal_store = None
            for store in self.namespace.Stores:
                # 查找默认的个人账号（不是共享邮箱）
                if "shared" not in store.DisplayName.lower() and store.DisplayName != "SharePoint Lists":
                    personal_store = store
                    self.logger.info(f"  🔍 使用个人账号: {store.DisplayName}")
                    break
            
            if not personal_store:
                self.logger.error("  ❌ 找不到个人账号")
                return False
            
            # 步骤2: 在个人账号的Drafts中创建邮件（这是关键！必须明确指定Drafts）
            # 不能用CreateItem(0)，因为会默认使用当前活跃的邮箱
            # 必须明确在个人账号的Drafts中创建
            personal_drafts = personal_store.GetDefaultFolder(3)  # 3 = Drafts folder
            mail = personal_drafts.Items.Add()  # 在Drafts中创建
            self.logger.info(f"  📧 邮件已在 {personal_store.DisplayName} 的Drafts中创建")
            
            # 设置基本信息
            mail.Subject = subject
            mail.HTMLBody = html_body
            mail.To = ";".join(to_addresses) if to_addresses else ""
            if cc_addresses:
                mail.CC = ";".join(cc_addresses)
            
            # 步骤3: 应用MAPI修复 - 这是实现"从公共邮箱发送"的关键
            self._apply_mapi_fixes_for_public_mailbox(mail, mailbox_name)
            
            # 步骤4: 添加附件
            if attachments:
                for att_path in attachments:
                    if os.path.exists(att_path):
                        mail.Attachments.Add(att_path)
                        self.logger.info(f"  ✅ 附件: {os.path.basename(att_path)}")
            
            # 步骤5: 保存
            mail.Save()
            self.logger.info(f"  💾 邮件已保存")
            
            # 步骤6: 发送或仅保存
            if save_draft_only:
                self.logger.info(f"  📋 已保存为草稿（未发送）")
            else:
                mail.Send()
                self.logger.info(f"  ✅ 邮件已发送")
            
            return True
            
        except Exception as e:
            self.logger.error(f"❌ 发送失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False
    
    def _apply_mapi_fixes_for_public_mailbox(self, mail, mailbox_name: str):
        """
        应用MAPI修复以实现从公共邮箱发送
        
        关键改进：这个方法用于从个人账号发送但显示为来自公共邮箱的邮件
        通过设置From和SentRepresenting属性来实现
        """
        try:
            self.logger.info(f"  🔧 应用MAPI修复...")
            pa = mail.PropertyAccessor
            
            # 公共邮箱的邮箱地址
            public_email = f"{mailbox_name}@shared.pg.com"
            
            # 修复1: 设置PR_SENDER_NAME (0x0C06001F) = 公共邮箱名称
            try:
                pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C06001F", mailbox_name)
                self.logger.debug(f"    ✅ PR_SENDER_NAME = {mailbox_name}")
            except Exception as e:
                self.logger.debug(f"    ⚠️ PR_SENDER_NAME失败: {e}")
            
            # 修复2: 设置PR_SENDER_EMAIL_ADDRESS (0x0C1F001F) = 公共邮箱邮箱地址
            try:
                pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0C1F001F", public_email)
                self.logger.debug(f"    ✅ PR_SENDER_EMAIL_ADDRESS = {public_email}")
            except Exception as e:
                self.logger.debug(f"    ⚠️ PR_SENDER_EMAIL_ADDRESS失败: {e}")
            
            # 修复3: 设置PR_SENT_REPRESENTING_NAME (0x0042001F) = 公共邮箱
            try:
                pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0042001F", mailbox_name)
                self.logger.debug(f"    ✅ PR_SENT_REPRESENTING_NAME = {mailbox_name}")
            except Exception as e:
                self.logger.debug(f"    ⚠️ PR_SENT_REPRESENTING_NAME失败: {e}")
            
            # 修复4: 设置PR_SENT_REPRESENTING_EMAIL_ADDRESS (0x0044001F) = 公共邮箱邮箱
            try:
                pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0044001F", public_email)
                self.logger.debug(f"    ✅ PR_SENT_REPRESENTING_EMAIL_ADDRESS = {public_email}")
            except Exception as e:
                self.logger.debug(f"    ⚠️ PR_SENT_REPRESENTING_EMAIL_ADDRESS失败: {e}")
            
            # 修复5: 设置PR_MESSAGE_FLAGS (0x0E070003) - 标记为已提交
            try:
                flags = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003")
                new_flags = flags | 64  # MSGFLAG_SUBMIT = 64
                pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x0E070003", new_flags)
                self.logger.debug(f"    ✅ PR_MESSAGE_FLAGS: {flags} -> {new_flags}")
            except Exception as e:
                self.logger.debug(f"    ⚠️ PR_MESSAGE_FLAGS失败: {e}")
            
            self.logger.info(f"  ✅ MAPI修复完成")
            
        except Exception as e:
            self.logger.warning(f"  ⚠️ MAPI修复异常: {e}")


# 快速测试函数
def test_public_mailbox_sender():
    """快速测试"""
    import win32com.client
    
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)
    
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
        sender = PublicMailboxAutoSender(outlook, logger)
        
        # 测试发送
        success = sender.send_from_public_mailbox(
            mailbox_name="ChinaPD_Cybersecurity_Robot",
            to_addresses=["liang.wq.1@pg.com"],
            subject="✅ 模块测试 - 自动发送",
            html_body="""
            <html>
            <head><meta charset="utf-8"></head>
            <body>
                <h2>公共邮箱发送模块测试</h2>
                <p>✅ 测试成功！</p>
            </body>
            </html>
            """
        )
        
        if success:
            print("✅ 测试成功")
        else:
            print("❌ 测试失败")
            
    except Exception as e:
        logger.error(f"❌ 测试错误: {e}")
        import traceback
        print(traceback.format_exc())


if __name__ == "__main__":
    test_public_mailbox_sender()
