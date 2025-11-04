# -*- coding: utf-8 -*-
"""
处理待审核请求数据并生成邮件内容（包含Revoked状态说明文本）
保存HTML和TXT格式的邮件内容到指定文件夹
支持多报表类型和Teams通知，所有Category通过配置文件统一匹配
"""

import pandas as pd
import numpy as np
import sys, os
import json
import requests
from datetime import datetime, date
import re

import warnings
warnings.filterwarnings('ignore')

# 默认配置 - 当配置文件不存在或某些配置缺失时使用
DEFAULT_CONFIG = {
    "reports": {
        "Pending review任务提醒": {
            "recipients": [],
            "cc": [],
            "cc1": {},  # 所有Category的匹配规则在这里配置
            "system_config": {
                "MAX_REMAINING_DAYS_FOR_REPORT": 10,
                "URGENCY_LEVELS": {
                    "非常紧急": 2,
                    "紧急": 4,
                    "常规": 10
                },
                "ITC_REPORT_DIR_NAME": "ITC report",
                "RAW_DATA_DIR_NAME": "RawData",
                "REMINDER_DIR_NAME": "Reminder",
                "LOG_DIR_NAME": "Log",
                "EMAIL_SUBJECT_PENDING": "Pending review任务提醒",
                "EMAIL_SUBJECT_REVOKED": "Revoked状态任务提醒",
                "EMAIL_ExitForm_REVOKED": "SSO的应用/加入域的系统，可以在1年内在系统里面移除并确认,否则24小时移除",
                "EMAIL_RoleChange_REVOKED": "请尽快和user联系，尽快处理，在30天内在ITC确认",
                "ITC_SYSTEM_LINK": "https://itc-tool.pg.com/ComplianceReport?siteId=193"
            }
        },
        "Revoked状态任务提醒": {
            "recipients": [],
            "cc": [],
            "cc1": {}  # 所有Category的匹配规则在这里配置
        }
        # 注释：已移除EVP Compliance相关配置（无此类型数据）
        # "EVP Compliance 情况和任务提醒": {
        #     "recipients": [],
        #     "cc": []
        # }
    },
    "Teams": {
        "webhook_url": "https://teams.microsoft.com/l/channel/19%3AcIJ0wUDE2YQPxiX78Gd2PNeT2r2tPtte9Bc90hoMz_k1%40thread.tacv2/General?groupId=b320358b-da36-47e8-9007-21fecd43e383&tenantId=3596192b-fdf5-4e2c-a6fa-acb706c963d8"
    }
}


def load_config(config_path=None):
    """加载配置文件，如果文件不存在则创建默认配置文件"""
    if not config_path:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(current_dir, "email_config.json")
    
    if not os.path.exists(config_path):
        print(f"配置文件不存在，将在 {config_path} 创建默认配置文件")
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return DEFAULT_CONFIG.copy()
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            user_config = json.load(f)
        
        config = DEFAULT_CONFIG.copy()
        
        # 加载所有报表配置
        if "reports" in user_config:
            for report_name, report_config in user_config["reports"].items():
                if report_name not in config["reports"]:
                    config["reports"][report_name] = {"recipients": [], "cc": [], "cc1": {}}
                
                if "recipients" in report_config:
                    config["reports"][report_name]["recipients"] = report_config["recipients"]
                
                if "cc" in report_config:
                    config["reports"][report_name]["cc"] = report_config["cc"]
                
                if "cc1" in report_config:
                    config["reports"][report_name]["cc1"] = report_config["cc1"]  # 读取所有Category的匹配规则
                
                if "system_config" in report_config:
                    if "system_config" not in config["reports"][report_name]:
                        config["reports"][report_name]["system_config"] = {}
                    for key, value in report_config["system_config"].items():
                        config["reports"][report_name]["system_config"][key] = value
        
        # 提取系统配置到顶层（保持兼容性）
        pending_system_config = config["reports"].get("Pending review任务提醒", {}).get("system_config", {})
        for key, value in pending_system_config.items():
            config[key] = value
        
        # 加载Teams配置
        if "Teams" in user_config:
            config["Teams"] = user_config["Teams"]
        
        # 确保必要配置存在
        required_system_configs = [
            "MAX_REMAINING_DAYS_FOR_REPORT", "URGENCY_LEVELS",
            "ITC_REPORT_DIR_NAME", "RAW_DATA_DIR_NAME",
            "REMINDER_DIR_NAME", "LOG_DIR_NAME",
            "EMAIL_SUBJECT_PENDING", "EMAIL_SUBJECT_REVOKED",
            "EMAIL_ExitForm_REVOKED", "EMAIL_RoleChange_REVOKED",
            "ITC_SYSTEM_LINK"
        ]
        
        for config_key in required_system_configs:
            if config_key not in config:
                print(f"警告：配置中缺少'{config_key}'，将使用默认值")
                config[config_key] = DEFAULT_CONFIG["reports"]["Pending review任务提醒"]["system_config"][config_key]
        
        # 确保紧急程度配置完整
        for level in ['非常紧急', '紧急', '常规']:
            if level not in config['URGENCY_LEVELS']:
                print(f"警告：配置文件中缺少'{level}'的紧急程度配置，将使用默认值")
                config['URGENCY_LEVELS'][level] = DEFAULT_CONFIG["reports"]["Pending review任务提醒"]["system_config"]["URGENCY_LEVELS"][level]
        
        return config
    except Exception as e:
        print(f"读取配置文件出错: {str(e)}，将使用默认配置")
        return DEFAULT_CONFIG.copy()


# 加载配置
CONFIG = load_config()


# 辅助函数：确保目录存在
def ensure_directory_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"已创建文件夹: {directory}")
    return directory


# 辅助函数：确保邮箱地址包含@pg.com后缀
def ensure_pg_email(email, username=None):
    if pd.isna(email) or email.strip() == '':
        if username and username.strip() != '':
            return f"{username.strip().lower().replace(' ', '.')}@pg.com"
        return ''
    
    email = email.strip()
    if '@' not in email:
        return f"{email.lower()}@pg.com"
    elif not email.endswith('@pg.com'):
        local_part = email.split('@')[0]
        return f"{local_part}@pg.com"
    return email


# 通用的cc1邮箱匹配函数（所有Category统一逻辑）
def get_matching_cc1_emails(category, report_type, strict_match=False):
    """
    根据分类(Category)和报表类型获取匹配的cc1邮箱列表
    匹配规则完全由email_config.json中的cc1配置决定
    """
    report_config = CONFIG["reports"].get(report_type, {})
    if "cc1" not in report_config:
        print(f"调试：{report_type} 未配置cc1，Category={category}")
        return []
    
    cc1_config = report_config["cc1"]
    category_clean = str(category).strip() if pd.notna(category) else ""
    
    # 1. 优先精确匹配配置中的key
    if category_clean in cc1_config:
        matched_emails = [ensure_pg_email(email) for email in cc1_config[category_clean]]
        print(f"调试：Category={category_clean} 精确匹配到cc1邮箱: {matched_emails}")
        return matched_emails
    
    # 2. 非严格模式下进行模糊匹配（所有Category统一逻辑）
    if not strict_match:
        for key in cc1_config:
            key_clean = str(key).strip()
            if (key_clean.lower() in category_clean.lower()) or (category_clean.lower() in key_clean.lower()):
                matched_emails = [ensure_pg_email(email) for email in cc1_config[key]]
                print(f"调试：Category={category_clean} 模糊匹配到cc1键={key_clean}，邮箱: {matched_emails}")
                return matched_emails
    
    print(f"调试：Category={category_clean} 未匹配到cc1邮箱")
    return []


# 辅助函数：安全解析日期列
def safe_parse_dates(df, date_columns):
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(
                df[col], 
                errors='coerce'
            )
    return df


# 数据读取与处理
def load_and_process_data(csv_file_path):
    df = pd.read_csv(
        csv_file_path,
        na_values=['', ' ', 'NA'],
        keep_default_na=True
    )

    df = df.replace(r'^\s*$', np.nan, regex=True)
    
    date_columns = ['Requested Date', 'Expiration Date', 'Log Date']
    df = safe_parse_dates(df, date_columns)

    for col in date_columns:
        if col in df.columns:
            na_count = df[col].isna().sum()
            if na_count > 0:
                print(f"警告：{col}列中有 {na_count} 个值无法解析为日期")

    df['is_new_request'] = df['Requester'].notna().astype(int)
    df['request_group'] = df['is_new_request'].cumsum()

    columns_to_fill = [
        'Requester', 'Requester Email', 'Request For', 'Request For Email', 
        'Requested Date', 'Area', 'Category', 'Category Description',
        'System/Solution', 'System/Solution Description', 'Approval Text',
        'Owner Guidelines', 'Expiration Date', 'Max Request Age (Days)',
        'Access Type', 'Temporary Access?', 'Privileged?', 'Status',
        'Confirmed?', 'Reason', 'Remark/Role', 'Employee Status',
        'Log Actor', 'Log Status', 'Log Date'
    ]

    for col in columns_to_fill:
        if col in df.columns:
            df[col] = df.groupby('request_group')[col].transform(
                lambda x: x.ffill().bfill()
            )
        
    return df


# 分析不同类型的请求
def analyze_requests(df):
    required_columns = ['Status', 'System/Solution', 'Request For', 'Category']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"数据中缺少必要的列：{col}")
    
    if 'Expiration Date' not in df.columns:
        print("警告：数据中没有'Expiration Date'列，将使用默认值")

    # 1. 处理Pending Review状态
    pending_df = df[
        (df['Status'] == 'Pending Review') &
        (df['System/Solution'].notna()) &
        (df['Request For'].notna()) &
        (df['Category'].notna())
    ].copy()
    
    # 2. 处理包含Revoked关键词的状态
    revoked_df = df[
        (df['Status'].str.contains('Revoked', case=False, na=False)) &
        (df['System/Solution'].notna()) &
        (df['Request For'].notna()) &
        (df['Category'].notna())
    ].copy()

    # 注释：已移除EVP Compliance数据筛选（无此类型数据）
    # 3. 处理EVP Compliance相关请求
    # evp_df = df[
    #     (df['Category'].str.contains('EVP Compliance', case=False, na=False)) &
    #     (df['System/Solution'].notna()) &
    #     (df['Request For'].notna())
    # ].copy()

    current_date = date.today()
    
    # 处理各类数据
    pending_results = process_pending_requests(pending_df, current_date)
    revoked_results = process_revoked_requests(revoked_df, current_date)
    # 注释：已移除EVP Compliance数据处理（无此类型数据）
    # evp_results = process_evp_compliance_requests(evp_df, current_date)
    
    return {
        'pending': pending_results,
        'revoked': revoked_results
        # 注释：已移除EVP Compliance结果返回（无此类型数据）
        # 'evp_compliance': evp_results
    }


def process_pending_requests(pending_df, current_date):
    """处理待审核请求"""
    report_type = "Pending review任务提醒"
    
    if pending_df.empty:
        print("信息：没有待审核的请求数据")
        empty_df = pd.DataFrame(columns=[
            'Action Owner', 'Action Owner Email', 'System Name', 'Category',
            '剩余天数', '紧急程度', 'Pending_review数量'
        ])
        return {
            'table': empty_df,
            'total_count': 0,
            'recipients': CONFIG["reports"][report_type].get('recipients', []),
            'cc': CONFIG["reports"][report_type].get('cc', []),
            'type': report_type
        }
    
    if 'Expiration Date' in pending_df.columns:
        pending_df['expiration_date_only'] = pending_df['Expiration Date'].dt.date
    else:
        pending_df['expiration_date_only'] = pd.NaT

    # 计算剩余天数
    def calculate_remaining_days(row):
        if pd.notna(row['expiration_date_only']):
            return max((row['expiration_date_only'] - current_date).days, 0)
        return CONFIG["MAX_REMAINING_DAYS_FOR_REPORT"]

    pending_df['剩余天数'] = pending_df.apply(calculate_remaining_days, axis=1).astype(int)
    
    # 应用邮件发送阈值
    filtered_df = pending_df[pending_df['剩余天数'] <= CONFIG["MAX_REMAINING_DAYS_FOR_REPORT"]].copy()
    excluded_count = len(pending_df) - len(filtered_df)
    if excluded_count > 0:
        print(f"根据配置，已排除 {excluded_count} 条剩余天数超过 {CONFIG['MAX_REMAINING_DAYS_FOR_REPORT']} 天的Pending请求")
    
    if filtered_df.empty:
        print(f"信息：没有剩余天数小于等于 {CONFIG['MAX_REMAINING_DAYS_FOR_REPORT']} 天的待审核请求")
        empty_df = pd.DataFrame(columns=[
            'Action Owner', 'Action Owner Email', 'System Name', 'Category',
            '剩余天数', '紧急程度', 'Pending_review数量'
        ])
        return {
            'table': empty_df,
            'total_count': 0,
            'recipients': CONFIG["reports"][report_type].get('recipients', []),
            'cc': CONFIG["reports"][report_type].get('cc', []),
            'type': report_type
        }

    # 标记紧急程度
    def mark_urgency(days):
        if days <= CONFIG['URGENCY_LEVELS']['非常紧急']:
            return '非常紧急'
        elif days <= CONFIG['URGENCY_LEVELS']['紧急']:
            return '紧急'
        elif days <= CONFIG['URGENCY_LEVELS']['常规']:
            return '常规'
        return '常规'

    filtered_df['紧急程度'] = filtered_df['剩余天数'].apply(mark_urgency)

    # 获取负责人和邮箱
    def get_action_owner_and_email(group):
        try:
            approval_statuses = ['Approved', 'PartiallyApproved']
            if 'Log Status' in group.columns:
                approval_records = group[group['Log Status'].isin(approval_statuses)]
            else:
                approval_records = pd.DataFrame()
            
            if not approval_records.empty and 'Log Date' in approval_records.columns:
                latest_approval = approval_records.loc[approval_records['Log Date'].idxmax()]
                email = ensure_pg_email(
                    latest_approval.get('Log Actor Email', ''),
                    latest_approval.get('Log Actor', '')
                )
                return pd.Series({
                    'Action Owner': latest_approval.get('Log Actor', '未知负责人'),
                    'Action Owner Email': email
                })
            else:
                requester = group['Requester'].iloc[0] if 'Requester' in group.columns and pd.notna(group['Requester'].iloc[0]) else '未知请求者'
                requester_email = group['Requester Email'].iloc[0] if 'Requester Email' in group.columns and pd.notna(group['Requester Email'].iloc[0]) else ''
                email = ensure_pg_email(requester_email, requester)
                return pd.Series({
                    'Action Owner': requester,
                    'Action Owner Email': email
                })
        except Exception as e:
            print(f"提取负责人信息时出错: {str(e)}")
            return pd.Series({
                'Action Owner': '提取失败',
                'Action Owner Email': ''
            })

    try:
        grouped = filtered_df.groupby('request_group')
        owner_info = grouped.apply(
            lambda x: pd.Series({
                'Action Owner': get_action_owner_and_email(x)['Action Owner'],
                'Action Owner Email': get_action_owner_and_email(x)['Action Owner Email'],
                'System/Solution': x['System/Solution'].iloc[0] if pd.notna(x['System/Solution'].iloc[0]) else '未知系统',
                'Category': x['Category'].iloc[0] if pd.notna(x['Category'].iloc[0]) else '未知分类',
                '剩余天数': x['剩余天数'].iloc[0],
                '紧急程度': x['紧急程度'].iloc[0]
            })
        ).reset_index(drop=True)
        
        for col in ['Action Owner', 'Action Owner Email', 'System/Solution', 'Category', '剩余天数', '紧急程度']:
            if col not in owner_info.columns:
                owner_info[col] = '未知' if col != '剩余天数' else 0
                
    except Exception as e:
        print(f"分组处理时出错: {str(e)}")
        owner_info = pd.DataFrame({
            'Action Owner': ['处理错误'],
            'Action Owner Email': [''],
            'System/Solution': ['未知系统'],
            'Category': ['未知分类'],
            '剩余天数': [0],
            '紧急程度': ['常规']
        })

    target_table = owner_info.groupby(
        ['Action Owner', 'Action Owner Email', 'System/Solution', 'Category', '剩余天数', '紧急程度']
    ).size().reset_index(name='Pending_review数量')

    target_table.rename(columns={'System/Solution': 'System Name'}, inplace=True)

    target_table['Action Owner'] = target_table['Action Owner'].fillna('未知负责人')
    target_table['Action Owner Email'] = target_table['Action Owner Email'].fillna('')
    target_table['System Name'] = target_table['System Name'].fillna('未知系统')
    target_table['Category'] = target_table['Category'].fillna('未知分类')

    total_count = target_table['Pending_review数量'].sum()
    total_row = pd.DataFrame([{
        'Action Owner': '总计',
        'Action Owner Email': '',
        'System Name': '',
        'Category': '',
        '剩余天数': '',
        '紧急程度': '',
        'Pending_review数量': total_count
    }])
    target_table = pd.concat([target_table, total_row], ignore_index=True)

    # 获取收件人和抄送（所有Category统一匹配逻辑）
    config_recipients = CONFIG["reports"][report_type].get("recipients", [])
    config_cc = CONFIG["reports"][report_type].get("cc", [])
    data_recipients = target_table[target_table['Action Owner Email'] != '']['Action Owner Email'].drop_duplicates().tolist()
    categories = target_table[target_table['Category'] != '']['Category'].drop_duplicates().tolist()
    cc1_emails = []
    
    # 所有Category使用统一匹配规则（由email_config.json中的cc1决定）
    for category in categories:
        # 严格匹配模式可通过配置控制，这里默认关闭（如需开启可改为True）
        cc1_for_category = get_matching_cc1_emails(category, report_type, strict_match=False)
        cc1_emails.extend(cc1_for_category)
    
    valid_recipients = list(set([ensure_pg_email(email) for email in config_recipients + data_recipients if email.strip()]))
    valid_cc = list(set([ensure_pg_email(email) for email in config_cc + cc1_emails if email.strip()]))
    valid_recipients = [email for email in valid_recipients if email]
    valid_cc = [email for email in valid_cc if email]
    
    # 输出调试信息（所有Category统一日志格式）
    print(f"调试：{report_type} 最终收件人: {valid_recipients}")
    print(f"调试：{report_type} 最终抄送: {valid_cc}")
    
    return {
        'table': target_table,
        'total_count': total_count,
        'recipients': valid_recipients,
        'cc': valid_cc,
        'type': report_type
    }


def process_revoked_requests(revoked_df, current_date):
    """处理Revoked状态请求"""
    report_type = "Revoked状态任务提醒"
    
    if revoked_df.empty:
        print("信息：没有包含Revoked关键词的请求数据")
        empty_df = pd.DataFrame(columns=[
            'Action Owner', 'Action Owner Email', 'System Name', 'Category',
            '状态', 'Revoked数量', '状态说明'
        ])
        return {
            'table': empty_df,
            'total_count': 0,
            'recipients': CONFIG["reports"][report_type].get('recipients', []),
            'cc': CONFIG["reports"][report_type].get('cc', []),
            'type': report_type
        }
    
    if 'Log Date' in revoked_df.columns:
        if not pd.api.types.is_datetime64_any_dtype(revoked_df['Log Date']):
            revoked_df['Log Date'] = pd.to_datetime(
                revoked_df['Log Date'], 
                errors='coerce'
            )
    
    # 根据Status添加状态说明
    def get_status_note(status):
        if pd.isna(status):
            return ""
        status_lower = str(status).lower()
        if 'exitform' in status_lower:
            return CONFIG["EMAIL_ExitForm_REVOKED"]
        elif 'rolechange' in status_lower:
            return CONFIG["EMAIL_RoleChange_REVOKED"]
        return ""
    
    revoked_df['状态说明'] = revoked_df['Status'].apply(get_status_note)
    
    # 获取负责人和邮箱
    def get_confirmed_actor_and_email(group):
        try:
            if 'Log Status' in group.columns:
                confirmed_records = group[group['Log Status'].str.contains('confirmed', case=False, na=False)]
            else:
                confirmed_records = pd.DataFrame()
            
            if not confirmed_records.empty and 'Log Date' in confirmed_records.columns:
                latest_confirmed = confirmed_records.loc[confirmed_records['Log Date'].idxmax()]
                email = ensure_pg_email(
                    latest_confirmed.get('Log Actor Email', ''),
                    latest_confirmed.get('Log Actor', '')
                )
                return pd.Series({
                    'Action Owner': latest_confirmed.get('Log Actor', '未知负责人'),
                    'Action Owner Email': email
                })
            else:
                if 'Log Date' in group.columns and not group['Log Date'].isna().all():
                    latest_log = group.loc[group['Log Date'].idxmax()]
                    email = ensure_pg_email(
                        latest_log.get('Log Actor Email', ''),
                        latest_log.get('Log Actor', '')
                    )
                    return pd.Series({
                        'Action Owner': latest_log.get('Log Actor', '未知负责人'),
                        'Action Owner Email': email
                    })
                else:
                    requester = group['Requester'].iloc[0] if 'Requester' in group.columns and pd.notna(group['Requester'].iloc[0]) else '未知请求者'
                    requester_email = group['Requester Email'].iloc[0] if 'Requester Email' in group.columns and pd.notna(group['Requester Email'].iloc[0]) else ''
                    email = ensure_pg_email(requester_email, requester)
                    return pd.Series({
                        'Action Owner': requester,
                        'Action Owner Email': email
                    })
        except Exception as e:
            print(f"提取Revoked负责人信息时出错: {str(e)}")
            return pd.Series({
                'Action Owner': '提取失败',
                'Action Owner Email': ''
            })

    try:
        grouped = revoked_df.groupby('request_group')
        owner_info = grouped.apply(
            lambda x: pd.Series({
                'Action Owner': get_confirmed_actor_and_email(x)['Action Owner'],
                'Action Owner Email': get_confirmed_actor_and_email(x)['Action Owner Email'],
                'System/Solution': x['System/Solution'].iloc[0] if pd.notna(x['System/Solution'].iloc[0]) else '未知系统',
                'Category': x['Category'].iloc[0] if pd.notna(x['Category'].iloc[0]) else '未知分类',
                'Status': x['Status'].iloc[0] if pd.notna(x['Status'].iloc[0]) else 'Revoked',
                '状态说明': x['状态说明'].iloc[0] if pd.notna(x['状态说明'].iloc[0]) else ''
            })
        ).reset_index(drop=True)
        
        for col in ['Action Owner', 'Action Owner Email', 'System/Solution', 'Category', 'Status', '状态说明']:
            if col not in owner_info.columns:
                owner_info[col] = '未知' if col != '状态说明' else ''
                
    except Exception as e:
        print(f"Revoked分组处理时出错: {str(e)}")
        owner_info = pd.DataFrame({
            'Action Owner': ['处理错误'],
            'Action Owner Email': [''],
            'System/Solution': ['未知系统'],
            'Category': ['未知分类'],
            'Status': 'Revoked',
            '状态说明': ''
        })

    target_table = owner_info.groupby(
        ['Action Owner', 'Action Owner Email', 'System/Solution', 'Category', 'Status', '状态说明']
    ).size().reset_index(name='Revoked数量')

    target_table.rename(columns={'System/Solution': 'System Name'}, inplace=True)

    target_table['Action Owner'] = target_table['Action Owner'].fillna('未知负责人')
    target_table['Action Owner Email'] = target_table['Action Owner Email'].fillna('')
    target_table['System Name'] = target_table['System Name'].fillna('未知系统')
    target_table['Category'] = target_table['Category'].fillna('未知分类')
    target_table['状态说明'] = target_table['状态说明'].fillna('')

    total_count = target_table['Revoked数量'].sum()
    total_row = pd.DataFrame([{
        'Action Owner': '总计',
        'Action Owner Email': '',
        'System Name': '',
        'Category': '',
        'Status': '',
        '状态说明': '',
        'Revoked数量': total_count
    }])
    target_table = pd.concat([target_table, total_row], ignore_index=True)

    # 获取收件人和抄送（所有Category统一匹配逻辑）
    config_recipients = CONFIG["reports"][report_type].get("recipients", [])
    config_cc = CONFIG["reports"][report_type].get("cc", [])
    data_recipients = target_table[target_table['Action Owner Email'] != '']['Action Owner Email'].drop_duplicates().tolist()
    categories = target_table[target_table['Category'] != '']['Category'].drop_duplicates().tolist()
    cc1_emails = []
    
    # 所有Category使用统一匹配规则（由email_config.json中的cc1决定）
    for category in categories:
        # 严格匹配模式可通过配置控制，这里默认关闭
        cc1_for_category = get_matching_cc1_emails(category, report_type, strict_match=False)
        cc1_emails.extend(cc1_for_category)
    
    valid_recipients = list(set([ensure_pg_email(email) for email in config_recipients + data_recipients if email.strip()]))
    valid_cc = list(set([ensure_pg_email(email) for email in config_cc + cc1_emails if email.strip()]))
    valid_recipients = [email for email in valid_recipients if email]
    valid_cc = [email for email in valid_cc if email]
    
    # 输出调试信息（所有Category统一日志格式）
    print(f"调试：{report_type} 最终收件人: {valid_recipients}")
    print(f"调试：{report_type} 最终抄送: {valid_cc}")
    
    return {
        'table': target_table,
        'total_count': total_count,
        'recipients': valid_recipients,
        'cc': valid_cc,
        'type': report_type
    }



# 生成邮件HTML内容
def generate_email_html(table_data, current_date, total_count, report_type, recipients, cc):
    urgency_styles = {
        '非常紧急': 'background-color:#ffebee;color:#b71c1c;font-weight:bold',
        '紧急': 'background-color:#fff3e0;color:#e65100',
        '常规': 'background-color:#e8f5e9;color:#2e7d32'
    }
    
    # 根据报表类型设置主题和标题文本
    if report_type == "Pending review任务提醒":
        subject = CONFIG.get('EMAIL_SUBJECT_PENDING', report_type)
        header_text = f"系统检测到当前有 <strong>{total_count}</strong> 条待审核（Pending Review）请求（剩余天数≤{CONFIG['MAX_REMAINING_DAYS_FOR_REPORT']}天）。为避免权限过期影响业务正常运行，请及时处理您负责的审核任务。"
        table_headers = ['负责人（Action Owner）', '负责人邮箱 (@pg.com)', '系统名称（System Name）', '分类（Category）', '剩余天数', '紧急程度', '待审核数量']
    elif report_type == "Revoked状态任务提醒":
        subject = CONFIG.get('EMAIL_SUBJECT_REVOKED', report_type)
        header_text = f"系统检测到当前有 <strong>{total_count}</strong> 条状态包含'Revoked'关键词的请求。请关注并处理这些请求。"
        table_headers = ['负责人（Action Owner）', '负责人邮箱 (@pg.com)', '系统名称（System Name）', '分类（Category）', '状态', '状态说明', 'Revoked数量']
    # 注释：已移除EVP Compliance邮件逻辑（无此类型数据）
    # elif report_type == "EVP Compliance 情况和任务提醒":
    #     subject = report_type
    #     header_text = f"系统检测到当前有 <strong>{total_count}</strong> 条EVP Compliance相关请求。请关注并处理这些请求。"
    #     table_headers = ['负责人（Action Owner）', '负责人邮箱 (@pg.com)', '系统名称（System Name）', '分类（Category）', '状态', '请求数量']
    else:
        subject = report_type
        header_text = f"系统检测到当前有 <strong>{total_count}</strong> 条相关请求。请关注并处理这些请求。"
        table_headers = ['负责人（Action Owner）', '负责人邮箱 (@pg.com)', '系统名称（System Name）', '分类（Category）', '状态', '请求数量']
    
    recipients_str = ", ".join(recipients) if recipients else "无"
    cc_str = ", ".join(cc) if cc else "无"
    
    html = f"""<html>
<head>
    <meta charset="UTF-8">
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 1200px; margin: 0 auto; padding: 20px; }}
        .email-header {{ margin-bottom: 25px; padding-bottom: 15px; border-bottom: 2px solid #f0f0f0; }}
        h2 {{ color: #2c3e50; margin-top: 0; }}
        .urgency-desc {{ margin: 20px 0; padding: 15px; background-color: #f5f5f5; border-radius: 6px; }}
        .recipients-info {{ margin: 15px 0; padding: 10px; background-color: #f9f9f9; border-radius: 4px; }}
        table {{ width: 100%; border-collapse: separate; border-spacing: 0; margin: 25px 0; box-shadow: 0 2px 5px rgba(0,0,0,0.05); border-radius: 6px; overflow: hidden; }}
        th {{ 
            padding: 12px 15px; 
            text-align: left; 
            background-color: #2c3e50; 
            color: white; 
            font-weight: bold; 
            position: sticky; 
            top: 0;
        }}
        td {{ 
            padding: 12px 15px; 
            text-align: left; 
            border-bottom: 1px solid #f0f0f0;
        }}
        tr:hover {{ background-color: #f9f9f9; }}
        tr:last-child td {{ border-bottom: none; }}
        .total-row {{ 
            background-color: #f1f8e9; 
            font-weight: bold; 
            border-top: 2px solid #2c3e50;
        }}
        .badge {{ 
            padding: 5px 10px; 
            border-radius: 15px; 
            font-size: 0.85em; 
            display: inline-block;
        }}
        .footer-note {{ margin-top: 30px; padding-top: 15px; border-top: 2px solid #f0f0f0; color: #666; }}
        .itc-link {{ color: #2196f3; text-decoration: underline; font-weight: bold; }}
        .status-note {{ color: #d32f2f; font-style: italic; }}
    </style>
</head>
<body>
    <div class="email-header">
        <h2>{subject} - {current_date}</h2>
        <p>{header_text}</p>
        <p>请通过Chrome 浏览器（其他浏览器可能存在兼容问题）登录<a href="{CONFIG['ITC_SYSTEM_LINK']}" class="itc-link" target="_blank">ITC系统链接</a>，点击"MyTasks"/"MyActions"完成相关任务处理。</p>
    </div>
    
    <div class="recipients-info">
        <p><strong>邮件接收人 (To):</strong> {recipients_str}</p>
        <p><strong>抄送 (CC):</strong> {cc_str}</p>
    </div>
    """
    
    # 根据报表类型添加不同的说明内容
    if report_type == "Pending review任务提醒":
        html += f"""
    <div class="urgency-desc">
        <p><strong>紧急程度说明：</strong></p>
        <ul>
            <li><span style="{urgency_styles['非常紧急']};padding:2px 6px;border-radius:3px">非常紧急</span>：剩余天数≤{CONFIG['URGENCY_LEVELS']['非常紧急']}天（需立即处理）</li>
            <li><span style="{urgency_styles['紧急']};padding:2px 6px;border-radius:3px">紧急</span>：剩余天数≤{CONFIG['URGENCY_LEVELS']['紧急']}天（建议当天处理）</li>
            <li><span style="{urgency_styles['常规']};padding:2px 6px;border-radius:3px">常规</span>：剩余天数≤{CONFIG['URGENCY_LEVELS']['常规']}天（请在过期前完成）</li>
        </ul>
    </div>
        """
    elif report_type == "Revoked状态任务提醒":
        html += f"""
    <div class="urgency-desc">
        <p><strong>状态说明：</strong></p>
        <ul>
            <li><strong>ExitForm相关：</strong> {CONFIG['EMAIL_ExitForm_REVOKED']}</li>
            <li><strong>RoleChange相关：</strong> {CONFIG['EMAIL_RoleChange_REVOKED']}</li>
        </ul>
    </div>
        """
    # 注释：已移除EVP Compliance说明内容（无此类型数据）
    # elif report_type == "EVP Compliance 情况和任务提醒":
    #     html += f"""
    # <div class="urgency-desc">
    #     <p><strong>EVP Compliance说明：</strong></p>
    #     <p>此报表包含所有与EVP合规性相关的请求，请按照公司合规政策及时处理。</p>
    # </div>
    #     """
    
    # 生成汇总表格
    html += f"""
    <table>
        <thead>
            <tr>
    """
    for header in table_headers:
        html += f"<th>{header}</th>"
    
    html += f"""
            </tr>
        </thead>
        <tbody>
    """
    
    for _, row in table_data.iterrows():
        if row['Action Owner'] == '总计':
            html += "<tr class='total-row'>"
            html += f"<td>{row['Action Owner']}</td>"
            for _ in range(len(table_headers) - 2):
                html += "<td></td>"
            if report_type == "Pending review任务提醒":
                html += f"<td>{row.get('Pending_review数量', '')}</td>"
            elif report_type == "Revoked状态任务提醒":
                html += f"<td>{row.get('Revoked数量', '')}</td>"
            # 注释：已移除EVP Compliance总计行逻辑（无此类型数据）
            # else:
            #     html += f"<td>{row.get('请求数量', '')}</td>"
            html += "</tr>"
        else:
            html += "<tr>"
            html += f"<td>{row['Action Owner']}</td>"
            html += f"<td>{row['Action Owner Email']}</td>"
            html += f"<td>{row['System Name']}</td>"
            html += f"<td>{row['Category']}</td>"
            
            if report_type == "Pending review任务提醒":
                html += f"<td>{row['剩余天数']}天</td>"
                urgency = row['紧急程度']
                html += f"<td><span class='badge' style='{urgency_styles[urgency]}'>{urgency}</span></td>"
                html += f"<td>{row['Pending_review数量']}</td>"
            elif report_type == "Revoked状态任务提醒":
                html += f"<td>{row['Status']}</td>"
                note = row['状态说明']
                if note:
                    html += f"<td class='status-note'>{note}</td>"
                else:
                    html += f"<td>{note}</td>"
                html += f"<td>{row['Revoked数量']}</td>"
            # 注释：已移除EVP Compliance表格行逻辑（无此类型数据）
            # else:  # EVP和其他报表类型
            #     html += f"<td>{row['Status']}</td>"
            #     html += f"<td>{row['请求数量']}</td>"
                
            html += "</tr>"
    
    html += """
        </tbody>
    </table>
    
    <div class="footer-note">
        <p>感谢您的及时处理！如有任何问题，请联系Site CSL团队支持。</p>
        <p>此致<br>GC PD 网络安全团队</p>
    </div>
</body>
</html>
"""
    return html, subject


# 从HTML提取纯文本内容
def html_to_text(html_content):
    text = re.sub(r'<style.*?</style>', '', html_content, flags=re.DOTALL)
    text = re.sub(r'<script.*?</script>', '', text, flags=re.DOTALL)
    text = re.sub(r'<.*?>', '', text)
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    return text


# 保存邮件内容为HTML和TXT到指定文件夹
def save_email_contents(html_content, output_dir, report_type):
    ensure_directory_exists(output_dir)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    prefix = report_type.replace(' ', '_').replace('任务提醒', '').lower()
    html_path = os.path.join(output_dir, f"{prefix}_email_{timestamp}.html")
    txt_path = os.path.join(output_dir, f"{prefix}_email_{timestamp}.txt")
    
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    text_content = html_to_text(html_content)
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(text_content)
    
    return html_path, txt_path


# 记录日志到Log文件夹
def log_message(message, log_dir):
    ensure_directory_exists(log_dir)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_date = datetime.now().strftime('%Y%m%d')
    log_file = os.path.join(log_dir, f"process_{log_date}.log")
    
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] {message}\n")


# 发送Teams消息
def send_teams_message(message, report_type, webhook_url=None):
    """发送消息到Teams频道"""
    if not webhook_url:
        webhook_url = CONFIG.get("Teams", {}).get("webhook_url", "")
    
    if not webhook_url or "请替换为你的" in webhook_url:
        print("Teams webhook URL未配置，跳过发送Teams消息")
        return False
    
    try:
        payload = {
            "@type": "MessageCard",
            "@context": "http://schema.org/extensions",
            "themeColor": "0076D7",
            "summary": f"{report_type}通知",
            "sections": [{
                "activityTitle": f"{report_type}",
                "activitySubtitle": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "text": message
            }]
        }
        
        response = requests.post(
            webhook_url,
            json=payload,
            headers={"Content-Type": "application/json"}
        )
        
        if response.status_code == 200:
            print(f"Teams消息发送成功: {report_type}")
            return True
        else:
            print(f"Teams消息发送失败，状态码: {response.status_code}, 响应: {response.text}")
            return False
    except Exception as e:
        print(f"发送Teams消息时出错: {str(e)}")
        return False


# 发送邮件的辅助函数
def send_report(report_data, reminder_dir, log_dir):
    report_type = report_data['type']
    total_count = report_data['total_count']
    recipients = report_data['recipients']
    cc = report_data['cc']
    table_data = report_data['table']
    
    if total_count == 0:
        return
    
    current_date_str = datetime.now().strftime('%Y年%m月%d日')
    
    generate_msg = f"生成{report_type}类型邮件内容..."
    print(generate_msg)
    log_message(generate_msg, log_dir)
    
    email_html, subject = generate_email_html(table_data, current_date_str, total_count, report_type, recipients, cc)
    html_path, txt_path = save_email_contents(email_html, reminder_dir, report_type)
    
    save_msg = f"{report_type}类型邮件内容已保存到 {reminder_dir}:\nHTML: {html_path}\nTXT: {txt_path}"
    print(save_msg)
    log_message(save_msg, log_dir)
    
    # 发送Teams消息
    teams_message = f"{report_type} - 共{total_count}条请求需要处理。详情请查看邮件。"
    send_teams_message(teams_message, report_type)
    
    # 获取当前脚本所在目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"当前脚本目录: {current_dir}")
    print("当前目录下的内容：")
    for item in os.listdir(current_dir):
        item_path = os.path.join(current_dir, item)
        if os.path.isdir(item_path):
            print(f"[目录] {item}")
        else:
            print(f"[文件] {item}")

    
    # 添加当前目录到搜索路径
    sys.path.append(current_dir)
    
    try:
        # 尝试导入邮件发送模块
        from email_sender import send_email
    except ImportError:
        # 尝试绝对路径（请替换成你的实际脚本目录）
        absolute_script_dir = r"C:\Users\liang.wq.1\Downloads\ITC_Scorecard"  # 仅路径，不包含文件名
        print(f"当前目录未找到email_sender，尝试绝对路径: {absolute_script_dir}")
        
        # 添加绝对路径到搜索路径
        sys.path.append(absolute_script_dir)
        
        try:
            from email_sender import send_email
            print(f"已从绝对路径成功导入email_sender")
        except ImportError:
            import_err_msg = f"当前目录和绝对路径({absolute_script_dir})均未找到email_sender.py，请检查文件位置"
            print(import_err_msg)
            log_message(import_err_msg, log_dir)
            return
    
    # 发送邮件
    try:
        send_msg = f"正在发送{report_type}类型邮件... 标题: {subject}, 收件人: {', '.join(recipients)}, 抄送: {', '.join(cc)}"
        print(send_msg)
        log_message(send_msg, log_dir)
        
        # 读取邮件发送配置
        try:
            # 尝试从主程序导入邮件配置函数
            main_script_path = os.path.join(current_dir, 'BatRun_ITCreport_downloader_rev1.py')
            if os.path.exists(main_script_path):
                # 动态导入主程序的配置函数
                import importlib.util
                spec = importlib.util.spec_from_file_location("main_script", main_script_path)
                if spec is not None and spec.loader is not None:
                    main_module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(main_module)
                    email_enabled, auto_send = main_module.get_email_settings()
                else:
                    raise ImportError("无法加载主程序模块")
            else:
                # 备用方案：从配置文件读取
                from email_sender import load_email_config
                email_config = load_email_config()
                email_enabled = email_config.get('system_config', {}).get('EMAIL_ENABLED', True)
                auto_send = email_config.get('system_config', {}).get('AUTO_SEND_EMAIL', False)
        except Exception:
            email_enabled = True
            auto_send = False  # 默认为预览模式
        
        if not email_enabled:
            print(f"{report_type}类型邮件功能已禁用")
            log_message(f"{report_type}类型邮件功能已禁用", log_dir)
            return
        
        success = send_email(subject, email_html, to_addrs=recipients, cc_addrs=cc, auto_send=auto_send)
        
        if auto_send:
            success_msg = f"{report_type}类型邮件已直接发送" if success else f"{report_type}类型邮件发送失败"
        else:
            success_msg = f"{report_type}类型邮件预览窗口已打开，请确认后发送" if success else f"{report_type}类型邮件发送失败"
        
        print(success_msg)
        log_message(success_msg, log_dir)
    except Exception as e:
        send_err_msg = f"发送{report_type}类型邮件时出错: {str(e)}"
        print(send_err_msg)
        log_message(send_err_msg, log_dir)


# 主函数
def main():
    current_script_path = os.path.abspath(__file__)
    current_dir = os.path.dirname(current_script_path)
    
    itc_report_dir = os.path.abspath(os.path.join(current_dir, CONFIG["ITC_REPORT_DIR_NAME"]))
    raw_data_dir = os.path.join(itc_report_dir, CONFIG["RAW_DATA_DIR_NAME"])
    reminder_dir = os.path.join(itc_report_dir, CONFIG["REMINDER_DIR_NAME"])
    log_dir = os.path.join(itc_report_dir, CONFIG["LOG_DIR_NAME"])
    
    print("="*60)
    print("当前配置:")
    print(f"邮件发送阈值: 剩余天数 ≤ {CONFIG['MAX_REMAINING_DAYS_FOR_REPORT']}天")
    print(f"紧急程度划分: {CONFIG['URGENCY_LEVELS']}")
    print(f"ExitForm状态说明: {CONFIG['EMAIL_ExitForm_REVOKED']}")
    print(f"RoleChange状态说明: {CONFIG['EMAIL_RoleChange_REVOKED']}")
    print(f"数据目录: {raw_data_dir}")
    print(f"报告目录: {reminder_dir}")
    print(f"已配置的报表类型: {', '.join(CONFIG['reports'].keys())}")
    print("="*60)
    
    ensure_directory_exists(raw_data_dir)
    ensure_directory_exists(reminder_dir)
    ensure_directory_exists(log_dir)
    
    start_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_message(f"开始处理ITC报告数据 - {start_time}", log_dir)
    
    csv_files = []
    for f in os.listdir(raw_data_dir):
        if f.lower().endswith('.csv'):
            file_path = os.path.join(raw_data_dir, f)
            modified_time = os.path.getmtime(file_path)
            csv_files.append((file_path, modified_time))
    
    if not csv_files:
        error_msg = f"错误：在 {raw_data_dir} 中未找到任何CSV文件"
        print(error_msg)
        log_message(error_msg, log_dir)
        return
    
    csv_files.sort(key=lambda x: x[1], reverse=True)
    latest_csv_path = csv_files[0][0]
    latest_csv_time = datetime.fromtimestamp(csv_files[0][1]).strftime('%Y-%m-%d %H:%M:%S')
    
    info_msg = f"使用最新CSV文件: {latest_csv_path} (修改时间: {latest_csv_time})"
    print(info_msg)
    log_message(info_msg, log_dir)
    
    process_msg = "正在读取和处理数据..."
    print(process_msg)
    log_message(process_msg, log_dir)
    
    try:
        df = load_and_process_data(latest_csv_path)
        
        # 输出所有Category的分布情况（辅助调试）
        if 'Category' in df.columns:
            print("\n===== 调试：所有Category分布 =====")
            print(df['Category'].value_counts().to_string())
            print("===================================\n")
        
        # 注释：更新分析说明（移除EVP相关）
        analyze_msg = "正在分析请求数据（包括Pending review和Revoked）..."
        print(analyze_msg)
        log_message(analyze_msg, log_dir)
        
        results = analyze_requests(df)
        
        # 发送各类报表（仅Pending和Revoked，EVP已移除）
        for report_type, report_data in results.items():
            send_report(report_data, reminder_dir, log_dir)
            
    except Exception as e:
        error_msg = f"处理数据时发生错误: {str(e)}"
        print(error_msg)
        log_message(error_msg, log_dir)
        return
    
    end_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    complete_msg = f"ITC报告数据处理完成 - {end_time}"
    log_message(complete_msg, log_dir)
    print(complete_msg)


if __name__ == "__main__":
    main()