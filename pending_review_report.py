# -*- coding: utf-8 -*-
"""
Pending / Revoked 报表处理、邮件生成与 Teams 通知
包含:
- CSV 读取(自动编码尝试)
- Pending/Revoked 分析与聚合
- 高视觉邮件模板(徽章/统计卡片)
- Teams 发送(卡片 + fallback 简单文本)
- JSON 结果输出 (处理 numpy / datetime)
"""

import os, sys, json, re, argparse, traceback, requests
import pandas as pd
import numpy as np
from datetime import datetime, date
import faulthandler
faulthandler.enable()
os.environ.setdefault("PANDAS_ARROW_DISABLED", "1")

DEFAULT_CONFIG = {
    "reports": {
        "Pending review任务提醒": {
            "recipients": [],
            "cc": [],
            "cc1": {},
            "system_config": {
                "MAX_REMAINING_DAYS_FOR_REPORT": 10,
                "URGENCY_LEVELS": {"非常紧急": 2, "紧急": 4, "常规": 10},
                "ITC_REPORT_DIR_NAME": "ITC report",
                "RAW_DATA_DIR_NAME": "RawData",
                "REMINDER_DIR_NAME": "Reminder",
                "LOG_DIR_NAME": "Log",
                "EMAIL_SUBJECT_PENDING": "Pending review任务提醒",
                "EMAIL_SUBJECT_REVOKED": "Revoked状态任务提醒",
                "EMAIL_ExitForm_REVOKED": "ExitForm:SSO的应用/加入域的系统或者没有Onekey系统权限就无法登录系统的，可以在1年内在系统里面移除并确认，否则24小时移除；换句话说，Onekey user的权限一定要求离职通知的24小时内移除",
                "EMAIL_RoleChange_REVOKED": "请在30天内移除并在ITC确认",
                "ITC_SYSTEM_LINK": "https://itc-tool.pg.com/ComplianceReport?siteId=193"
            }
        },
        "Revoked状态任务提醒": {"recipients": [], "cc": [], "cc1": {}}
    },
    "Teams": {"webhook_url": ""}
}

def load_config(path=None):
    if not path:
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_config.json")
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=4)
        return DEFAULT_CONFIG.copy()
    try:
        with open(path, "r", encoding="utf-8") as f:
            user = json.load(f)
        cfg = DEFAULT_CONFIG.copy()
        if "reports" in user:
            for k, v in user["reports"].items():
                if k not in cfg["reports"]:
                    cfg["reports"][k] = {
                        "recipients": [], "cc": [], "cc1": {},
                        "system_config": DEFAULT_CONFIG["reports"]["Pending review任务提醒"]["system_config"].copy()
                    }
                for field in ["recipients", "cc", "cc1", "system_config"]:
                    if field in v:
                        cfg["reports"][k][field] = v[field]
        if "Teams" in user:
            cfg["Teams"] = user["Teams"]
        base_levels = DEFAULT_CONFIG["reports"]["Pending review任务提醒"]["system_config"]["URGENCY_LEVELS"]
        for lvl, val in base_levels.items():
            if lvl not in cfg["reports"]["Pending review任务提醒"]["system_config"]["URGENCY_LEVELS"]:
                cfg["reports"]["Pending review任务提醒"]["system_config"]["URGENCY_LEVELS"][lvl] = val
        return cfg
    except Exception:
        return DEFAULT_CONFIG.copy()

CONFIG = load_config()
_SYS = CONFIG["reports"]["Pending review任务提醒"]["system_config"]

def get_cfg(key):
    return CONFIG["reports"]["Pending review任务提醒"]["system_config"].get(
        key,
        DEFAULT_CONFIG["reports"]["Pending review任务提醒"]["system_config"].get(key)
    )

def cfg_values():
    return {
        "MAX_REMAINING_DAYS_FOR_REPORT": get_cfg("MAX_REMAINING_DAYS_FOR_REPORT"),
        "URGENCY_LEVELS": get_cfg("URGENCY_LEVELS"),
        "EMAIL_SUBJECT_PENDING": get_cfg("EMAIL_SUBJECT_PENDING"),
        "EMAIL_SUBJECT_REVOKED": get_cfg("EMAIL_SUBJECT_REVOKED"),
        "ITC_SYSTEM_LINK": get_cfg("ITC_SYSTEM_LINK"),
        "EMAIL_ExitForm_REVOKED": get_cfg("EMAIL_ExitForm_REVOKED"),
        "EMAIL_RoleChange_REVOKED": get_cfg("EMAIL_RoleChange_REVOKED"),
    }

def ensure_directory_exists(p):
    if not os.path.exists(p):
        os.makedirs(p, exist_ok=True)
    return p

def log_message(msg, log_dir):
    ensure_directory_exists(log_dir)
    lf = os.path.join(log_dir, f"process_{datetime.now().strftime('%Y%m%d')}.log")
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
        with open(lf, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def ensure_pg_email(email, username=None):
    if pd.isna(email) or str(email).strip() == "":
        if username:
            return f"{username.strip().lower().replace(' ', '.')}@pg.com"
        return ""
    email = str(email).strip()
    if '@' not in email:
        return f"{email.lower()}@pg.com"
    if not email.lower().endswith("@pg.com"):
        return email.split("@")[0] + "@pg.com"
    return email

def make_json_safe(obj):
    import numpy as _np
    from datetime import datetime as _dt, date as _date
    if isinstance(obj, (_np.integer,)): return int(obj)
    if isinstance(obj, (_np.floating,)): return float(obj)
    if isinstance(obj, (_dt, _date)): return obj.isoformat()
    if isinstance(obj, dict): return {k: make_json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)): return [make_json_safe(v) for v in obj]
    return obj

def fix_mojibake(text):
    if not isinstance(text, str): return text
    if any(x in text for x in ["â€“", "Ã", "å", "æ", "é"]):
        for enc in ["latin1", "cp1252"]:
            try:
                return text.encode(enc).decode("utf-8")
            except Exception:
                pass
    return text

def apply_mojibake_fix(df):
    if df is None or df.empty:
        return df
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].apply(fix_mojibake)
    return df

def detect_file_encoding(path, sample_size=4096):
    cands = ["utf-8-sig", "utf-8", "gbk", "cp936", "latin1"]
    with open(path, "rb") as f:
        raw = f.read(sample_size)
    for enc in cands:
        try:
            raw.decode(enc)
            return enc
        except Exception:
            pass
    return "latin1"

def load_and_process_data(csv_file_path):
    base_log_dir = os.path.join(os.path.dirname(csv_file_path), "..")
    log_message(f"开始读取CSV: {csv_file_path}", base_log_dir)
    enc_guess = detect_file_encoding(csv_file_path)
    log_message(f"初步编码猜测: {enc_guess}", base_log_dir)
    df = None
    tried = []
    for enc in [enc_guess, "utf-8-sig", "utf-8", "gbk", "cp936", "latin1"]:
        if enc in tried: continue
        tried.append(enc)
        try:
            df = pd.read_csv(csv_file_path, encoding=enc, dtype=str,
                             na_values=["", " ", "NA"], keep_default_na=True,
                             on_bad_lines="skip")
            log_message(f"使用编码 {enc} 读取成功。", base_log_dir)
            break
        except Exception as e:
            log_message(f"编码 {enc} 失败: {e}", base_log_dir)
    if df is None:
        raise RuntimeError("无法读取CSV。")
    # 禁用 pandas future warning
    pd.set_option('future.no_silent_downcasting', True)
    df = df.replace(r'^\s*$', np.nan, regex=True)
    try:
        df = df.infer_objects(copy=False)
    except Exception:
        pass
    df = apply_mojibake_fix(df)
    for col in ["Requested Date", "Expiration Date", "Log Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    df["is_new_request"] = df["Requester"].notna().astype(int) if "Requester" in df.columns else 0
    df["request_group"] = df["is_new_request"].cumsum()
    fill_cols = [
        "Requester","Requester Email","Request For","Request For Email","Requested Date","Area","Category","Category Description",
        "System/Solution","System/Solution Description","Approval Text","Owner Guidelines","Expiration Date","Max Request Age (Days)",
        "Access Type","Temporary Access?","Privileged?","Status","Confirmed?","Reason","Remark/Role","Employee Status",
        "Log Actor","Log Status","Log Date","Request ID"
    ]
    for col in fill_cols:
        if col in df.columns:
            df[col] = df.groupby("request_group")[col].transform(lambda x: x.ffill().bfill())
    log_message(f"读取完成: 行数={len(df)} 列数={len(df.columns)}", base_log_dir)
    return df

SITE_COLUMNS_PRIORITY = ["Site", "Site ID", "SiteID", "Site_Id", "Area", "Category"]

def extract_site_tokens(row):
    for col in SITE_COLUMNS_PRIORITY:
        if col in row.index and pd.notna(row[col]) and str(row[col]).strip():
            raw = str(row[col]).strip()
            parts = re.split(r"[,\s;/]+", raw)
            toks = [p.strip().upper() for p in parts if p.strip()]
            if toks:
                return toks
    return []

def match_cc1_emails_by_sites(site_tokens, report_type):
    cc1_cfg = CONFIG["reports"].get(report_type, {}).get("cc1", {})
    found = set()
    for token in site_tokens:
        for key, emails in cc1_cfg.items():
            kn = key.strip().upper()
            if token == kn or token in kn or kn in token:
                for e in emails:
                    found.add(ensure_pg_email(e))
    return sorted(found)

def analyze_requests(df):
    log_message("开始分析数据", os.getcwd())
    for col in ["Status", "System/Solution", "Request For", "Category"]:
        if col not in df.columns:
            raise ValueError(f"缺少列: {col}")
    pending_df = df[(df["Status"] == "Pending Review") & df["System/Solution"].notna() & df["Request For"].notna() & df["Category"].notna()].copy()
    revoked_df = df[(df["Status"].str.contains("Revoked", case=False, na=False)) & df["System/Solution"].notna() & df["Request For"].notna() & df["Category"].notna()].copy()
    log_message(f"Pending Review 行: {len(pending_df)} Revoked 行: {len(revoked_df)}", os.getcwd())
    today = date.today()
    return {
        "pending": process_pending_requests(pending_df, today),
        "revoked": process_revoked_requests(revoked_df, today)
    }

def process_pending_requests(pending_df, current_date):
    rpt = "Pending review任务提醒"
    cv = cfg_values()
    max_days = cv["MAX_REMAINING_DAYS_FOR_REPORT"]
    urgency = cv["URGENCY_LEVELS"]
    if pending_df.empty:
        empty = pd.DataFrame(columns=["Action Owner","Action Owner Email","System Name","Category","剩余天数","紧急程度","Pending_review数量"])
        return {"table": empty, "total_count": 0, "recipients": CONFIG["reports"][rpt].get("recipients", []), "cc": CONFIG["reports"][rpt].get("cc", []), "type": rpt, "items": []}
    pending_df["expiration_date_only"] = pending_df["Expiration Date"].dt.date if "Expiration Date" in pending_df.columns else pd.NaT
    def remaining(row):
        if pd.notna(row["expiration_date_only"]):
            return max((row["expiration_date_only"] - current_date).days, 0)
        return max_days
    pending_df["剩余天数"] = pending_df.apply(remaining, axis=1).astype(int)
    filtered = pending_df[pending_df["剩余天数"] <= max_days].copy()
    if filtered.empty:
        empty = pd.DataFrame(columns=["Action Owner","Action Owner Email","System Name","Category","剩余天数","紧急程度","Pending_review数量"])
        return {"table": empty, "total_count": 0, "recipients": CONFIG["reports"][rpt].get("recipients", []), "cc": CONFIG["reports"][rpt].get("cc", []), "type": rpt, "items": []}
    def mark(d):
        if d <= urgency["非常紧急"]: return "非常紧急"
        if d <= urgency["紧急"]: return "紧急"
        return "常规"
    filtered["紧急程度"] = filtered["剩余天数"].apply(mark)
    def owner_info(g):
        approvals = g[g["Log Status"].isin(["Approved","PartiallyApproved"])] if "Log Status" in g.columns else pd.DataFrame()
        if not approvals.empty and "Log Date" in approvals.columns:
            latest = approvals.loc[approvals["Log Date"].idxmax()]
            em = ensure_pg_email(latest.get("Log Actor Email",""), latest.get("Log Actor",""))
            return latest.get("Log Actor","未知"), em
        rq = g["Requester"].iloc[0] if "Requester" in g.columns and pd.notna(g["Requester"].iloc[0]) else "未知"
        rq_email = g["Requester Email"].iloc[0] if "Requester Email" in g.columns else ""
        return rq, ensure_pg_email(rq_email, rq)
    rows = []
    for _, grp in filtered.groupby("request_group"):
        ao, ao_email = owner_info(grp)
        first = grp.iloc[0]
        site_tokens = extract_site_tokens(first)
        rows.append({
            "Action Owner": ao,
            "Action Owner Email": ao_email,
            "System/Solution": first.get("System/Solution"),
            "Category": first.get("Category"),
            "SiteTokens": site_tokens,
            "剩余天数": int(first.get("剩余天数", 0)),
            "紧急程度": first.get("紧急程度"),
            "Request ID": first.get("Request ID", "N/A")
        })
    df_owner = pd.DataFrame(rows)
    agg = df_owner.groupby(["Action Owner","Action Owner Email","System/Solution","Category","剩余天数","紧急程度"]).size().reset_index(name="Pending_review数量")
    agg.rename(columns={"System/Solution": "System Name"}, inplace=True)
    total = int(agg["Pending_review数量"].sum())
    agg = pd.concat([agg, pd.DataFrame([{"Action Owner": "总计","Action Owner Email":"","System Name":"","Category":"","剩余天数":"","紧急程度":"","Pending_review数量": total}])], ignore_index=True)
    config_rec = CONFIG["reports"][rpt].get("recipients", [])
    config_cc = CONFIG["reports"][rpt].get("cc", [])
    data_rec = df_owner["Action Owner Email"].dropna().unique().tolist()
    all_site_tokens = set(t for toks in df_owner["SiteTokens"] for t in toks)
    cc1_emails = match_cc1_emails_by_sites(all_site_tokens, rpt)
    recipients = sorted(list(set([ensure_pg_email(e) for e in config_rec + data_rec if e])))
    cc_all = sorted(list(set([ensure_pg_email(e) for e in config_cc + cc1_emails if e])))
    cc_all = [e for e in cc_all if e not in recipients]
    return {"table": agg, "total_count": total, "recipients": recipients, "cc": cc_all, "type": rpt, "items": rows}

def process_revoked_requests(revoked_df, current_date):
    rpt = "Revoked状态任务提醒"
    cv = cfg_values()
    exit_note = cv["EMAIL_ExitForm_REVOKED"]
    role_note = cv["EMAIL_RoleChange_REVOKED"]
    if revoked_df.empty:
        empty = pd.DataFrame(columns=["Action Owner","Action Owner Email","System Name","Category","状态","状态说明","Revoked数量"])
        return {"table": empty, "total_count": 0, "recipients": CONFIG["reports"][rpt].get("recipients", []), "cc": CONFIG["reports"][rpt].get("cc", []), "type": rpt, "items": []}
    def status_note(st):
        if pd.isna(st): return ""
        s = str(st).lower()
        if "exitform" in s: return exit_note
        if "rolechange" in s: return role_note
        return ""
    revoked_df["状态说明"] = revoked_df["Status"].apply(status_note)
    def owner(g):
        confirmed = g[g["Log Status"].str.contains("confirmed", case=False, na=False)] if "Log Status" in g.columns else pd.DataFrame()
        if not confirmed.empty and "Log Date" in confirmed.columns:
            latest = confirmed.loc[confirmed["Log Date"].idxmax()]
            em = ensure_pg_email(latest.get("Log Actor Email",""), latest.get("Log Actor",""))
            return latest.get("Log Actor","未知"), em
        if "Log Date" in g.columns and not g["Log Date"].isna().all():
            latest = g.loc[g["Log Date"].idxmax()]
            em = ensure_pg_email(latest.get("Log Actor Email",""), latest.get("Log Actor",""))
            return latest.get("Log Actor","未知"), em
        rq = g["Requester"].iloc[0] if "Requester" in g.columns and pd.notna(g["Requester"].iloc[0]) else "未知"
        rq_email = g["Requester Email"].iloc[0] if "Requester Email" in g.columns else ""
        return rq, ensure_pg_email(rq_email, rq)
    rows = []
    for _, grp in revoked_df.groupby("request_group"):
        ao, ao_email = owner(grp)
        first = grp.iloc[0]
        site_tokens = extract_site_tokens(first)
        rows.append({
            "Action Owner": ao,
            "Action Owner Email": ao_email,
            "System/Solution": first.get("System/Solution"),
            "Category": first.get("Category"),
            "SiteTokens": site_tokens,
            "Status": first.get("Status"),
            "状态说明": first.get("状态说明"),
            "Request ID": first.get("Request ID", "N/A")
        })
    df_owner = pd.DataFrame(rows)
    agg = df_owner.groupby(["Action Owner","Action Owner Email","System/Solution","Category","Status","状态说明"]).size().reset_index(name="Revoked数量")
    agg.rename(columns={"System/Solution": "System Name"}, inplace=True)
    total = int(agg["Revoked数量"].sum())
    agg = pd.concat([agg, pd.DataFrame([{"Action Owner":"总计","Action Owner Email":"","System Name":"","Category":"","Status":"","状态说明":"","Revoked数量": total}])], ignore_index=True)
    config_rec = CONFIG["reports"][rpt].get("recipients", [])
    config_cc = CONFIG["reports"][rpt].get("cc", [])
    data_rec = df_owner["Action Owner Email"].dropna().unique().tolist()
    all_site_tokens = set(t for toks in df_owner["SiteTokens"] for t in toks)
    cc1_emails = match_cc1_emails_by_sites(all_site_tokens, rpt)
    recipients = sorted(list(set([ensure_pg_email(e) for e in config_rec + data_rec if e])))
    cc_all = sorted(list(set([ensure_pg_email(e) for e in config_cc + cc1_emails if e])))
    cc_all = [e for e in cc_all if e not in recipients]
    return {"table": agg, "total_count": total, "recipients": recipients, "cc": cc_all, "type": rpt, "items": rows}

def dataframe_to_markdown(df):
    if df.empty: return "_无数据_"
    headers = df.columns.tolist()
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join(["---"] * len(headers)) + " |"
    ]
    for _, r in df.iterrows():
        lines.append("| " + " | ".join("" if pd.isna(r[h]) else str(r[h]) for h in headers) + " |")
    return "\n".join(lines)

def build_teams_markdown(report_data, subject):
    cv = cfg_values()
    if report_data["type"] == "Pending review任务提醒":
        # 计算紧急程度统计
        df_body = report_data["table"][report_data["table"]["Action Owner"] != "总计"].copy()
        stats_extreme = int((df_body.get("紧急程度") == "非常紧急").sum()) if "紧急程度" in df_body.columns else 0
        stats_urgent = int((df_body.get("紧急程度") == "紧急").sum()) if "紧急程度" in df_body.columns else 0
        stats_normal = int((df_body.get("紧急程度") == "常规").sum()) if "紧急程度" in df_body.columns else 0
        
        # 构建优美的 markdown 格式（无完整表格，避免超时）
        urgency_lines = []
        if stats_extreme > 0:
            urgency_lines.append(f"🔴 **非常紧急**: {stats_extreme} 条 (需立即处理)")
        if stats_urgent > 0:
            urgency_lines.append(f"🟠 **紧急**: {stats_urgent} 条 (建议当天处理)")
        if stats_normal > 0:
            urgency_lines.append(f"🟢 **常规**: {stats_normal} 条 (请在过期前完成)")
        
        urgency_section = "\n".join(urgency_lines) if urgency_lines else "无紧急项"
        
        # 获取前5条明细摘要
        detail_lines = []
        for idx, (_, row) in enumerate(df_body.iterrows()):
            if idx >= 5:
                remaining = len(df_body) - 5
                detail_lines.append(f"... 还有 {remaining} 条")
                break
            ao = row.get("Action Owner", "")
            sys = row.get("System Name", "")
            cat = row.get("Category", "")
            urgency = row.get("紧急程度", "")
            qty = row.get("Pending_review数量", "")
            if urgency == "非常紧急":
                emoji = "🔴"
            elif urgency == "紧急":
                emoji = "🟠"
            else:
                emoji = "🟢"
            detail_lines.append(f"{emoji} {ao} | {sys} | {cat} ({qty}条)")
        
        details_section = "\n".join(detail_lines) if detail_lines else "无明细"
        
        intro = f"""### {subject}

**✅ 系统检测到当前有 {report_data['total_count']} 条待审核请求**

**紧急程度统计：**
{urgency_section}

**待审核摘要：**
{details_section}

---

**处理要求：**
为避免权限过期影响业务正常运行，请及时处理您负责的审核任务。

请通过 Chrome 浏览器（其他浏览器可能存在兼容问题）登录 [ITC 系统]({cv['ITC_SYSTEM_LINK']})，点击 **MyTasks / MyActions** 完成相关任务处理。
"""
        return intro
    else:
        # Revoked 消息
        detail_lines = []
        df_body = report_data["table"][report_data["table"]["Action Owner"] != "总计"].copy()
        for idx, (_, row) in enumerate(df_body.iterrows()):
            if idx >= 5:
                remaining = len(df_body) - 5
                detail_lines.append(f"... 还有 {remaining} 条")
                break
            ao = row.get("Action Owner", "")
            sys = row.get("System Name", "")
            status = row.get("Status", "")
            qty = row.get("Revoked数量", "")
            detail_lines.append(f"⚠️  {ao} | {sys} | {status} ({qty}条)")
        
        details_section = "\n".join(detail_lines) if detail_lines else "无明细"
        
        intro = f"""### {subject}

**⚠️ 当前 Revoked 总数：{report_data['total_count']}**

**Revoked 摘要：**
{details_section}

---

**处理要求：**
请核查状态说明并在系统中完成权限确认与清理。

[查看 ITC 系统]({cv['ITC_SYSTEM_LINK']})
"""
        return intro

def html_to_text(html):
    txt = re.sub(r"<style.*?</style>", "", html, flags=re.DOTALL)
    txt = re.sub(r"<script.*?</script>", "", txt, flags=re.DOTALL)
    txt = re.sub(r"<[^>]+>", " ", txt)
    return " ".join(txt.split())

def save_email_contents(html, out_dir, report_type):
    ensure_directory_exists(out_dir)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix = report_type.replace(" ", "_")
    hp = os.path.join(out_dir, f"{prefix}_email_{ts}.html")
    tp = os.path.join(out_dir, f"{prefix}_email_{ts}.txt")
    with open(hp, "w", encoding="utf-8") as f:
        f.write(html)
    with open(tp, "w", encoding="utf-8") as f:
        f.write(html_to_text(html))
    return hp, tp

def format_cn_date(s):
    try:
        dt = datetime.strptime(s, "%Y-%m-%d")
        return f"{dt.year}年{dt.month}月{dt.day}日"
    except Exception:
        return s

def _compute_pending_urgency_stats(table_data):
    df = table_data[table_data.get("Action Owner") != "总计"].copy()
    stats = {
        "非常紧急": int((df.get("紧急程度") == "非常紧急").sum()) if "紧急程度" in df.columns else 0,
        "紧急": int((df.get("紧急程度") == "紧急").sum()) if "紧急程度" in df.columns else 0,
        "常规": int((df.get("紧急程度") == "常规").sum()) if "紧急程度" in df.columns else 0
    }
    stats["总计"] = sum(stats.values())
    return stats
# ...existing code...

def generate_email_html(table_data, current_date, total_count, report_type, recipients, cc):
    cv = cfg_values()
    link = cv["ITC_SYSTEM_LINK"]
    cn_date = format_cn_date(current_date)
    to_str = ", ".join(recipients) if recipients else "无"
    cc_str = ", ".join(cc) if cc else "无"

    def fmt_days(v):
        if v is None or str(v).strip() == "":
            return ""
        try:
            return f"{int(v)}天"
        except Exception:
            return str(v)

    if report_type == "Pending review任务提醒":
        subject = f"{cv['EMAIL_SUBJECT_PENDING']} - {cn_date}"
        df_body = table_data[table_data["Action Owner"] != "总计"].copy()
        stats_extreme = int((df_body.get("紧急程度") == "非常紧急").sum()) if "紧急程度" in df_body.columns else 0
        stats_urgent = int((df_body.get("紧急程度") == "紧急").sum()) if "紧急程度" in df_body.columns else 0
        stats_normal = int((df_body.get("紧急程度") == "常规").sum()) if "紧急程度" in df_body.columns else 0
        rows_html = []
        for _, row in table_data.iterrows():
            if row.get("Action Owner") == "总计":
                rows_html.append(
                    "<tr class='total-row'>"
                    "<td>总计</td><td></td><td></td><td></td>"
                    "<td></td><td></td>"
                    f"<td>{row.get('Pending_review数量','')}</td></tr>"
                )
            else:
                urg = row.get("紧急程度", "")
                badge_class = {
                    "非常紧急": "badge-critical",
                    "紧急": "badge-warning",
                    "常规": "badge-normal"
                }.get(urg, "badge-normal")
                highlight_row = {
                    "非常紧急": "row-critical",
                    "紧急": "row-warning"
                }.get(urg, "")
                rows_html.append(
                    f"<tr class='{highlight_row}'>"
                    f"<td>{row.get('Action Owner','')}</td>"
                    f"<td>{row.get('Action Owner Email','')}</td>"
                    f"<td>{row.get('System Name','')}</td>"
                    f"<td>{row.get('Category','')}</td>"
                    f"<td>{fmt_days(row.get('剩余天数',''))}</td>"
                    f"<td><span class='badge {badge_class}'>{urg}</span></td>"
                    f"<td>{row.get('Pending_review数量','')}</td>"
                    "</tr>"
                )

        html = f"""<html><head><meta charset="UTF-8">
<style>
body {{
  font-family:"Segoe UI",Arial,sans-serif; background:#f5f7fa;
  margin:0; padding:24px; color:#1f2937; line-height:1.6; font-size:11pt;
}}
h2 {{ margin:0 0 16px; font-weight:600; color:#0f4c81; letter-spacing:.5px; font-size:14pt; }}
p {{ margin:0 0 12px; font-size:11pt; }}
.section-title {{
  font-weight:600; color:#0f4c81; margin:30px 0 8px; font-size:12.5pt;
  border-left:4px solid #0f4c81; padding-left:8px;
}}
.info-box {{
  background:#ffffff; border:1px solid #e2e8f0; border-radius:10px;
  padding:14px 18px; margin-bottom:18px; box-shadow:0 1px 3px rgba(0,0,0,0.05);
  font-size:11pt;
}}
.summary-cards {{ display:flex; gap:14px; flex-wrap:wrap; margin:18px 0 8px; }}
.card {{
  background:#fff; border:1px solid #e2e8f0; border-radius:10px;
  padding:10px 14px; min-width:140px; box-shadow:0 1px 2px rgba(0,0,0,0.05);
}}
.card h4 {{ margin:0 0 4px; font-size:11pt; font-weight:600; color:#475569; }}
.card .num {{ font-size:18pt; font-weight:600; }}
.card-critical .num {{ color:#c62828; }}
.card-warning .num {{ color:#ef6c00; }}
.card-normal .num {{ color:#2e7d32; }}

.urgency-desc p {{ margin:4px 0; }}
.line-critical {{ color:#c62828; font-weight:600; }}
.line-warning {{ color:#ef6c00; font-weight:600; }}
.line-normal {{ color:#2e7d32; font-weight:600; }}

table {{
  width:100%; border-collapse:separate; border-spacing:0;
  background:#fff; border:1px solid #d9e3ec; border-radius:12px;
  overflow:hidden; margin-top:6px;
}}
thead th {{
  background:linear-gradient(90deg,#0f4c81,#1769aa);
  color:#fff; padding:10px 12px; font-size:12.5px; letter-spacing:.6px;
  text-align:left;
}}
tbody td {{
  padding:8px 12px; font-size:12.5px; border-top:1px solid #eef2f7;
}}
tbody tr:nth-child(even) td {{ background:#f9fbfd; }}
.row-critical td {{ background:#fff5f5; }}
.row-warning td {{ background:#fff9ed; }}
.total-row td {{
  background:#eef6ff; font-weight:600; border-top:2px solid #c2d9f3;
}}

.badge {{
  display:inline-block; padding:4px 11px; border-radius:15px;
  font-size:11px; font-weight:600; letter-spacing:.5px;
}}
.badge-critical {{ background:#ffebee; color:#c62828; }}
.badge-warning {{ background:#fff3e0; color:#ef6c00; }}
.badge-normal {{ background:#e8f5e9; color:#2e7d32; }}

.footer {{
  margin-top:28px; font-size:11pt; color:#64748b;
  border-top:1px solid #e2e8f0; padding-top:14px;
}}
a {{ color:#0f4c81; text-decoration:none; }}
a:hover {{ text-decoration:underline; }}

.signature {{
  margin-top:20px; font-size:11.5pt; font-weight:500; color:#0f4c81;
}}
</style></head><body>
<h2>{subject}</h2>
<div class="info-box">
<p>系统检测到当前有 <strong>{total_count}</strong> 条待审核（Pending Review）请求（剩余天数 ≤ 10 天）。为避免权限过期影响业务正常运行，请及时处理您负责的审核任务。</p>
<p>请通过 <strong>Chrome 浏览器</strong>（其他浏览器可能存在兼容问题）登录 ITC 系统：<a href="{link}" target="_blank">{link}</a>，点击 <strong>MyTasks / MyActions</strong> 完成相关任务处理。</p>
<p><b>邮件接收人 (To):</b> {to_str}<br><b>抄送 (CC):</b> {cc_str}</p>
</div>

<div class="section-title">紧急程度说明</div>
<div class="info-box urgency-desc">
<p class="line-critical">非常紧急： 剩余天数 ≤ {cv['URGENCY_LEVELS']['非常紧急']} 天（需立即处理）</p>
<p class="line-warning">紧急： 剩余天数 ≤ {cv['URGENCY_LEVELS']['紧急']} 天（建议当天处理）</p>
<p class="line-normal">常规： 剩余天数 ≤ {cv['URGENCY_LEVELS']['常规']} 天（请在过期前完成）</p>
</div>

<div class="summary-cards">
  <div class="card card-critical"><h4>非常紧急</h4><div class="num">{stats_extreme}</div></div>
  <div class="card card-warning"><h4>紧急</h4><div class="num">{stats_urgent}</div></div>
  <div class="card card-normal"><h4>常规</h4><div class="num">{stats_normal}</div></div>
  <div class="card"><h4>总计</h4><div class="num">{total_count}</div></div>
</div>

<div class="section-title">待审核明细</div>
<table>
<thead><tr>
<th>负责人 (Action Owner)</th>
<th>负责人邮箱 (@pg.com)</th>
<th>系统名称 (System Name)</th>
<th>分类 (Category)</th>
<th>剩余天数</th>
<th>紧急程度</th>
<th>待审核数量</th>
</tr></thead>
<tbody>{''.join(rows_html)}</tbody>
</table>

<div class="footer">
感谢您的及时处理！如有任何问题，请联系 Site CSL 团队支持。
</div>
<div class="signature">此致<br>GC PD 网络安全团队</div>
</body></html>"""
        return html, subject

    subject = f"{cv['EMAIL_SUBJECT_REVOKED']} - {cn_date}"
    rows_html = []
    for _, row in table_data.iterrows():
        if row.get("Action Owner") == "总计":
            rows_html.append(
                "<tr class='total-row'>"
                "<td>总计</td><td></td><td></td><td></td><td></td><td></td>"
                f"<td>{row.get('Revoked数量','')}</td></tr>"
            )
        else:
            rows_html.append(
                "<tr>"
                f"<td>{row.get('Action Owner','')}</td>"
                f"<td>{row.get('Action Owner Email','')}</td>"
                f"<td>{row.get('System Name','')}</td>"
                f"<td>{row.get('Category','')}</td>"
                f"<td>{row.get('Status','')}</td>"
                f"<td>{row.get('状态说明','')}</td>"
                f"<td>{row.get('Revoked数量','')}</td>"
                "</tr>"
            )
    html = f"""<html><head><meta charset="UTF-8">
<style>
body{{font-family:"Segoe UI",Arial,sans-serif;background:#f6f8fa;padding:24px;color:#1f2937;line-height:1.6;font-size:11pt}}
h2{{margin:0 0 16px;font-weight:600;color:#0f4c81;letter-spacing:.5px;font-size:14pt}}
table{{width:100%;border-collapse:separate;border-spacing:0;background:#fff;border:1px solid #d9e3ec;border-radius:12px;overflow:hidden;margin-top:6px}}
thead th{{background:#374151;color:#fff;padding:10px 12px;font-size:11pt;text-align:left;letter-spacing:.6px}}
tbody td{{padding:8px 12px;font-size:11pt;border-top:1px solid #eef2f7}}
tbody tr:nth-child(even) td{{background:#f9fbfd}}
.total-row td{{background:#eef6ff;font-weight:600;border-top:2px solid #c2d9f3}}
.info-box{{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px 18px;margin-bottom:18px;box-shadow:0 1px 3px rgba(0,0,0,0.05);font-size:11pt}}
.footer{{margin-top:24px;font-size:11pt;color:#64748b;border-top:1px solid #e2e8f0;padding-top:14px}}
a{{color:#0f4c81;text-decoration:none}}a:hover{{text-decoration:underline}}
.signature{{margin-top:20px;font-size:11.5pt;font-weight:500;color:#0f4c81}}
</style></head><body>
<h2>{subject}</h2>
<div class="info-box">
<p>当前 Revoked 总数 <strong>{total_count}</strong>。请核查状态说明并在系统中完成权限确认与清理。</p>
<p>系统链接：<a href="{link}" target="_blank">{link}</a></p>
<p><b>邮件接收人 (To):</b> {to_str}<br><b>抄送 (CC):</b> {cc_str}</p>
</div>
<table>
<thead><tr>
<th>负责人 (Action Owner)</th>
<th>负责人邮箱 (@pg.com)</th>
<th>系统名称 (System Name)</th>
<th>分类 (Category)</th>
<th>状态 (Status)</th>
<th>状态说明</th>
<th>数量</th>
</tr></thead>
<tbody>{''.join(rows_html)}</tbody>
</table>
<div class="footer">提示: 请根据状态说明及时处理（ExitForm / RoleChange）。</div>
<div class="signature">此致<br>GC PD 网络安全团队</div>
</body></html>"""
    return html, subject

def send_to_teams_simple_markdown(subject, markdown_content, log_dir):
    try:
        tc_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "teams_config.json")
        url = ""
        if os.path.exists(tc_path):
            with open(tc_path, "r", encoding="utf-8") as f:
                tcfg = json.load(f)
            if tcfg.get("enabled"):
                def_name = tcfg.get("default_webhook")
                url = tcfg.get("webhooks", {}).get(def_name, "")
        if not url:
            url = CONFIG.get("Teams", {}).get("webhook_url", "").strip()
    except Exception:
        url = CONFIG.get("Teams", {}).get("webhook_url", "").strip()
    if not url:
        log_message("Teams simple fallback 无URL", log_dir)
        return False
    payload = {"text": f"{subject}\n{markdown_content[:7000]}"}
    try:
        r = requests.post(url, json=payload, timeout=25)
        log_message(f"SimpleWebhook HTTP {r.status_code}", log_dir)
        return 200 <= r.status_code < 300
    except Exception as e:
        log_message(f"Teams simple 异常: {e}", log_dir)
        return False

def send_health_probe(log_dir):
    url = CONFIG.get("Teams", {}).get("webhook_url", "").strip()
    if not url:
        log_message("健康检测: 无URL", log_dir)
        return
    try:
        r = requests.post(url, json={"text": "健康探测"}, timeout=10)
        log_message(f"健康检测状态: {r.status_code}", log_dir)
    except Exception as e:
        log_message(f"健康检测异常: {e}", log_dir)

SCRIPT_VERSION = "pending_report_v9_trace"

def _trace_banner():
    print(f"[TRACE] pending_review_report loaded VERSION={SCRIPT_VERSION} FILE={__file__}")

_trace_banner()

def send_report(report_data, reminder_dir, log_dir):
    log_message(f"[DEBUG] send_report() 被调用, type={report_data.get('type')}, total_count={report_data.get('total_count')}", log_dir)
    log_message(f"[VER {SCRIPT_VERSION}] 开始发送报告: {report_data['type']}", log_dir)
    if report_data["total_count"] == 0:
        log_message(f"[VER {SCRIPT_VERSION}] {report_data['type']} 无数据跳过", log_dir)
        return
    now_str = datetime.now().strftime("%Y-%m-%d")
    email_html, subject = generate_email_html(report_data["table"], now_str, report_data["total_count"],
                                              report_data["type"], report_data["recipients"], report_data["cc"])
    log_message(f"[VER {SCRIPT_VERSION}] 邮件HTML生成 subject={subject}", log_dir)
    html_path, _ = save_email_contents(email_html, reminder_dir, report_data["type"])
    log_message(f"[VER {SCRIPT_VERSION}] 保存邮件文件: {html_path}", log_dir)

    send_email_func = None
    email_enabled = True
    try:
        from email_sender import send_email as _send_email, load_email_config
        ecfg = load_email_config()
        email_enabled = ecfg.get("system_config", {}).get("EMAIL_ENABLED", True)
        send_email_func = _send_email
        log_message(f"[VER {SCRIPT_VERSION}] 邮件模块加载完成 ENABLED={email_enabled}", log_dir)
    except Exception as e:
        log_message(f"[VER {SCRIPT_VERSION}] 邮件模块加载失败: {e}", log_dir)

    if email_enabled and send_email_func and (report_data["recipients"] or report_data["cc"]):
        try:
            log_message(f"[VER {SCRIPT_VERSION}] 邮件发送开始...", log_dir)
            ok = send_email_func(subject, email_html,
                                 to_addrs=report_data["recipients"], cc_addrs=report_data["cc"])
            log_message(f"[VER {SCRIPT_VERSION}] 邮件发送结果={ok}", log_dir)
        except Exception as e:
            log_message(f"[VER {SCRIPT_VERSION}] 邮件发送异常: {e}", log_dir)
            log_message(traceback.format_exc(), log_dir)
    else:
        log_message(f"[VER {SCRIPT_VERSION}] 邮件阶段跳过 ENABLED={email_enabled} to={len(report_data['recipients']) if report_data else 0} cc={len(report_data['cc']) if report_data else 0}", log_dir)

    log_message(f"[VER {SCRIPT_VERSION}] 准备进入Teams阶段", log_dir)
    md = build_teams_markdown(report_data, subject)
    urgent_flag = False
    rule_key = "normal_issues"
    if report_data["type"] == "Pending review任务提醒":
        urgent_flag = any(it.get("紧急程度") == "非常紧急" for it in report_data["items"])
        rule_key = "urgent_issues" if urgent_flag else "normal_issues"
    elif report_data["type"] == "Revoked状态任务提醒":
        rule_key = "revoked_issues"
    log_message(f"[VER {SCRIPT_VERSION}] 规则判定 rule_key={rule_key} urgent={urgent_flag}", log_dir)

    teams_success = False
    try:
        log_message(f"[VER {SCRIPT_VERSION}] 即将导入 teams_sender", log_dir)
        if "teams_sender" in sys.modules:
            del sys.modules["teams_sender"]
        import teams_sender
        log_message(f"[VER {SCRIPT_VERSION}] 已导入 teams_sender 路径={getattr(teams_sender,'__file__','?')}", log_dir)
        tc = teams_sender.load_teams_config()
        log_message(f"[VER {SCRIPT_VERSION}] Teams配置 enabled={tc.get('enabled')} default={tc.get('default_webhook')} webhooks={list(tc.get('webhooks',{}).keys())}", log_dir)
        rules = tc.get("notification_rules", {})
        webhook_name = rules.get(rule_key, {}).get("webhook", tc.get("default_webhook",""))
        log_message(f"[VER {SCRIPT_VERSION}] 选定 webhook_name={webhook_name}", log_dir)
        if not tc.get("enabled"):
            log_message(f"[VER {SCRIPT_VERSION}] Teams未启用跳过", log_dir)
        else:
            ok, msg = teams_sender.send_teams_message(subject, md, webhook_name or tc.get("default_webhook",""),
                                                      urgent_flag, teams_config=tc)
            log_message(f"[VER {SCRIPT_VERSION}] 卡片发送结果 ok={ok} msg={msg}", log_dir)
            teams_success = ok
    except Exception as e:
        log_message(f"[VER {SCRIPT_VERSION}] Teams发送异常: {e}", log_dir)
        log_message(traceback.format_exc(), log_dir)

    if not teams_success:
        log_message(f"[VER {SCRIPT_VERSION}] 尝试 fallback simple", log_dir)
        fb = send_to_teams_simple_markdown(subject, md, log_dir)
        log_message(f"[VER {SCRIPT_VERSION}] fallback结果={fb}", log_dir)
        teams_success = teams_success or fb

    log_message(f"[VER {SCRIPT_VERSION}] 最终Teams状态={teams_success}", log_dir)

def main(selected_csv_path=None):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    itc_dir = os.path.join(base_dir, get_cfg("ITC_REPORT_DIR_NAME"))
    raw_dir = os.path.join(itc_dir, get_cfg("RAW_DATA_DIR_NAME"))
    reminder_dir = os.path.join(itc_dir, get_cfg("REMINDER_DIR_NAME"))
    log_dir = os.path.join(itc_dir, get_cfg("LOG_DIR_NAME"))
    for p in [raw_dir, reminder_dir, log_dir]: ensure_directory_exists(p)
    log_message(f"目录初始化: ITC_DIR='{itc_dir}' RAW='RawData' REMINDER='Reminder' LOG='Log'", log_dir)
    log_message("开始处理报告", log_dir)

    if selected_csv_path and os.path.exists(selected_csv_path):
        csv_path = selected_csv_path
        log_message(f"使用指定CSV: {csv_path}", log_dir)
    else:
        csv_files = [(os.path.join(raw_dir, f), os.path.getmtime(os.path.join(raw_dir, f)))
                     for f in os.listdir(raw_dir) if f.lower().endswith(".csv")]
        if not csv_files:
            log_message("未找到CSV文件", log_dir)
            return 1
        csv_files.sort(key=lambda x: x[1], reverse=True)
        csv_path = csv_files[0][0]
        log_message(f"选取最新CSV: {csv_path}", log_dir)

    summary = {}
    try:
        log_message("读取数据开始", log_dir)
        df = load_and_process_data(csv_path)
        log_message("读取数据完成", log_dir)
        if "Category" in df.columns:
            cats = df["Category"].dropna().value_counts().to_dict()
            log_message(f"Category分布: {json.dumps(cats, ensure_ascii=False)}", log_dir)
        log_message("分析开始", log_dir)
        results = analyze_requests(df)
        log_message(f"分析完成 Pending={results['pending']['total_count']} Revoked={results['revoked']['total_count']}", log_dir)
        log_message(f"[DEBUG] 即将循环遍历结果 results.keys()={list(results.keys())}", log_dir)
        for rpt in results.values():
            log_message(f"[DEBUG] 循环中: rpt type={rpt.get('type')}, total={rpt.get('total_count')}", log_dir)
            send_report(rpt, reminder_dir, log_dir)
        summary = {
            "pending_count": int(results["pending"]["total_count"]),
            "revoked_count": int(results["revoked"]["total_count"]),
            "pending_review_items": results["pending"].get("items", []),
            "revoked_items": results["revoked"].get("items", [])
        }
    except Exception as e:
        log_message(f"处理异常: {e}", log_dir)
        log_message(traceback.format_exc(), log_dir)
        summary = {"error": str(e)}

    result_path = os.path.join(base_dir, "a_results.json")
    try:
        with open(result_path, "w", encoding="utf-8") as f:
            json.dump(make_json_safe(summary), f, ensure_ascii=False, indent=4)
        log_message(f"结果写入: {result_path}", log_dir)
    except Exception as e:
        log_message(f"结果写入失败: {e}", log_dir)

# ...existing code...
    if "error" not in summary:
        log_message(f"完成 Summary Pending={summary['pending_count']} Revoked={summary['revoked_count']}", log_dir)
        return 0
    else:
        log_message(f"完成但出错: {summary['error']}", log_dir)
        return 1

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--csv-path", default=None)
    args = parser.parse_args()
    code = main(args.csv_path)
    sys.exit(code)