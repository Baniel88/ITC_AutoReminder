# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
"""
诊断 Revoked 数据问题
检查数据处理流程中 revoked_count 和 revoked_categories 的来源
"""

import os
import sys
import pandas as pd
import json
from datetime import datetime

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 全局日志列表
diagnostic_logs = []

def log_diagnostic(message, level="INFO"):
    """记录诊断日志"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_entry = f"[{timestamp}] [{level}] {message}"
    diagnostic_logs.append(log_entry)
    print(log_entry)


def save_diagnostic_log():
    """保存诊断日志到文件"""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(script_dir, "ITC report", "Log")
        
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        log_filename = f"diagnostic_revoked_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path = os.path.join(log_dir, log_filename)
        
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(diagnostic_logs))
        
        log_diagnostic(f"✅ 诊断日志已保存到: {log_path}", "INFO")
        return log_path
        
    except Exception as e:
        log_diagnostic(f"❌ 保存诊断日志失败: {str(e)}", "ERROR")
        return None


def find_latest_rawdata_file():
    """查找最新的 RawData 文件"""
    log_diagnostic("=" * 80, "INFO")
    log_diagnostic("开始查找最新的 RawData 文件", "INFO")
    log_diagnostic("=" * 80, "INFO")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    raw_data_dir = os.path.join(script_dir, "ITC report", "RawData")
    
    log_diagnostic(f"RawData 目录: {raw_data_dir}", "INFO")
    
    if not os.path.exists(raw_data_dir):
        log_diagnostic(f"❌ RawData 目录不存在: {raw_data_dir}", "ERROR")
        return None
    
    files = []
    for root, _, filenames in os.walk(raw_data_dir):
        for filename in filenames:
            if filename.endswith(('.csv', '.xlsx', '.xls')) and not filename.startswith('.'):
                filepath = os.path.join(root, filename)
                mtime = os.path.getmtime(filepath)
                files.append((filepath, mtime))
                log_diagnostic(f"找到文件: {filename} (修改时间: {datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')})", "DEBUG")
    
    if not files:
        log_diagnostic(f"❌ RawData 目录中没有找到数据文件", "ERROR")
        return None
    
    # 返回最新的文件
    latest_file = max(files, key=lambda x: x[1])[0]
    log_diagnostic(f"✅ 找到最新文件: {os.path.basename(latest_file)}", "INFO")
    log_diagnostic(f"完整路径: {latest_file}", "DEBUG")
    log_diagnostic(f"修改时间: {datetime.fromtimestamp(os.path.getmtime(latest_file)).strftime('%Y-%m-%d %H:%M:%S')}", "DEBUG")
    
    return latest_file


def analyze_revoked_data(csv_path):
    """分析 CSV 文件中的 Revoked 数据"""
    log_diagnostic("=" * 80, "INFO")
    log_diagnostic(f"开始分析文件: {os.path.basename(csv_path)}", "INFO")
    log_diagnostic("=" * 80, "INFO")
    
    try:
        # 读取 CSV 文件
        log_diagnostic("正在读取文件...", "INFO")
        if csv_path.endswith('.csv'):
            df = pd.read_csv(csv_path, encoding='utf-8-sig')
        else:
            df = pd.read_excel(csv_path)
        
        log_diagnostic(f"✅ 成功读取文件", "INFO")
        log_diagnostic(f"📊 总记录数: {len(df)}", "INFO")
        log_diagnostic(f"📋 列数: {len(df.columns)}", "INFO")
        log_diagnostic(f"📋 列名列表: {list(df.columns)}", "DEBUG")
        
        # 检查是否有 Status 列
        status_columns = [col for col in df.columns if 'status' in col.lower() or 'state' in col.lower()]
        log_diagnostic(f"🔍 找到 {len(status_columns)} 个可能的状态列: {status_columns}", "INFO")
        
        # 检查 Revoked 数据
        total_revoked = 0
        revoked_details = []
        
        for col in status_columns:
            log_diagnostic(f"\n--- 分析列: '{col}' ---", "INFO")
            
            unique_values = df[col].unique()
            log_diagnostic(f"唯一值数量: {len(unique_values)}", "DEBUG")
            log_diagnostic(f"唯一值: {list(unique_values)}", "DEBUG")
            
            # 查找包含 'revoked' 的记录
            revoked_mask = df[col].astype(str).str.contains('revoked', case=False, na=False)
            revoked_count = revoked_mask.sum()
            
            log_diagnostic(f"包含 'revoked' 的记录数: {revoked_count}", "INFO")
            
            if revoked_count > 0:
                total_revoked += revoked_count
                log_diagnostic(f"⚠️ 在列 '{col}' 中找到 {revoked_count} 条 Revoked 记录！", "WARNING")
                
                # 显示 Revoked 记录的详细信息
                revoked_df = df[revoked_mask]
                
                # 查找 Category 列
                category_columns = [c for c in df.columns if 'category' in c.lower()]
                if category_columns:
                    category_col = category_columns[0]
                    revoked_categories = revoked_df[category_col].unique()
                    log_diagnostic(f"涉及的 Category ({len(revoked_categories)}个): {list(revoked_categories)}", "INFO")
                    
                    # 记录详细信息
                    for cat in revoked_categories:
                        cat_count = (revoked_df[category_col] == cat).sum()
                        revoked_details.append({
                            'category': cat,
                            'count': cat_count,
                            'status_column': col
                        })
                        log_diagnostic(f"  - {cat}: {cat_count} 条记录", "DEBUG")
                else:
                    log_diagnostic("⚠️ 未找到 Category 列", "WARNING")
                
                # 显示前几条记录的关键信息
                log_diagnostic(f"\n前 3 条 Revoked 记录的关键信息:", "DEBUG")
                key_columns = [c for c in df.columns if any(k in c.lower() for k in ['id', 'category', 'status', 'request', 'user'])]
                if key_columns:
                    sample_df = revoked_df[key_columns].head(3)
                    log_diagnostic(f"\n{sample_df.to_string()}", "DEBUG")
        
        if total_revoked == 0:
            log_diagnostic("\n✅ 文件中没有找到包含 'revoked' 的记录", "INFO")
            log_diagnostic("\n检查其他可能的状态值...", "INFO")
            
            # 显示所有状态列的唯一值统计
            for col in status_columns:
                log_diagnostic(f"\n列 '{col}' 的值分布:", "DEBUG")
                value_counts = df[col].value_counts()
                for val, count in value_counts.head(10).items():
                    log_diagnostic(f"  {val}: {count}", "DEBUG")
        else:
            log_diagnostic(f"\n📊 总计在文件中找到 {total_revoked} 条 Revoked 记录", "INFO")
            log_diagnostic(f"📋 详细分布: {json.dumps(revoked_details, ensure_ascii=False, indent=2)}", "DEBUG")
        
        return total_revoked, revoked_details
        
    except Exception as e:
        log_diagnostic(f"\n❌ 分析过程出错: {str(e)}", "ERROR")
        import traceback
        log_diagnostic(f"错误堆栈:\n{traceback.format_exc()}", "ERROR")
        return 0, []


def check_pending_review_script():
    """检查 pending_review_report.py 的处理逻辑"""
    log_diagnostic("\n" + "=" * 80, "INFO")
    log_diagnostic("检查 pending_review_report.py 的处理逻辑...", "INFO")
    log_diagnostic("=" * 80, "INFO")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    report_script = os.path.join(script_dir, "pending_review_report.py")
    
    if not os.path.exists(report_script):
        log_diagnostic(f"❌ 未找到报表处理脚本: {report_script}", "ERROR")
        return
    
    log_diagnostic(f"✅ 找到脚本: {os.path.basename(report_script)}", "INFO")
    
    try:
        with open(report_script, 'r', encoding='utf-8') as f:
            script_content = f.read()
        
        log_diagnostic(f"脚本文件大小: {len(script_content)} 字符", "DEBUG")
        
        # 查找 revoked 相关的代码
        import re
        
        log_diagnostic("\n🔍 搜索 revoked 相关代码...", "INFO")
        
        revoked_patterns = [
            (r"revoked_count\s*=.*", "revoked_count 赋值"),
            (r"revoked_categories\s*=.*", "revoked_categories 赋值"),
            (r"['\"]revoked['\"]", "字符串 'revoked'"),
            (r"Revoked", "关键词 Revoked"),
            (r"def.*revoked.*\(", "revoked 相关函数")
        ]
        
        for pattern, description in revoked_patterns:
            matches = re.findall(f".*{pattern}.*", script_content, re.IGNORECASE)
            if matches:
                log_diagnostic(f"\n找到 {description} ({len(matches)} 处):", "INFO")
                for i, match in enumerate(matches[:5], 1):  # 只显示前5个
                    log_diagnostic(f"  {i}. {match.strip()}", "DEBUG")
            else:
                log_diagnostic(f"未找到 {description}", "DEBUG")
        
        # 查找日志输出相关代码
        log_diagnostic("\n🔍 搜索日志输出相关代码...", "INFO")
        log_patterns = [
            (r"log_message.*revoked", "revoked 相关日志"),
            (r"print.*revoked", "revoked 相关打印")
        ]
        
        for pattern, description in log_patterns:
            matches = re.findall(f".*{pattern}.*", script_content, re.IGNORECASE)
            if matches:
                log_diagnostic(f"\n找到 {description} ({len(matches)} 处):", "INFO")
                for i, match in enumerate(matches[:5], 1):
                    log_diagnostic(f"  {i}. {match.strip()}", "DEBUG")
        
    except Exception as e:
        log_diagnostic(f"❌ 读取脚本文件失败: {str(e)}", "ERROR")


def check_latest_log():
    """检查最新的处理日志"""
    log_diagnostic("\n" + "=" * 80, "INFO")
    log_diagnostic("检查最新的处理日志...", "INFO")
    log_diagnostic("=" * 80, "INFO")
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "ITC report", "Log")
    
    if not os.path.exists(log_dir):
        log_diagnostic(f"❌ 日志目录不存在: {log_dir}", "ERROR")
        return
    
    log_diagnostic(f"日志目录: {log_dir}", "INFO")
    
    try:
        log_files = [f for f in os.listdir(log_dir) if f.endswith('.log') and not f.startswith('diagnostic')]
        if not log_files:
            log_diagnostic("❌ 没有找到日志文件", "WARNING")
            return
        
        log_diagnostic(f"找到 {len(log_files)} 个日志文件", "INFO")
        
        # 获取最新的日志文件
        latest_log = max([os.path.join(log_dir, f) for f in log_files], key=os.path.getmtime)
        log_diagnostic(f"✅ 最新日志文件: {os.path.basename(latest_log)}", "INFO")
        log_diagnostic(f"修改时间: {datetime.fromtimestamp(os.path.getmtime(latest_log)).strftime('%Y-%m-%d %H:%M:%S')}", "DEBUG")
        
        with open(latest_log, 'r', encoding='utf-8') as f:
            log_content = f.read()
        
        log_diagnostic(f"日志文件大小: {len(log_content)} 字符", "DEBUG")
        
        # 查找 revoked 相关的日志
        import re
        revoked_log_lines = re.findall(r".*[Rr]evoked.*", log_content)
        
        if revoked_log_lines:
            log_diagnostic(f"\n找到 {len(revoked_log_lines)} 条 revoked 相关日志:", "WARNING")
            log_diagnostic("\n最近 10 条 revoked 相关日志:", "INFO")
            for i, line in enumerate(revoked_log_lines[-10:], 1):
                log_diagnostic(f"  {i}. {line.strip()}", "DEBUG")
        else:
            log_diagnostic("\n✅ 日志中没有 revoked 相关记录", "INFO")
        
        # 查找数据统计相关的日志
        log_diagnostic("\n🔍 搜索数据统计相关日志...", "INFO")
        stat_patterns = [
            r"总记录数.*\d+",
            r"待审核.*\d+",
            r"排除.*\d+",
            r"发现.*\d+.*条"
        ]
        
        for pattern in stat_patterns:
            matches = re.findall(f".*{pattern}.*", log_content, re.IGNORECASE)
            if matches:
                log_diagnostic(f"\n找到匹配 '{pattern}':", "DEBUG")
                for match in matches[-3:]:  # 显示最后3条
                    log_diagnostic(f"  {match.strip()}", "DEBUG")
        
    except Exception as e:
        log_diagnostic(f"❌ 检查日志失败: {str(e)}", "ERROR")


def check_get_processing_summary():
    """检查 get_processing_summary 函数的实现"""
    log_diagnostic("\n" + "=" * 80, "INFO")
    log_diagnostic("检查 get_processing_summary() 函数...", "INFO")
    log_diagnostic("=" * 80, "INFO")
    
    try:
        from BatRun_ITCreport_downloader_rev1 import get_processing_summary
        
        log_diagnostic("✅ 成功导入 get_processing_summary 函数", "INFO")
        
        # 调用函数查看返回值
        log_diagnostic("\n正在调用 get_processing_summary()...", "INFO")
        summary = get_processing_summary()
        
        log_diagnostic(f"\n📊 get_processing_summary() 返回值:", "INFO")
        log_diagnostic(f"{json.dumps(summary, indent=2, ensure_ascii=False)}", "INFO")
        
        # 重点检查
        log_diagnostic(f"\n⚠️ 重点检查:", "INFO")
        log_diagnostic(f"  - total_records: {summary.get('total_records', 'N/A')}", "INFO")
        log_diagnostic(f"  - urgent_pending: {summary.get('urgent_pending', 'N/A')}", "INFO")
        log_diagnostic(f"  - normal_pending: {summary.get('normal_pending', 'N/A')}", "INFO")
        log_diagnostic(f"  - revoked_count: {summary.get('revoked_count', 'N/A')}", "INFO")
        log_diagnostic(f"  - revoked_categories: {summary.get('revoked_categories', 'N/A')}", "INFO")
        log_diagnostic(f"  - excluded_long_term: {summary.get('excluded_long_term', 'N/A')}", "INFO")
        
        revoked_count = summary.get('revoked_count', 0)
        if revoked_count > 0:
            log_diagnostic(f"\n⚠️ get_processing_summary() 返回了非零的 revoked_count: {revoked_count}!", "WARNING")
            log_diagnostic(f"这可能是因为:", "WARNING")
            log_diagnostic(f"  1. 从旧的日志文件中读取了过期数据", "WARNING")
            log_diagnostic(f"  2. 日志解析逻辑有问题", "WARNING")
            log_diagnostic(f"  3. pending_review_report.py 生成了错误的日志", "WARNING")
            
            revoked_categories = summary.get('revoked_categories', [])
            if revoked_categories:
                log_diagnostic(f"\nrevoked_categories 内容: {revoked_categories}", "WARNING")
        else:
            log_diagnostic(f"\n✅ revoked_count 为 0，这是正确的（如果实际数据中没有 revoked 记录）", "INFO")
        
        return summary
        
    except ImportError as e:
        log_diagnostic(f"\n❌ 无法导入 get_processing_summary: {str(e)}", "ERROR")
        return None
    except Exception as e:
        log_diagnostic(f"\n❌ 检查失败: {str(e)}", "ERROR")
        import traceback
        log_diagnostic(f"错误堆栈:\n{traceback.format_exc()}", "ERROR")
        return None


def generate_diagnostic_report(rawdata_revoked_count, rawdata_revoked_details, summary):
    """生成诊断报告"""
    log_diagnostic("\n" + "=" * 80, "INFO")
    log_diagnostic("生成诊断报告", "INFO")
    log_diagnostic("=" * 80, "INFO")
    
    log_diagnostic("\n📊 诊断结果汇总:", "INFO")
    log_diagnostic(f"\n1. RawData 文件中的 Revoked 数据:", "INFO")
    log_diagnostic(f"   - Revoked 记录数: {rawdata_revoked_count}", "INFO")
    if rawdata_revoked_details:
        log_diagnostic(f"   - 详细分布:", "INFO")
        for detail in rawdata_revoked_details:
            log_diagnostic(f"     * {detail['category']}: {detail['count']} 条 (来自列: {detail['status_column']})", "INFO")
    else:
        log_diagnostic(f"   - 没有详细分布信息", "INFO")
    
    if summary:
        log_diagnostic(f"\n2. get_processing_summary() 返回的 Revoked 数据:", "INFO")
        log_diagnostic(f"   - Revoked 记录数: {summary.get('revoked_count', 'N/A')}", "INFO")
        log_diagnostic(f"   - Revoked Categories: {summary.get('revoked_categories', 'N/A')}", "INFO")
        
        # 对比分析
        summary_revoked_count = summary.get('revoked_count', 0)
        if rawdata_revoked_count != summary_revoked_count:
            log_diagnostic(f"\n⚠️ 数据不一致!", "WARNING")
            log_diagnostic(f"   RawData 文件: {rawdata_revoked_count} 条", "WARNING")
            log_diagnostic(f"   Summary 返回: {summary_revoked_count} 条", "WARNING")
            log_diagnostic(f"   差异: {abs(rawdata_revoked_count - summary_revoked_count)} 条", "WARNING")
        else:
            log_diagnostic(f"\n✅ 数据一致: RawData 和 Summary 的 revoked_count 都是 {rawdata_revoked_count}", "INFO")
    
    log_diagnostic("\n" + "=" * 80, "INFO")
    log_diagnostic("💡 问题分析和建议:", "INFO")
    log_diagnostic("=" * 80, "INFO")
    
    if rawdata_revoked_count == 0 and (summary and summary.get('revoked_count', 0) > 0):
        log_diagnostic("\n🔴 问题: RawData 中没有 revoked 记录，但 get_processing_summary() 返回了非零值", "WARNING")
        log_diagnostic("\n可能原因:", "INFO")
        log_diagnostic("1. get_processing_summary() 从旧的日志文件中读取了过期数据", "INFO")
        log_diagnostic("2. 日志文件未被及时更新或清理", "INFO")
        log_diagnostic("3. 日志解析正则表达式匹配了错误的内容", "INFO")
        log_diagnostic("\n建议解决方案:", "INFO")
        log_diagnostic("1. 修改 get_processing_summary() 函数，只读取最新的日志文件", "INFO")
        log_diagnostic("2. 在处理新数据前清理旧的日志文件", "INFO")
        log_diagnostic("3. 改进日志解析逻辑，添加时间戳验证", "INFO")
        log_diagnostic("4. 考虑直接从数据处理结果读取，而不是从日志解析", "INFO")
        
    elif rawdata_revoked_count > 0 and (summary and summary.get('revoked_count', 0) == 0):
        log_diagnostic("\n🔴 问题: RawData 中有 revoked 记录，但 get_processing_summary() 返回了 0", "WARNING")
        log_diagnostic("\n可能原因:", "INFO")
        log_diagnostic("1. pending_review_report.py 没有正确处理 revoked 数据", "INFO")
        log_diagnostic("2. 日志输出格式不正确", "INFO")
        log_diagnostic("3. get_processing_summary() 的日志解析逻辑有问题", "INFO")
        log_diagnostic("\n建议解决方案:", "INFO")
        log_diagnostic("1. 检查 pending_review_report.py 的 revoked 处理逻辑", "INFO")
        log_diagnostic("2. 确保 revoked 数据被正确写入日志", "INFO")
        log_diagnostic("3. 验证 get_processing_summary() 的正则表达式是否正确", "INFO")
        
    elif rawdata_revoked_count == 0 and (summary and summary.get('revoked_count', 0) == 0):
        log_diagnostic("\n✅ 结论: 数据正常", "INFO")
        log_diagnostic("RawData 中确实没有 revoked 记录，Summary 返回也正确", "INFO")
        log_diagnostic("\n如果您在 Teams 通知中看到 revoked 数据:", "INFO")
        log_diagnostic("1. 可能是使用了测试数据", "INFO")
        log_diagnostic("2. 建议检查 teams_sender.py 的测试代码", "INFO")
        log_diagnostic("3. 确保生产环境使用真实的 log_summary 数据", "INFO")
        
    else:
        log_diagnostic("\n✅ 数据一致且正常", "INFO")
        log_diagnostic(f"RawData 和 Summary 都显示有 {rawdata_revoked_count} 条 revoked 记录", "INFO")


def main():
    """主函数"""
    log_diagnostic("🔧 Revoked 数据诊断工具", "INFO")
    log_diagnostic("=" * 80, "INFO")
    log_diagnostic(f"诊断开始时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "INFO")
    log_diagnostic("=" * 80, "INFO")
    
    # 1. 查找最新的 RawData 文件
    log_diagnostic("\n📍 步骤 1: 查找最新的 RawData 文件", "INFO")
    latest_file = find_latest_rawdata_file()
    
    rawdata_revoked_count = 0
    rawdata_revoked_details = []
    
    if latest_file:
        # 2. 分析文件中的 Revoked 数据
        log_diagnostic("\n📍 步骤 2: 分析文件中的 Revoked 数据", "INFO")
        rawdata_revoked_count, rawdata_revoked_details = analyze_revoked_data(latest_file)
    else:
        log_diagnostic("\n⚠️ 无法继续分析，未找到 RawData 文件", "WARNING")
    
    # 3. 检查处理脚本
    log_diagnostic("\n📍 步骤 3: 检查 pending_review_report.py", "INFO")
    check_pending_review_script()
    
    # 4. 检查日志文件
    log_diagnostic("\n📍 步骤 4: 检查最新的处理日志", "INFO")
    check_latest_log()
    
    # 5. 检查 get_processing_summary 函数
    log_diagnostic("\n📍 步骤 5: 检查 get_processing_summary 函数", "INFO")
    summary = check_get_processing_summary()
    
    # 6. 生成诊断报告
    log_diagnostic("\n📍 步骤 6: 生成诊断报告", "INFO")
    generate_diagnostic_report(rawdata_revoked_count, rawdata_revoked_details, summary)
    
    # 7. 保存诊断日志
    log_diagnostic("\n" + "=" * 80, "INFO")
    log_diagnostic("✅ 诊断完成！", "INFO")
    log_diagnostic(f"诊断结束时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", "INFO")
    log_diagnostic("=" * 80, "INFO")
    
    log_path = save_diagnostic_log()
    
    if log_path:
        print(f"\n📄 完整诊断日志已保存到:")
        print(f"   {log_path}")
        print(f"\n💡 您可以查看此日志文件获取详细的诊断信息")


if __name__ == "__main__":
    main()
    input("\n按回车键退出...")