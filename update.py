# -*- coding: utf-8 -*-
"""
A股情绪指标一键全自动更新脚本 (One-Click Auto Update Script)
使用步骤：
1. 将 Wind 最新导出的数据保存或追加至 '副本万得全A.xlsx'
2. 运行本脚本：python update.py (或双击 一键更新.bat)
3. 脚本将自动完成：
   - 重新计算 252 交易日滚动百分位与综合情绪指标
   - 重新生成高分辨率走势图表
   - 自动生成交互式 index.html 与 docs 网页
   - 自动同步文件并 Git Commit & Push 至 GitHub
"""

import os
import sys
import subprocess
import shutil
import pandas as pd

# 确保控制台输出 UTF-8 编码避免 Windows GBK 报错
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

WORKSPACE_DIR = r'D:\vscode_workspace\情绪指标'
DOCS_DIR = os.path.join(WORKSPACE_DIR, 'docs')

def print_step(step_name):
    print(f"\n{'='*50}")
    print(f">> {step_name}")
    print(f"{'='*50}")

def main():
    os.chdir(WORKSPACE_DIR)

    # 1. 检查原始数据文件
    print_step("步骤 1: 检查原始数据文件")
    raw_excel = os.path.join(WORKSPACE_DIR, '副本万得全A.xlsx')
    if not os.path.exists(raw_excel):
        print(f"[错误] 未找到原始数据文件: {raw_excel}")
        sys.exit(1)
    
    df_raw = pd.read_excel(raw_excel)
    print(f"[OK] 原始数据已载入: 共 {len(df_raw)} 行历史数据")
    print(f"     起始日期: {df_raw.iloc[0, 1]}")
    print(f"     最新日期: {df_raw.iloc[-1, 1]}")

    # 2. 运行核心量化计算
    print_step("步骤 2: 计算滚动252日分位数与综合指标")
    res = subprocess.run([sys.executable, "sentiment_indicator.py"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    if res.returncode != 0:
        print(f"[错误] 计算失败:\n{res.stderr}")
        sys.exit(1)
    print(res.stdout)

    # 3. 复制图表为 sentiment_chart.png
    src_chart = os.path.join(WORKSPACE_DIR, '情绪指标_图表.png')
    dst_chart = os.path.join(WORKSPACE_DIR, 'sentiment_chart.png')
    if os.path.exists(src_chart):
        shutil.copy2(src_chart, dst_chart)

    # 4. 生成交互式网页 index.html
    print_step("步骤 3: 重新生成交互式研报网页 (index.html)")
    res_html = subprocess.run([sys.executable, "generate_html.py"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    if res_html.returncode != 0:
        print(f"[错误] 网页生成失败:\n{res_html.stderr}")
        sys.exit(1)
    print(res_html.stdout)

    # 5. 同步所有必要资产至 docs/ 目录
    print_step("步骤 4: 同步文件至 docs 目录")
    os.makedirs(DOCS_DIR, exist_ok=True)
    files_to_docs = [
        'index.html',
        'sentiment_chart.png',
        'sentiment_indicator.py',
        'generate_html.py',
        '情绪指标_结果.xlsx',
        '情绪指标_图表.png',
        '副本万得全A.xlsx',
        'update.py',
        '一键更新.bat'
    ]
    for f in files_to_docs:
        src = os.path.join(WORKSPACE_DIR, f)
        dst = os.path.join(DOCS_DIR, f)
        if os.path.exists(src):
            shutil.copy2(src, dst)
            print(f"     已同步: docs/{f}")

    # 6. 读取最新数据并打印看板总结
    print_step("步骤 5: 最新读数速览")
    df_res = pd.read_excel(os.path.join(WORKSPACE_DIR, '情绪指标_结果.xlsx'))
    cols = ['date', 'turnover', 'top3', 'rise', 'margin', 'p_turnover', 'p_top3', 'p_rise', 'p_margin', 'composite']
    df_res.columns = cols
    latest_row = df_res.iloc[-1]
    
    date_str = pd.Timestamp(latest_row['date']).strftime('%Y-%m-%d')
    comp = latest_row['composite']
    status = "过热 (>80)" if comp > 80 else ("过冷 (<20)" if comp < 20 else "中性偏冷" if comp < 50 else "中性偏暖")
    
    print(f"交易日期: {date_str}")
    print(f"综合情绪指标: {comp:.2f} ({status})")
    print(f"  • 换手率分位:     {latest_row['p_turnover']:.2f}% (绝对值 {latest_row['turnover']:.2f}%)")
    print(f"  • 行业集中度分位: {latest_row['p_top3']:.2f}% (绝对值 {latest_row['top3']:.2f}%)")
    print(f"  • 上涨个股分位:   {latest_row['p_rise']:.2f}% (绝对值 {latest_row['rise']:.2f}%)")
    margin_str = f"{latest_row['p_margin']:.2f}%" if pd.notnull(latest_row['p_margin']) else "未更新"
    print(f"  • 融资买入分位:   {margin_str}")

    # 7. Git 自动提交与推送
    print_step("步骤 6: 自动提交与推送到 GitHub")
    commit_msg = f"data: update sentiment data up to {date_str} (sentiment: {comp:.2f})"
    
    try:
        subprocess.run(["git", "add", "."], check=True)
        # 检查是否有文件改动
        status_res = subprocess.run(["git", "status", "--porcelain"], capture_output=True, text=True)
        if status_res.stdout.strip():
            subprocess.run(["git", "commit", "-m", commit_msg], check=True)
            print("     [OK] 本地 Git Commit 完成")
            
            # Push 到 main 和 master
            subprocess.run(["git", "push", "origin", "main"], check=True)
            subprocess.run(["git", "push", "origin", "main:master"], check=False)
            print("     [OK] 远程 GitHub Push 推送成功")
        else:
            print("     [Info] 数据与代码已是最新，无需重复提交")
    except Exception as e:
        print(f"     [提示] Git 推送结果: {e}")

    print(f"\n{'='*50}")
    print("更新完成！网页与数据已全部刷新。")
    print("在线访问地址: https://wsjjasper.github.io/share/")
    print(f"{'='*50}\n")

if __name__ == '__main__':
    main()
