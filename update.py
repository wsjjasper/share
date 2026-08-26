# -*- coding: utf-8 -*-
"""
A股情绪指标一键全自动更新脚本 (One-Click Auto Update & Spot Check Pipeline)
使用步骤：
1. 运行本脚本：python update.py (或双击 一键更新.bat)
2. 脚本将全自动完成：
   - 步骤 1: 自动拉取公开金融数据，执行 Spot Check 历史交叉比对与数据补全
   - 步骤 2: 重新计算 252 交易日滚动百分位与 4 大指标综合情绪值
   - 步骤 3: 重新生成高清研报走势图表
   - 步骤 4: 自动生成多周期交互式 index.html 与 docs 网页
   - 步骤 5: 自动同步文件至 docs 并 Git Commit & Push 至 GitHub 仓库
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

    # 1. 自动抓取与 Spot Check 校验
    print_step("步骤 1: 自动化数据抓取与 Spot Check 交叉校验")
    res_fetch = subprocess.run([sys.executable, "auto_fetch_daily.py"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    print(res_fetch.stdout)
    if res_fetch.stderr:
        print(f"[提示] 抓取日志:\n{res_fetch.stderr}")

    # 2. 检查原始数据文件
    print_step("步骤 2: 检查原始数据文件状态")
    raw_excel = os.path.join(WORKSPACE_DIR, '副本万得全A.xlsx')
    if not os.path.exists(raw_excel):
        raw_excel = os.path.join(WORKSPACE_DIR, '副本万得全A_latest.xlsx')
        
    df_raw = pd.read_excel(raw_excel)
    print(f"[OK] 原始数据已载入: 共 {len(df_raw)} 行历史数据")
    print(f"     起始日期: {df_raw.iloc[0, 1]}")
    print(f"     最新日期: {df_raw.iloc[-1, 1]}")

    # 3. 运行核心量化计算
    print_step("步骤 3: 计算滚动252日分位数与综合指标")
    res_calc = subprocess.run([sys.executable, "sentiment_indicator.py"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    if res_calc.returncode != 0:
        print(f"[错误] 计算失败:\n{res_calc.stderr}")
        sys.exit(1)
    print(res_calc.stdout)

    # 4. 复制图表为 sentiment_chart.png
    src_chart = os.path.join(WORKSPACE_DIR, '情绪指标_图表.png')
    dst_chart = os.path.join(WORKSPACE_DIR, 'sentiment_chart.png')
    if os.path.exists(src_chart):
        shutil.copy2(src_chart, dst_chart)

    # 5. 生成交互式网页 index.html
    print_step("步骤 4: 重新生成交互式研报网页 (index.html)")
    res_html = subprocess.run([sys.executable, "generate_html.py"], capture_output=True, text=True, encoding='utf-8', errors='ignore')
    if res_html.returncode != 0:
        print(f"[错误] 网页生成失败:\n{res_html.stderr}")
        sys.exit(1)
    print(res_html.stdout)

    # 6. 同步所有必要资产至 docs/ 目录
    print_step("步骤 5: 同步文件至 docs 目录")
    os.makedirs(DOCS_DIR, exist_ok=True)
    files_to_docs = [
        'index.html',
        'sentiment_chart.png',
        'sentiment_indicator.py',
        'generate_html.py',
        'auto_fetch_daily.py',
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

    # 7. 读取最新数据并打印看板总结
    print_step("步骤 6: 最新读数速览")
    excel_res_path = os.path.join(WORKSPACE_DIR, '情绪指标_结果.xlsx')
    if not os.path.exists(excel_res_path):
        excel_res_path = os.path.join(WORKSPACE_DIR, '情绪指标_结果_latest.xlsx')
    df_res = pd.read_excel(excel_res_path)
    cols = ['date', 'turnover', 'top3', 'rise', 'margin', 'p_turnover', 'p_top3', 'p_rise', 'p_margin', 'composite']
    df_res.columns = cols
    latest_row = df_res.iloc[-1]
    
    date_str = str(latest_row['date'])[:10]
    comp_val = latest_row['composite']
    status_text = "过热" if comp_val > 80 else ("过冷" if comp_val < 20 else "中性")
    
    print(f"交易日期: {date_str}")
    print(f"综合情绪指标: {comp_val:.2f} ({status_text})")
    print(f"  • 换手率分位:     {latest_row['p_turnover']:.2f}% (绝对值 {latest_row['turnover']:.2f}%)")
    print(f"  • 行业集中度分位: {latest_row['p_top3']:.2f}% (绝对值 {latest_row['top3']:.2f}%)")
    print(f"  • 上涨个股分位:   {latest_row['p_rise']:.2f}% (绝对值 {latest_row['rise']:.2f}%)")
    if pd.notna(latest_row['p_margin']):
        print(f"  • 融资买入分位:   {latest_row['p_margin']:.2f}% (绝对值 {latest_row['margin']:.2f}%)")
    else:
        print(f"  • 融资买入分位:   未更新")

    # 8. 自动 Git Commit 与 Push
    print_step("步骤 7: 自动提交与推送到 GitHub")
    try:
        subprocess.run(["git", "add", "."], check=True, capture_output=True)
        commit_msg = f"data: auto-fetch & update sentiment data up to {date_str} (sentiment: {comp_val:.2f})"
        subprocess.run(["git", "commit", "-m", commit_msg], capture_output=True)
        print("     [OK] 本地 Git Commit 完成")

        # Push to main and master
        push_res_main = subprocess.run(["git", "push", "origin", "main"], capture_output=True, text=True)
        push_res_master = subprocess.run(["git", "push", "origin", "main:master"], capture_output=True, text=True)
        
        if push_res_main.returncode == 0 or push_res_master.returncode == 0:
            print("     [OK] 远程 GitHub Push 推送成功")
        else:
            print(f"     [警告] Push 出现提示:\n{push_res_main.stderr}\n{push_res_master.stderr}")

    except Exception as e:
        print(f"     [错误] Git 操作失败: {e}")

    print("\n" + "="*50)
    print("更新完成！网页与数据已全部刷新。")
    print("在线访问地址: https://wsjjasper.github.io/share/")
    print("="*50 + "\n")

if __name__ == '__main__':
    main()
