# -*- coding: utf-8 -*-
"""
美股情绪指标计算 —— 对应 A 股的 sentiment_indicator.py

读 us_market_data.csv, 对四个子指标各算滚动 252 日分位, 等权合成 0~100 综合指标,
写出 us_sentiment_result.csv (列名与网页前端约定一致)。

与 A 股的唯一方法论差异: **VIX 反向取分位**。
VIX 高 = 恐慌 = 情绪低, 方向与其余三项相反, 因此用 100 - 分位 参与合成。
存进 margin_pct 的仍是 VIX 原值, 图表展示的是真实 VIX。
"""

import os
import sys

import numpy as np
import pandas as pd

from sentiment_core import rolling_percentile_rank, WINDOW

# 确保控制台输出 UTF-8, 避免 Windows GBK 下打印 • 等字符直接抛 UnicodeEncodeError
# (A 股的 auto_fetch_daily.py / update.py 都有这段, 这两个脚本原来漏了)
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE_DIR, 'us_market_data.csv')
OUT = os.path.join(BASE_DIR, 'us_sentiment_result.csv')

# 前端 rawData 约定的键名, 与 A 股共用同一套 (含义由页面的 MARKETS 配置决定)
OUT_COLS = ['date', 'turnover', 'top3_ind', 'rise_pct', 'margin_pct',
            'pct_turnover', 'pct_top3_ind', 'pct_rise_pct', 'pct_margin_pct',
            'composite_sentiment']

# 合成综合值所需的最少子指标个数。
# 美股休市日 (如阵亡将士纪念日、劳动节) yfinance 仍会给出 VIX 行, 而板块 ETF 与
# 成分股都没有数据, outer merge 于是造出只有 VIX 一项的行; 那样算出来的"综合值"
# 其实就是 VIX 单项分位冒充四指标合成。要求至少 3 项才出值, 把这类行挡掉。
# (取 3 而非 4 是与 A 股一致: A 股 2010-03-31 前没有两融, 合法地只有 3 项。)
MIN_SUBS = 3


def main():
    if not os.path.exists(SRC):
        print(f"[错误] 未找到 {SRC}, 请先运行 python us_fetch_daily.py")
        sys.exit(1)

    df = pd.read_csv(SRC)
    df['date'] = pd.to_datetime(df['date'])
    df = df.sort_values('date').reset_index(drop=True)
    print(f"载入 {len(df)} 个交易日 ({df['date'].iloc[0].date()} ~ {df['date'].iloc[-1].date()})")

    out = pd.DataFrame({'date': df['date']})
    out['turnover'] = df['dollar_volume_b']
    out['top3_ind'] = df['top3_share_pct']
    out['rise_pct'] = df['rise_pct']
    out['margin_pct'] = df['vix']

    print(f"计算滚动 {WINDOW} 日分位数...")
    out['pct_turnover'] = rolling_percentile_rank(out['turnover'], WINDOW)
    out['pct_top3_ind'] = rolling_percentile_rank(out['top3_ind'], WINDOW)
    out['pct_rise_pct'] = rolling_percentile_rank(out['rise_pct'], WINDOW)

    # VIX 反向: 高 VIX = 恐慌 = 情绪低
    vix_pct = rolling_percentile_rank(out['margin_pct'], WINDOW)
    out['pct_margin_pct'] = 100.0 - vix_pct
    print("  VIX 已反向取分位 (100 - 原始分位)")

    pct_cols = ['pct_turnover', 'pct_top3_ind', 'pct_rise_pct', 'pct_margin_pct']
    n_subs = out[pct_cols].notna().sum(axis=1)
    out['composite_sentiment'] = out[pct_cols].mean(axis=1)   # 缺失项跳过, 与 A 股一致
    out.loc[n_subs < MIN_SUBS, 'composite_sentiment'] = np.nan

    dropped = out.loc[(n_subs > 0) & (n_subs < MIN_SUBS), 'date']
    if len(dropped):
        print(f"  剔除子指标不足 {MIN_SUBS} 项的交易日 {len(dropped)} 个 (多为美股休市日): "
              + ', '.join(d.strftime('%Y-%m-%d') for d in dropped))

    out = out[out['composite_sentiment'].notna()].copy()
    if out.empty:
        print(f"[错误] 没有任何一天凑齐足够历史 (滚动窗口需要 {int(WINDOW * 0.5)} 个有效值), 不写出")
        sys.exit(1)

    out['date'] = out['date'].dt.strftime('%Y-%m-%d')
    out = out[OUT_COLS].round(2)
    out.to_csv(OUT, index=False)

    print(f"\n[OK] 写入 {OUT}: {len(out)} 个交易日 ({out['date'].iloc[0]} ~ {out['date'].iloc[-1]})")
    last = out.iloc[-1]
    status = '历史高位' if last['composite_sentiment'] > 80 else ('历史低位' if last['composite_sentiment'] < 20 else '中性')
    print(f"\n最新读数 {last['date']}: 综合情绪 {last['composite_sentiment']:.2f} ({status})")
    fmt = lambda v: '—' if pd.isna(v) else f"{v:.2f}"
    print(f"  • 成交额分位:   {fmt(last['pct_turnover'])}%  (绝对值 {fmt(last['turnover'])} 十亿美元)")
    print(f"  • 板块集中度:   {fmt(last['pct_top3_ind'])}%  (绝对值 {fmt(last['top3_ind'])}%)")
    print(f"  • 上涨占比分位: {fmt(last['pct_rise_pct'])}%  (绝对值 {fmt(last['rise_pct'])}%)")
    print(f"  • VIX 反向分位: {fmt(last['pct_margin_pct'])}%  (VIX {fmt(last['margin_pct'])})")
    print("\n下一步: python generate_html.py")


if __name__ == '__main__':
    main()
