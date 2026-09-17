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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE_DIR, 'us_market_data.csv')
OUT = os.path.join(BASE_DIR, 'us_sentiment_result.csv')

# 前端 rawData 约定的键名, 与 A 股共用同一套 (含义由页面的 MARKETS 配置决定)
OUT_COLS = ['date', 'turnover', 'top3_ind', 'rise_pct', 'margin_pct',
            'pct_turnover', 'pct_top3_ind', 'pct_rise_pct', 'pct_margin_pct',
            'composite_sentiment']


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
    out['composite_sentiment'] = out[pct_cols].mean(axis=1)   # 缺失项跳过, 与 A 股一致

    out = out[out['composite_sentiment'].notna()].copy()
    if out.empty:
        print(f"[错误] 没有任何一天凑齐足够历史 (滚动窗口需要 {int(WINDOW * 0.5)} 个有效值), 不写出")
        sys.exit(1)

    out['date'] = out['date'].dt.strftime('%Y-%m-%d')
    out = out[OUT_COLS].round(2)
    out.to_csv(OUT, index=False)

    print(f"\n[OK] 写入 {OUT}: {len(out)} 个交易日 ({out['date'].iloc[0]} ~ {out['date'].iloc[-1]})")
    last = out.iloc[-1]
    status = '过热' if last['composite_sentiment'] > 80 else ('过冷' if last['composite_sentiment'] < 20 else '中性')
    print(f"\n最新读数 {last['date']}: 综合情绪 {last['composite_sentiment']:.2f} ({status})")
    fmt = lambda v: '—' if pd.isna(v) else f"{v:.2f}"
    print(f"  • 成交额分位:   {fmt(last['pct_turnover'])}%  (绝对值 {fmt(last['turnover'])} 十亿美元)")
    print(f"  • 板块集中度:   {fmt(last['pct_top3_ind'])}%  (绝对值 {fmt(last['top3_ind'])}%)")
    print(f"  • 上涨占比分位: {fmt(last['pct_rise_pct'])}%  (绝对值 {fmt(last['rise_pct'])}%)")
    print(f"  • VIX 反向分位: {fmt(last['pct_margin_pct'])}%  (VIX {fmt(last['margin_pct'])})")
    print("\n下一步: python generate_html.py")


if __name__ == '__main__':
    main()
