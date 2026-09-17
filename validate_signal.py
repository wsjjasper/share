# -*- coding: utf-8 -*-
"""
信号有效性检验 —— 回答"这个指标能不能用来择时", 而不是"数据抓得对不对"。

仓库里已经有一整套**数据口径**的校验门 (backfill_industry_sw / backfill_breadth_sina /
rebuild_turnover_sw, 不达标拒写)。这个脚本把同样的纪律延伸到**信号层**: 改了子指标、
换了口径、调了权重之后, 跑一遍看它到底有没有前瞻信息, 而不是靠感觉判断。

四项检验:

1. IC (信息系数)  —— 指标与未来 5/20/60 日收益的秩相关。
2. 非重叠显著性    —— 关键一项。重叠窗口会把 p 值严重高估: 用重叠样本算 60 日 IC
                     可以得到 p=1.8e-16, 看着无可辩驳; 但把样本切成 60 组互不重叠
                     的子样本后, 只有 12% 的组显著 —— 与随机无异。真有信号的话,
                     绝大多数子样本都该显著。
3. 分档单调性      —— 按指标高低分 5 档看后续收益。一个"过热该谨慎"的指标, 收益
                     应当随档位单调下降。实测 A 股 >80 档后续 20 日收益 +0.55%,
                     反而高于全样本 +0.31% —— 阈值语义不成立。
4. 同步性          —— 指标与**当日**收益的相关。这项高说明它在描述已经发生的行情,
                     不是在预测。实测 A 股综合 0.42、上涨个股占比 0.80 —— 后者几乎
                     就是"今天涨没涨"的另一种写法。

用法:
    python validate_signal.py                 # 两个市场都跑
    python validate_signal.py --market cn
    python validate_signal.py --market us --horizons 5,20
    python validate_signal.py --start 2018-01-01 --end 2026-08-25
"""

import argparse
import json
import math
import os
import sys

import numpy as np
import pandas as pd
import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# A 股口径一致区间: 换手率自 2015-01-01 起是申万口径(见 rebuild_turnover_sw.py),
# 行业集中度与宽度在 2026-08-26 之前是 Wind 原值。留一年让滚动窗口填满, 故从 2016 起。
CN_DEFAULT_START, CN_DEFAULT_END = '2016-01-01', '2026-08-26'
SIG_LEVEL = 0.05
COINCIDENT_FLAG = 0.30      # 与当日收益相关高于此值即视为同步指标
MIN_BUCKET_N = 30           # 分档样本少于此数则单调性不作数
HORIZONS = (5, 20, 60)


# ---------- 统计 ----------

def spearman(x, y):
    """秩相关 + 正态近似 p 值 (避免引入 scipy; n>30 时足够用)。"""
    x, y = np.asarray(x, float), np.asarray(y, float)
    ok = ~(np.isnan(x) | np.isnan(y))
    x, y = x[ok], y[ok]
    n = len(x)
    if n < 10:
        return np.nan, np.nan, n
    rx = pd.Series(x).rank().to_numpy()
    ry = pd.Series(y).rank().to_numpy()
    if rx.std() == 0 or ry.std() == 0:
        return np.nan, np.nan, n
    ic = float(np.corrcoef(rx, ry)[0, 1])
    z = ic * math.sqrt(max(n - 1, 1))
    p = 2.0 * (1.0 - 0.5 * (1.0 + math.erf(abs(z) / math.sqrt(2.0))))
    return ic, p, n


def non_overlap_check(ind, fwd, h):
    """
    把重叠样本切成 h 组互不重叠的子样本, 每组独立算 IC。

    返回 (显著组占比, IC 中位数, 组数)。这是判断 IC 是否真实的关键一步 ——
    重叠窗口下相邻观测高度自相关, 朴素 p 值会小到毫无意义。
    """
    ics, sigs = [], []
    for off in range(h):
        sx, sy = ind[off::h], fwd[off::h]
        ic, p, n = spearman(sx, sy)
        if not np.isnan(ic) and n >= 20:
            ics.append(ic)
            sigs.append(p < SIG_LEVEL)
    if not ics:
        return np.nan, np.nan, 0
    return float(np.mean(sigs)), float(np.median(ics)), len(ics)


# ---------- 取数 ----------

def load_cn(start, end):
    """A 股: 用全部历史重算四项分位再合成 —— 结果文件只有 2024-09-24 起, 样本太少。"""
    from sentiment_core import rolling_percentile_rank, WINDOW

    xls = os.path.join(BASE_DIR, '副本万得全A.xlsx')
    if not os.path.exists(xls):
        raise RuntimeError(f'未找到 {xls}')
    df = pd.read_excel(xls)
    d = pd.to_datetime(df.iloc[:, 1], errors='coerce')
    sub = pd.DataFrame({
        'date': d,
        '换手率': pd.to_numeric(df.iloc[:, 2], errors='coerce'),
        '行业集中度': pd.to_numeric(df.iloc[:, 3], errors='coerce') * 100,
        '上涨占比': pd.to_numeric(df.iloc[:, 4], errors='coerce') * 100,
        '融资买入': pd.to_numeric(df.iloc[:, 5], errors='coerce') * 100,
    }).dropna(subset=['date'])
    sub.loc[sub['融资买入'] == 0, '融资买入'] = np.nan     # 与 sentiment_indicator 一致
    subs = ['换手率', '行业集中度', '上涨占比', '融资买入']
    for c in subs:
        sub['p_' + c] = rolling_percentile_rank(sub[c], WINDOW)
    sub['composite'] = sub[['p_' + c for c in subs]].mean(axis=1)
    sub['date'] = sub['date'].dt.normalize()

    # 万得全A 指数免费拿不到, 用上证综指代理 (方向性足够, 绝对幅度会有差异)
    r = requests.get(
        'http://money.finance.sina.com.cn/quotes_service/api/json_v2.php/CN_MarketData.getKLineData',
        params={'symbol': 'sh000001', 'scale': 240, 'ma': 'no', 'datalen': 3000},
        headers={'User-Agent': 'Mozilla/5.0', 'Referer': 'http://finance.sina.com.cn/'},
        timeout=(8, 25))
    k = pd.DataFrame(json.loads(r.text))
    idx = pd.DataFrame({'date': pd.to_datetime(k['day']).dt.normalize(),
                        'close': pd.to_numeric(k['close'])})

    m = sub.merge(idx, on='date').sort_values('date').reset_index(drop=True)
    m = m[(m['date'] >= start) & (m['date'] < end)].reset_index(drop=True)
    return m, subs, '上证综指'


def load_us(start, end):
    """美股: 直接用结果文件 + 标普500。"""
    import yfinance as yf

    csv = os.path.join(BASE_DIR, 'us_sentiment_result.csv')
    if not os.path.exists(csv):
        raise RuntimeError(f'未找到 {csv}, 先跑 us_fetch_daily.py 与 us_sentiment_indicator.py')
    u = pd.read_csv(csv)
    u['date'] = pd.to_datetime(u['date']).dt.normalize()
    ren = {'pct_turnover': 'p_成交额', 'pct_top3_ind': 'p_板块集中度',
           'pct_rise_pct': 'p_上涨占比', 'pct_margin_pct': 'p_VIX反向',
           'composite_sentiment': 'composite'}
    u = u.rename(columns=ren)
    subs = ['成交额', '板块集中度', '上涨占比', 'VIX反向']

    sp = yf.download('^GSPC', start='2022-09-01', end='2026-12-31',
                     progress=False, auto_adjust=False)
    cl = sp['Close']
    cl = cl.iloc[:, 0] if hasattr(cl, 'columns') else cl
    idx = pd.DataFrame({'date': pd.to_datetime(cl.index).normalize(), 'close': cl.values})

    m = u.merge(idx, on='date').sort_values('date').reset_index(drop=True)
    if start:
        m = m[m['date'] >= start]
    if end:
        m = m[m['date'] < end]
    return m.reset_index(drop=True), subs, '标普500'


# ---------- 检验 ----------

def report(name, m, subs, idx_name, horizons):
    print('=' * 68)
    print(f'{name}  ({len(m)} 个交易日, {m["date"].min():%Y-%m-%d} ~ {m["date"].max():%Y-%m-%d}, '
          f'基准 {idx_name})')
    print('=' * 68)
    if len(m) < 200:
        print('[警告] 样本不足 200 个交易日, 下面任何结论都不可信\n')

    for h in horizons:
        m[f'fwd{h}'] = m['close'].shift(-h) / m['close'] - 1

    verdict = {'ic': False, 'mono': False, 'coincident': False}

    print('\n【1/4】综合指标 IC 与未来收益')
    print(f"{'持有期':<8}{'IC':>9}{'朴素p':>10}{'有效样本':>10}{'非重叠显著占比':>16}")
    for h in horizons:
        ic, p, n = spearman(m['composite'], m[f'fwd{h}'])
        v = m[['composite', f'fwd{h}']].dropna()
        share, med, ngrp = non_overlap_check(v['composite'].to_numpy(),
                                             v[f'fwd{h}'].to_numpy(), h)
        share_s = '—' if np.isnan(share) else f'{share*100:.0f}% ({ngrp}组)'
        print(f"{str(h)+'日':<8}{ic:>9.4f}{p:>10.4f}{n//h:>10d}{share_s:>16}")
        # 只有朴素显著 *且* 多数非重叠子样本也显著, 才算真有信号
        if p < SIG_LEVEL and not np.isnan(share) and share >= 0.5:
            verdict['ic'] = True

    print('\n【2/4】分档后续收益 (以 20 日为准)')
    hh = 20 if 20 in horizons else horizons[0]
    v = m[['composite', f'fwd{hh}']].dropna().copy()
    labels = ['<20', '20-40', '40-60', '60-80', '>80']
    v['档'] = pd.cut(v['composite'], [0, 20, 40, 60, 80, 100], labels=labels)
    print(f"{'档位':<9}{'天数':>7}{'平均':>9}{'中位':>9}{'胜率':>8}")
    means, counts = [], []
    for lab in labels:
        s = v.loc[v['档'] == lab, f'fwd{hh}']
        if len(s) == 0:
            means.append(np.nan); counts.append(0); continue
        means.append(s.mean()); counts.append(len(s))
        print(f"{lab:<9}{len(s):>7}{s.mean()*100:>8.2f}%{s.median()*100:>8.2f}%{(s>0).mean()*100:>7.1f}%")
    print(f"{'全样本':<9}{len(v):>7}{v[f'fwd{hh}'].mean()*100:>8.2f}%"
          f"{v[f'fwd{hh}'].median()*100:>8.2f}%{(v[f'fwd{hh}']>0).mean()*100:>7.1f}%")
    ok = [x for x in means if not np.isnan(x)]
    if len(ok) >= 4:
        # 只有 5 个档位, 走不了 spearman() 的 n>=10 门槛, 这里直接算秩相关
        a = pd.Series(np.arange(len(ok))).rank().to_numpy()
        b = pd.Series(ok).rank().to_numpy()
        mono = float(np.corrcoef(a, b)[0, 1])
        thin = [c for c in counts if 0 < c < MIN_BUCKET_N]
        print(f"  档位与收益的秩相关: {mono:+.3f}  "
              f"(要支持'高分位=该谨慎', 应显著为负)")
        # 只有 5 个点, 单调性本身证据很弱; 极端档样本又常常很少。
        # 不加这道限制, 一个 14 天的极端档就能把秩相关拉到 -1.000 并打出"成立"。
        if thin:
            print(f"  [注意] 有 {len(thin)} 个档位样本不足 {MIN_BUCKET_N} 天 "
                  f"(最少 {min(thin)} 天), 单调性不可当作证据")
        verdict['mono'] = (mono <= -0.8) and not thin

    print('\n【3/4】同步性 —— 与当日收益的相关')
    m['ret0'] = m['close'].pct_change()
    ic0, p0, _ = spearman(m['composite'], m['ret0'])
    print(f"  综合指标  IC={ic0:>7.4f}")
    for c in subs:
        col = 'p_' + c
        if col in m.columns:
            ics, _, _ = spearman(m[col], m['ret0'])
            tag = '  <- 与当日行情高度重合' if abs(ics) >= 0.6 else ''
            print(f"  {c:<12}IC={ics:>7.4f}{tag}")
    verdict['coincident'] = abs(ic0) >= COINCIDENT_FLAG

    print(f'\n【4/4】各子指标对未来 {hh} 日收益的 IC (Bonferroni 阈值 '
          f'{SIG_LEVEL/len(subs):.4f})')
    for c in subs:
        col = 'p_' + c
        if col not in m.columns:
            continue
        ic, p, _ = spearman(m[col], m[f'fwd{hh}'])
        mark = '显著' if p < SIG_LEVEL / len(subs) else ('擦线' if p < SIG_LEVEL else '不显著')
        print(f"  {c:<12}IC={ic:>7.4f}  p={p:>6.4f}  {mark}")

    print('\n--- 判定 ---')
    if verdict['ic']:
        print('  ✓ 存在通过非重叠检验的前瞻信号')
    else:
        print('  ✗ 没有任何持有期的前瞻信号能通过非重叠检验 —— 不支持用于择时')
    if verdict['mono']:
        print('  ~ 分档收益随指标单调下降 (5 个点的单调性只是弱证据, 不等于可交易)')
    else:
        print('  ✗ 分档收益不单调或样本过薄 —— "高分位=过热该谨慎"没有数据支持')
    if verdict['coincident']:
        print(f'  ! 与当日收益相关 {ic0:.2f} —— 这是同步(描述)指标, 不是领先指标')
    print()
    return verdict


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--market', default='both', choices=['cn', 'us', 'both'])
    ap.add_argument('--start', default=None)
    ap.add_argument('--end', default=None)
    ap.add_argument('--horizons', default=','.join(str(h) for h in HORIZONS))
    args = ap.parse_args()
    hs = tuple(int(x) for x in args.horizons.split(',') if x.strip())

    if args.market in ('cn', 'both'):
        try:
            m, subs, idx = load_cn(args.start or CN_DEFAULT_START, args.end or CN_DEFAULT_END)
            report('A 股', m, subs, idx, hs)
        except Exception as e:
            print(f'[A股] 检验失败: {type(e).__name__}: {e}\n')

    if args.market in ('us', 'both'):
        try:
            m, subs, idx = load_us(args.start, args.end)
            report('美股', m, subs, idx, hs)
        except Exception as e:
            print(f'[美股] 检验失败: {type(e).__name__}: {e}\n')

    print('提示: 这是对指标统计属性的检验, 不构成投资建议。')


if __name__ == '__main__':
    main()
