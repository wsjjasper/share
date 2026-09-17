# -*- coding: utf-8 -*-
"""
用申万一级行业数据重建整列「换手率」(一次性脚本)

为什么要整列重建, 而不是像行业集中度那样回填几行:
    Wind 原来的换手率列无法用任何外部免费源复现。实测 51 个交易日:
      · Wind 成交额 vs 交易所官方成交额   相关 1.0000  (成交额完全对得上)
      · 交易所换手率 vs 交易所成交额       相关 0.9876  (真实换手率该有的样子)
      · Wind 换手率 vs 交易所换手率        相关 0.6770
      · Wind 换手率 vs Wind 自己的成交额   相关 0.6810  <- 关键
    Wind 这一列相对任何真实市值序列都带约 12% 的特异波动: 它的隐含分母日环比
    中位数 4.05%, 而真实流通市值只有 1.26% —— 真实市值不可能这样跳。
    所以它不是 成交额/市值, 重建不出来。

    但滚动百分位只要求「列内口径一致」, 不要求匹配 Wind。把整列换成一个真实、
    内部自洽的序列, 比留着一列无法维护、也无法续写的数字更有用。

口径:
    换手率 = Σ(换手率ᵢ × 流通市值ᵢ) / Σ(流通市值ᵢ)   (申万一级 31 个行业)
    申万的「流通市值」是自由流通口径, 分母比传统流通市值小约 2.7 倍, 因此数值
    水平比 Wind 原列高约 2.7 倍。这不影响分位数 —— 分位数只看列内相对位置。

    重建覆盖 Excel 的全部历史(申万一级可回溯到 2000 年), 因此不留口径断点。

用法:
    python rebuild_turnover_sw.py                  # 只校验+预览, 不写入 (默认)
    python rebuild_turnover_sw.py --apply          # 校验通过后写入
    python rebuild_turnover_sw.py --refetch        # 忽略缓存重新抓取
"""

import argparse
import os
import sys

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, '副本万得全A.xlsx')
# *.tmp 已在 .gitignore 中; 抓一次 26 年要几分钟, 校验和写入两次运行复用同一份
CACHE_PATH = os.path.join(BASE_DIR, '_sw_turnover_cache.tmp')

COL_DATE, COL_TURNOVER, COL_TOTAL_AMT = 1, 2, 6

# 重建后的自检门槛。这里不能再跟 Wind 比 —— 目标就是换掉 Wind 那一列,
# 所以改成检查"它确实是一个真实的换手率序列"。
MIN_COVERAGE = 0.95      # 区间内 Excel 交易日被覆盖的比例
SANE_LO, SANE_HI = 0.2, 20.0     # 单日换手率的合理区间 (%)
MIN_CORR_AMT = 0.85     # 与成交额的相关; 真实换手率分母缓变, 应当高度同步
MAX_DENOM_MOVE = 0.03   # 隐含分母日环比变动的中位数上限 (真实市值约 1.3%)


def fetch_turnover(start_date, end_date, refetch=False):
    """按区间取回每个交易日的全市场换手率, 复用日常管线的抓取实现。"""
    if not refetch and os.path.exists(CACHE_PATH):
        # convert_axes=False: 否则 pandas 把索引解成 Timestamp,
        # str() 后变成 '2014-02-21 00:00:00', 与 'YYYY-MM-DD' 对不上
        s = pd.read_json(CACHE_PATH, typ='series', convert_axes=False, convert_dates=False)
        print(f"     [缓存] 复用 {CACHE_PATH} ({len(s)} 个交易日); 加 --refetch 可强制重抓")
        return {str(k)[:10]: float(v) for k, v in s.items()}

    from auto_fetch_daily import (_sw_fetch_raw, SW_CHUNK_DAYS, SW_MIN_INDUSTRIES)

    parts, cur, end_ts = [], pd.Timestamp(start_date), pd.Timestamp(end_date)
    while cur <= end_ts:
        chunk_end = min(cur + pd.Timedelta(days=SW_CHUNK_DAYS - 1), end_ts)
        print(f"     · 区间 {cur.date()} ~ {chunk_end.date()}", flush=True)
        part = _sw_fetch_raw(cur.strftime('%Y-%m-%d'), chunk_end.strftime('%Y-%m-%d'))
        if part is not None and len(part) > 0:
            parts.append(part)
        cur = chunk_end + pd.Timedelta(days=1)
    if not parts:
        raise RuntimeError('申万接口未返回可用数据')

    raw = pd.concat(parts, ignore_index=True)
    raw = raw[['发布日期', '指数名称', '换手率', '流通市值']].drop_duplicates(
        subset=['发布日期', '指数名称'])
    raw['发布日期'] = pd.to_datetime(raw['发布日期'])
    for c in ('换手率', '流通市值'):
        raw[c] = pd.to_numeric(raw[c], errors='coerce')

    out = {}
    for d, g in raw.groupby('发布日期'):
        g = g.dropna(subset=['换手率', '流通市值'])
        mcap = g['流通市值'].sum()
        if len(g) < SW_MIN_INDUSTRIES or not (mcap > 0):
            continue
        out[d.strftime('%Y-%m-%d')] = float((g['换手率'] * g['流通市值']).sum() / mcap)

    pd.Series(out).to_json(CACHE_PATH)
    print(f"     [OK] 取到 {len(out)} 个交易日, 已缓存至 {CACHE_PATH}")
    return out


def validate(tr_map, df, start, end):
    """重建序列的自检: 覆盖率、取值合理性、与成交额的同步性、隐含分母的平滑度。"""
    d = pd.to_datetime(df.iloc[:, COL_DATE], errors='coerce').dt.normalize()
    sel = (d >= pd.Timestamp(start)) & (d <= pd.Timestamp(end))
    keys = d[sel].dt.strftime('%Y-%m-%d')
    vals = pd.Series([tr_map.get(k, np.nan) for k in keys], index=keys.index, dtype=float)
    amt = pd.to_numeric(df.iloc[:, COL_TOTAL_AMT], errors='coerce')[sel]

    ok = True
    n_all, n_hit = int(sel.sum()), int(vals.notna().sum())
    cov = n_hit / n_all if n_all else 0.0
    print(f"\n重建自检 ({start} ~ {end})")
    print(f"  Excel 交易日     : {n_all}")
    print(f"  申万可覆盖       : {n_hit}  ({cov * 100:.2f}%)   (阈值 >= {MIN_COVERAGE * 100:.0f}%)")
    if cov < MIN_COVERAGE:
        ok = False

    v = vals.dropna()
    bad = int(((v < SANE_LO) | (v > SANE_HI)).sum())
    print(f"  取值区间         : {v.min():.2f}% ~ {v.max():.2f}%  均值 {v.mean():.2f}%  "
          f"越界 {bad} 个   (合理区间 {SANE_LO}~{SANE_HI})")
    if bad:
        ok = False

    # 相关性必须逐年算再取中位数: 成交额是金额、换手率是比率, 分母(市值)
    # 十几年里长了两个数量级 —— 跨全历史直接算相关没有意义
    # (Wind 原列自己跨 26 年也只有 0.169, 而年内有 0.72)。
    both = pd.DataFrame({'tr': vals, 'amt': amt, 'y': d[sel].dt.year}).dropna()
    per_year = both.groupby('y')[['tr', 'amt']].apply(
        lambda g: g['tr'].corr(g['amt']) if len(g) > 30 else np.nan).dropna()
    corr = float(per_year.median()) if len(per_year) else np.nan
    print(f"  与成交额相关     : 逐年中位数 {corr:.4f}  "
          f"(最差年份 {per_year.min():.4f})   (阈值 >= {MIN_CORR_AMT})")
    if not (corr >= MIN_CORR_AMT):
        ok = False

    denom = (both['amt'] / (both['tr'] / 100)).sort_index()
    move = denom.pct_change().abs().median()
    print(f"  隐含分母日环比   : 中位 {move * 100:.2f}%   (阈值 <= {MAX_DENOM_MOVE * 100:.0f}%)")
    if not (move <= MAX_DENOM_MOVE):
        ok = False

    old = pd.to_numeric(df.iloc[:, COL_TURNOVER], errors='coerce')[sel].dropna()
    print(f"\n  原列(Wind)       : 均值 {old.mean():.2f}%  范围 {old.min():.2f}~{old.max():.2f}")
    print(f"  新列(申万)       : 均值 {v.mean():.2f}%  范围 {v.min():.2f}~{v.max():.2f}")
    print(f"  水平比值         : {v.mean() / old.mean():.2f}x  (自由流通口径, 不影响分位数)")
    print(f"\n结论: {'通过 —— 可安全重建' if ok else '未通过 —— 拒绝写入'}")
    return ok, vals, sel


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--start', default=None, help='默认 Excel 最早一行')
    ap.add_argument('--end', default=None, help='默认 Excel 最后一行')
    ap.add_argument('--apply', action='store_true', help='校验通过后真正写入 Excel')
    ap.add_argument('--refetch', action='store_true', help='忽略缓存重新抓取')
    args = ap.parse_args()

    if not os.path.exists(EXCEL_PATH):
        print(f"[错误] 未找到 {EXCEL_PATH}")
        sys.exit(1)

    df = pd.read_excel(EXCEL_PATH)
    d = pd.to_datetime(df.iloc[:, COL_DATE], errors='coerce')
    start = args.start or d.min().strftime('%Y-%m-%d')
    end = args.end or d.max().strftime('%Y-%m-%d')

    print(f">> 抓取申万一级行业换手率 ({start} ~ {end})")
    tr_map = fetch_turnover(start, end, refetch=args.refetch)

    ok, vals, sel = validate(tr_map, df, start, end)

    idx = df.index[sel]
    print(f"\n重建预览 (前 5 行 / 后 5 行, 共 {len(idx)} 行):")
    old = pd.to_numeric(df.iloc[:, COL_TURNOVER], errors='coerce')
    show = list(idx[:5]) + list(idx[-5:])
    for i in show:
        nv = vals.get(i, np.nan)
        print(f"  {d[i]:%Y-%m-%d}  {old[i]:>6.2f}%  ->  "
              f"{'—' if pd.isna(nv) else f'{nv:6.2f}%'}")

    if not args.apply:
        print("\n[未写入] 这是校验模式。确认无误后加 --apply 执行写入。")
        return
    if not ok:
        print("\n[拒绝写入] 自检未通过。")
        sys.exit(2)

    n = 0
    for i in idx:
        nv = vals.get(i, np.nan)
        # 抓不到的交易日留空, 不保留旧口径的数字 —— 混口径比缺失更糟
        df.iat[i, COL_TURNOVER] = np.nan if pd.isna(nv) else float(nv)
        n += 1
    df.to_excel(EXCEL_PATH, index=False)
    print(f"\n[OK] 已重建换手率 {n} 行 (其中 {int(vals.notna().sum())} 行有值), 写入 {EXCEL_PATH}")
    print("     记得把 auto_fetch_daily.SW_TURNOVER_ENABLED 置为 True, 让日常管线续写同口径。")
    print("     接着运行 python update.py 重算分位数并发布。")


if __name__ == '__main__':
    main()
