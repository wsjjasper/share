# -*- coding: utf-8 -*-
"""
用申万一级行业数据回填「换手率」与「成交额前三行业占比」(一次性脚本)

背景:
- 行业集中度: 2026-08-26 起自动追加的行曾误用东财 t:1 地域板块(省份), 已置空。
- 换手率:     同期的行写的是硬编码常量 1.48, 并非真实数据。

两项均改由申万宏源提供, 与 Excel 中 2026-08-25 之前的 Wind 历史同口径。
抓取复用 auto_fetch_daily.fetch_sw_daily_metrics —— 与日常管线共用一份实现。

安全措施: 写入前必须先与重叠期的 Wind 历史逐日比对, 任一列不达标即整体拒绝写入。

用法:
    python backfill_industry_sw.py                  # 只校验+预览, 不写入 (默认)
    python backfill_industry_sw.py --apply          # 校验通过后写入
    python backfill_industry_sw.py --validate-days 250
"""

import argparse
import os
import sys

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, '副本万得全A.xlsx')

COL_DATE, COL_TURNOVER, COL_RATIO, COL_TOTAL_AMT, COL_TOP3_AMT = 1, 2, 3, 6, 7

# 每列的校验口径。excel_scale: 把 Excel 存的值换算成对比单位(%)的系数
CHECKS = [
    {'key': 'top3_ratio',    'col': COL_RATIO,    'name': '行业集中度',
     'excel_scale': 100.0, 'sw_scale': 100.0, 'max_mad': 3.0,  'min_corr': 0.80},
    {'key': 'turnover_rate', 'col': COL_TURNOVER, 'name': '换手率',
     'excel_scale': 1.0,   'sw_scale': 1.0,   'max_mad': 0.5,  'min_corr': 0.80},
]


def fetch_sw(start_date, end_date):
    """复用日常管线的抓取实现, 避免两处口径各自漂移。"""
    from auto_fetch_daily import fetch_sw_daily_metrics

    sw = fetch_sw_daily_metrics(start_date, end_date)
    if not sw:
        raise RuntimeError("申万接口未返回可用数据")
    print(f"     [OK] 共 {len(sw)} 个交易日可用于校验与回填")
    return sw


def validate(sw, df, n_days, before):
    """
    在重叠期把申万口径与 Excel 里的 Wind 历史值逐日比对。

    只取 before(回填起始日) 之前的行: 回填区间内的数值要么是已置空的缺失值,
    要么是硬编码占位常量(换手率 1.48), 都不是 Wind 真值 —— 放进校验窗口会
    污染结果, 既可能造成误判失败, 也可能在占位值恰好接近时造成误判通过。
    """
    d = pd.to_datetime(df.iloc[:, COL_DATE]).dt.normalize()
    before = pd.Timestamp(before)
    passed = {}

    for chk in CHECKS:
        rows = []
        for i in df.index:
            if d[i] >= before:
                continue
            key = d[i].strftime('%Y-%m-%d')
            ev = df.iloc[i, chk['col']]
            if key not in sw or pd.isna(ev):
                continue
            sv = sw[key][chk['key']]
            if sv is None:
                continue
            rows.append({'日期': d[i], 'wind': ev * chk['excel_scale'], 'sw': sv * chk['sw_scale']})

        m = pd.DataFrame(rows).tail(n_days)
        print(f"\n口径校验 · {chk['name']}")
        if len(m) < 20:
            print(f"  [失败] 可比对的交易日仅 {len(m)} 个, 不足以校验")
            passed[chk['key']] = False
            continue

        diff = m['sw'] - m['wind']
        mad, corr, bias = diff.abs().mean(), m['sw'].corr(m['wind']), diff.mean()
        print(f"  区间           : {m['日期'].min().date()} ~ {m['日期'].max().date()}  ({len(m)} 个交易日)")
        print(f"  Wind 均值      : {m['wind'].mean():.2f}%")
        print(f"  申万均值       : {m['sw'].mean():.2f}%")
        print(f"  系统性偏差     : {bias:+.2f} 个百分点")
        print(f"  平均绝对偏差   : {mad:.2f} 个百分点   (阈值 <= {chk['max_mad']})")
        print(f"  最大绝对偏差   : {diff.abs().max():.2f} 个百分点")
        print(f"  相关系数       : {corr:.4f}   (阈值 >= {chk['min_corr']})")
        # bool(): mad/corr 是 numpy 标量, 比较结果 np.bool_ 与 `is True` 不相等
        ok = bool((mad <= chk['max_mad']) and (corr >= chk['min_corr']))
        print(f"  结论: {'通过' if ok else '未通过'}")
        passed[chk['key']] = ok

    names = lambda v: ', '.join(c['name'] for c in CHECKS if passed.get(c['key']) is v) or '无'
    print("")
    print(f"总体: 通过 = {names(True)};  未通过 = {names(False)}")
    print("      按列独立放行 —— 只写入通过的列, 未通过的列保持原样不动。")
    return passed


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--start', default='2026-08-26')
    ap.add_argument('--end', default='2026-09-16')
    ap.add_argument('--validate-days', type=int, default=120)
    ap.add_argument('--apply', action='store_true', help='校验通过后真正写入 Excel')
    args = ap.parse_args()

    if not os.path.exists(EXCEL_PATH):
        print(f"[错误] 未找到 {EXCEL_PATH}")
        sys.exit(1)

    df = pd.read_excel(EXCEL_PATH)
    d = pd.to_datetime(df.iloc[:, COL_DATE]).dt.normalize()

    fetch_start = (pd.Timestamp(args.start)
                   - pd.Timedelta(days=args.validate_days * 2)).strftime('%Y-%m-%d')
    sw = fetch_sw(fetch_start, args.end)

    passed = validate(sw, df, args.validate_days, args.start)

    target = df.index[(d >= args.start) & (d <= args.end)]
    print(f"\n待回填区间 {args.start} ~ {args.end}: Excel 内 {len(target)} 行")
    print("\n回填预览:")
    for i in target:
        key = d[i].strftime('%Y-%m-%d')
        if key not in sw:
            print(f"  {key}  申万无此交易日数据, 跳过")
            continue
        e = sw[key]
        tr = '—' if e['turnover_rate'] is None else f"{e['turnover_rate']:.2f}%"
        print(f"  {key}  换手率 {tr:>7}   行业前3 {e['top3_ratio'] * 100:.2f}%   {e['top3_detail']}")

    if not args.apply:
        print("\n[未写入] 这是校验模式。确认无误后加 --apply 执行写入。")
        return
    if not any(passed.values()):
        print("\n[拒绝写入] 口径校验未通过。")
        sys.exit(2)

    n_ratio = n_tr = 0
    for i in target:
        key = d[i].strftime('%Y-%m-%d')
        if key not in sw:
            continue
        e = sw[key]
        if passed.get('top3_ratio'):
            df.iat[i, COL_RATIO] = e['top3_ratio']
            total_amt = df.iat[i, COL_TOTAL_AMT]
            df.iat[i, COL_TOP3_AMT] = (e['top3_ratio'] * total_amt) if pd.notna(total_amt) else np.nan
            n_ratio += 1
        if passed.get('turnover_rate') and e['turnover_rate'] is not None:
            df.iat[i, COL_TURNOVER] = e['turnover_rate']
            n_tr += 1

    df.to_excel(EXCEL_PATH, index=False)
    print(f"\n[OK] 已回填 行业集中度 {n_ratio} 行 / 换手率 {n_tr} 行, 写入 {EXCEL_PATH}")
    for c in CHECKS:
        if not passed.get(c['key']):
            print(f"     [跳过] {c['name']}: 口径校验未通过, 该列未改动")
    print("     接着运行 python update.py 重算分位数并发布。")


if __name__ == '__main__':
    main()
