# -*- coding: utf-8 -*-
"""
用申万一级行业数据回填「成交额前三行业占比」(一次性脚本)

背景: 2026-08-26 起自动追加的行, 该列曾误用东财 t:1 地域板块(省份), 现已置空。
本脚本改用申万宏源的一级行业口径 —— 与 Excel 里 2026-08-25 之前的 Wind
历史数据同口径(申万一级 31 个行业), 因此不会引入量纲断层。

数据源: akshare.index_analysis_daily_sw(symbol="一级行业")
        每行含 指数代码/指数名称/发布日期/成交额占比, 前3名占比相加即为目标值。

用法:
    python backfill_industry_sw.py                 # 只校验, 不写入 (默认)
    python backfill_industry_sw.py --apply         # 校验通过后写入 Excel
    python backfill_industry_sw.py --validate-days 250
"""

import argparse
import os
import sys

import numpy as np
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, '副本万得全A.xlsx')

COL_DATE, COL_RATIO, COL_TOTAL_AMT, COL_TOP3_AMT = 1, 3, 6, 7

# 校验阈值: 与 Wind 历史数据的一致性要求
MAX_MEAN_ABS_DIFF = 3.0   # 平均绝对偏差上限 (百分点)
MIN_CORRELATION = 0.80    # 相关系数下限


def fetch_sw_top3(start_date, end_date):
    """
    返回 DataFrame[日期, top3_ratio, 明细]

    抓取与单位判定复用 auto_fetch_daily.fetch_sw_industry_top3 —— 两处共用一份实现,
    避免日常管线与回填脚本的口径各自漂移。
    """
    from auto_fetch_daily import fetch_sw_industry_top3

    sw = fetch_sw_industry_top3(start_date, end_date)
    if not sw:
        raise RuntimeError("申万接口未返回可用数据")

    out = pd.DataFrame(
        [{'日期': pd.Timestamp(k), 'top3_ratio': v[0], '明细': v[1]} for k, v in sw.items()]
    ).sort_values('日期').reset_index(drop=True)
    print(f"     [OK] 共 {len(out)} 个交易日可用于校验与回填")
    return out


def validate(sw, df_excel, n_days):
    """在重叠期把申万口径与 Excel 里的 Wind 历史值逐日比对。"""
    d = pd.to_datetime(df_excel.iloc[:, COL_DATE]).dt.normalize()
    wind = pd.DataFrame({'日期': d, 'wind': df_excel.iloc[:, COL_RATIO]}).dropna()
    merged = wind.merge(sw[['日期', 'top3_ratio']], on='日期', how='inner')
    merged = merged.tail(n_days)

    if len(merged) < 20:
        print(f"[失败] 重叠交易日仅 {len(merged)} 个, 不足以校验口径")
        return False, merged

    merged['diff'] = (merged['top3_ratio'] - merged['wind']) * 100
    mean_abs = merged['diff'].abs().mean()
    corr = merged['top3_ratio'].corr(merged['wind'])
    bias = merged['diff'].mean()

    print(f"\n口径校验 (重叠 {len(merged)} 个交易日, {merged['日期'].min().date()} ~ {merged['日期'].max().date()})")
    print(f"  Wind 均值      : {merged['wind'].mean() * 100:.2f}%")
    print(f"  申万均值       : {merged['top3_ratio'].mean() * 100:.2f}%")
    print(f"  系统性偏差     : {bias:+.2f} 个百分点")
    print(f"  平均绝对偏差   : {mean_abs:.2f} 个百分点   (阈值 <= {MAX_MEAN_ABS_DIFF})")
    print(f"  最大绝对偏差   : {merged['diff'].abs().max():.2f} 个百分点")
    print(f"  相关系数       : {corr:.4f}   (阈值 >= {MIN_CORRELATION})")

    ok = (mean_abs <= MAX_MEAN_ABS_DIFF) and (corr >= MIN_CORRELATION)
    print(f"  结论: {'通过 —— 申万与 Wind 同口径, 可安全回填' if ok else '未通过 —— 口径不一致, 拒绝写入'}")
    return ok, merged


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--start', default='2026-08-26', help='回填起始日 (含)')
    ap.add_argument('--end', default='2026-09-16', help='回填结束日 (含)')
    ap.add_argument('--validate-days', type=int, default=120, help='用多少个重叠交易日校验口径')
    ap.add_argument('--apply', action='store_true', help='校验通过后真正写入 Excel')
    args = ap.parse_args()

    if not os.path.exists(EXCEL_PATH):
        print(f"[错误] 未找到 {EXCEL_PATH}")
        sys.exit(1)

    df = pd.read_excel(EXCEL_PATH)
    d = pd.to_datetime(df.iloc[:, COL_DATE]).dt.normalize()

    # 拉取范围要同时覆盖回填区间与校验区间
    fetch_start = (pd.Timestamp(args.start) - pd.Timedelta(days=args.validate_days * 2)).strftime('%Y-%m-%d')
    sw = fetch_sw_top3(fetch_start, args.end)

    ok, _ = validate(sw, df, args.validate_days)

    target = (d >= args.start) & (d <= args.end)
    print(f"\n待回填区间 {args.start} ~ {args.end}: Excel 内 {target.sum()} 行, "
          f"其中当前为空 {df.loc[target, df.columns[COL_RATIO]].isna().sum()} 行")

    sw_map = sw.set_index('日期')
    preview = []
    for i in df.index[target]:
        day = d[i]
        if day in sw_map.index:
            r = sw_map.loc[day]
            preview.append((day.strftime('%Y-%m-%d'), r['top3_ratio'], r['明细']))
        else:
            preview.append((day.strftime('%Y-%m-%d'), None, '申万无此交易日数据'))

    print("\n回填预览:")
    for day, ratio, detail in preview:
        print(f"  {day}  {'—' if ratio is None else '%.2f%%' % (ratio * 100)}   {detail}")

    if not args.apply:
        print("\n[未写入] 这是校验模式。确认无误后加 --apply 执行写入。")
        return

    if not ok:
        print("\n[拒绝写入] 口径校验未通过。")
        sys.exit(2)

    n = 0
    for i in df.index[target]:
        day = d[i]
        if day not in sw_map.index:
            continue
        ratio = float(sw_map.loc[day, 'top3_ratio'])
        df.iat[i, COL_RATIO] = ratio
        total_amt = df.iat[i, COL_TOTAL_AMT]
        # 第7列(前三行业合计,亿元)由占比 × 当日总成交额导出
        df.iat[i, COL_TOP3_AMT] = ratio * total_amt if pd.notna(total_amt) else np.nan
        n += 1

    df.to_excel(EXCEL_PATH, index=False)
    print(f"\n[OK] 已回填 {n} 行并写入 {EXCEL_PATH}")
    print("     接着运行 python update.py 重算分位数并发布。")


if __name__ == '__main__':
    main()
