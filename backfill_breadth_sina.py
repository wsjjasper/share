# -*- coding: utf-8 -*-
"""
用逐只个股日线重建「上涨个股占比」(一次性脚本)

背景:
    这一列 2026-08-26 起的 16 行是硬编码占位常量 0.585 (58.50%)。它之前的真实
    数据在 8.05% ~ 93.88% 之间剧烈波动 —— 市场宽度不可能连续 16 天纹丝不动。
    分位被钉在 64~65, 把综合情绪值抬高了约 9~12 分。

为什么要逐只算:
    没有任何免费接口按日期返回全市场涨跌家数。已逐一排除:
      · 交易所官方汇总 (stock_sse_deal_daily / stock_szse_summary) —— 只有成交额、
        市值、换手率, 没有涨跌家数
      · akshare 全部带 date 参数的函数 —— 只有涨停板池 (stock_zt_pool_em), 口径不同
      · 乐咕赚钱效应 (stock_market_activity_legu) —— 是快照且页面结构已变, 接口报错
      · 东财 stock_zh_a_spot_em —— 快照, 无历史
    唯一可靠的办法是把全市场个股日线拉下来自己数。个股日线是按区间返回的,
    所以请求数只跟股票数有关 (约 5500 次), 跟需要多少个交易日无关。

数据源:
    股票清单  新浪 Market_Center.getHQNodeData  (含北交所, 约 5560 只)
    日线      新浪 CN_MarketData.getKLineData   (实测 0.28s/只; 东财会限流封连接)
    上涨判定  当日收盘 > 前一日收盘 (同一序列内相邻两根K线)

用法:
    python backfill_breadth_sina.py                 # 抓取+校验+预览, 不写入
    python backfill_breadth_sina.py --apply         # 校验通过后写入
    python backfill_breadth_sina.py --refetch       # 忽略缓存重抓
    python backfill_breadth_sina.py --limit 800     # 只抽前 N 只 (快速试跑)
"""

import argparse
import json
import os
import sys
import time

import numpy as np
import pandas as pd
import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, '副本万得全A.xlsx')
CACHE_PATH = os.path.join(BASE_DIR, '_breadth_cache.tmp')     # *.tmp 已被 .gitignore 忽略

COL_DATE, COL_RISE_DEC, COL_UP_CNT, COL_MEMBERS, COL_RISE_PCT = 1, 4, 8, 9, 11

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
LIST_URL = ("http://vip.stock.finance.sina.com.cn/quotes_service/api/json_v2.php/"
            "Market_Center.getHQNodeData")
KLINE_URL = ("http://money.finance.sina.com.cn/quotes_service/api/json_v2.php/"
             "CN_MarketData.getKLineData")
DATALEN = 60             # 每只取最近 N 根日线。请求数只跟股票数有关, 与这个值无关;
                         # 取 60 是为了在回填区间之外还留约 44 个交易日与 Wind 对账
SAVE_EVERY = 250         # 每处理这么多只存一次盘, 断了可以续
MIN_STOCKS_PER_DAY = 3000   # 某日参与统计的个股数低于此值则该日不可信

# 校验门槛: 与 Excel 里 2026-08-26 之前的 Wind 真实宽度逐日比对
MAX_MAD = 3.0            # 平均绝对偏差 (个百分点)
MIN_CORR = 0.95          # 相关系数; 宽度波动极大, 真同源应该非常高


def _session():
    s = requests.Session()
    s.headers.update({'User-Agent': UA, 'Referer': 'http://finance.sina.com.cn/'})
    return s


def fetch_symbols(s):
    """新浪全 A 清单 (hs_a 含沪深, 另含北交所 bj 前缀)。"""
    codes, page = [], 1
    while page <= 100:
        r = s.get(LIST_URL, params={'page': page, 'num': 100, 'sort': 'symbol',
                                    'asc': 1, 'node': 'hs_a'}, timeout=(8, 20))
        try:
            arr = json.loads(r.text)
        except Exception:
            break
        if not arr:
            break
        codes += [x['symbol'] for x in arr]
        page += 1
    return codes


def fetch_bars(s, sym, retries=3):
    """取一只个股最近 DATALEN 根日线; 失败返回 None。"""
    for a in range(retries):
        try:
            r = s.get(KLINE_URL, params={'symbol': sym, 'scale': 240, 'ma': 'no',
                                         'datalen': DATALEN}, timeout=(8, 20))
            txt = r.text.strip()
            if not txt or txt == 'null':
                return []
            return json.loads(txt)
        except Exception:
            if a == retries - 1:
                return None
            time.sleep(0.6 * (a + 1))
    return None


def collect(limit=None, refetch=False):
    """
    逐只累加每个交易日的 (上涨数, 参与数)。

    不保存逐只K线 —— 只累加计数器, 内存和缓存都很小; 另存已处理清单以便断点续跑。
    """
    state = {'up': {}, 'tot': {}, 'done': []}
    if not refetch and os.path.exists(CACHE_PATH):
        state = json.load(open(CACHE_PATH, encoding='utf-8'))
        print(f"     [缓存] 已处理 {len(state['done'])} 只; 加 --refetch 可重抓")

    s = _session()
    print(">> 获取股票清单...")
    syms = fetch_symbols(s)
    if not syms:
        raise RuntimeError('股票清单获取失败')
    if limit:
        syms = syms[:limit]
    print(f"     [OK] {len(syms)} 只")

    done = set(state['done'])
    todo = [x for x in syms if x not in done]
    print(f">> 需要抓取 {len(todo)} 只 (已完成 {len(done)})")
    t0, failed = time.time(), 0
    for i, sym in enumerate(todo, 1):
        bars = fetch_bars(s, sym)
        if bars is None:
            failed += 1
        elif bars:
            prev = None
            for b in bars:
                try:
                    c = float(b['close'])
                except (TypeError, ValueError, KeyError):
                    prev = None
                    continue
                day = b['day'][:10]
                if prev is not None and prev > 0:
                    state['tot'][day] = state['tot'].get(day, 0) + 1
                    if c > prev:
                        state['up'][day] = state['up'].get(day, 0) + 1
                prev = c
        state['done'].append(sym)
        if i % SAVE_EVERY == 0:
            json.dump(state, open(CACHE_PATH, 'w', encoding='utf-8'))
            el = time.time() - t0
            print(f"     [{i}/{len(todo)}] {el/60:.1f} 分钟, 失败 {failed}, "
                  f"预计还需 {(len(todo)-i)*el/i/60:.1f} 分钟", flush=True)
    json.dump(state, open(CACHE_PATH, 'w', encoding='utf-8'))
    print(f"     [OK] 完成, 失败 {failed} 只, 覆盖 {len(state['tot'])} 个交易日")
    return state


def build_series(state):
    """把计数器整理成 {日期: 上涨占比%}, 剔除参与数过少的日子。"""
    out = {}
    for day, tot in state['tot'].items():
        if tot < MIN_STOCKS_PER_DAY:
            continue
        out[day] = state['up'].get(day, 0) / tot * 100
    return out, {d: state['tot'][d] for d in out}


def validate(series, totals, df, start):
    d = pd.to_datetime(df.iloc[:, COL_DATE], errors='coerce').dt.normalize()
    wind = pd.to_numeric(df.iloc[:, COL_RISE_DEC], errors='coerce') * 100
    rows = []
    for i in df.index:
        if pd.isna(d[i]) or d[i] >= pd.Timestamp(start):
            continue
            # 校验窗口只取 --start 之前的行。回填区间里的值本就是坏的
            # (占位常量, 或同一个快照被写给多个日期), 放进来只会把
            # MAD 抬高、把相关压低, 把一个合格的重建误判成不合格。
        k = d[i].strftime('%Y-%m-%d')
        if k in series and pd.notna(wind[i]):
            rows.append({'日期': d[i], 'wind': wind[i], 'new': series[k], 'n': totals[k]})
    m = pd.DataFrame(rows)
    print(f"\n口径校验 · 上涨个股占比")
    if len(m) < 20:
        print(f"  [失败] 可比对交易日仅 {len(m)} 个, 不足以校验")
        return False, m
    diff = m['new'] - m['wind']
    mad, corr = diff.abs().mean(), m['new'].corr(m['wind'])
    print(f"  区间           : {m['日期'].min():%Y-%m-%d} ~ {m['日期'].max():%Y-%m-%d}  "
          f"({len(m)} 个交易日)")
    print(f"  参与个股数     : 中位 {int(m['n'].median())}")
    print(f"  Wind 均值      : {m['wind'].mean():.2f}%   区间 {m['wind'].min():.2f}~{m['wind'].max():.2f}")
    print(f"  重建均值       : {m['new'].mean():.2f}%   区间 {m['new'].min():.2f}~{m['new'].max():.2f}")
    print(f"  系统性偏差     : {diff.mean():+.2f} 个百分点")
    print(f"  平均绝对偏差   : {mad:.2f} 个百分点   (阈值 <= {MAX_MAD})")
    print(f"  最大绝对偏差   : {diff.abs().max():.2f} 个百分点")
    print(f"  相关系数       : {corr:.4f}   (阈值 >= {MIN_CORR})")
    ok = bool(mad <= MAX_MAD and corr >= MIN_CORR)
    print(f"  结论: {'通过' if ok else '未通过'}")
    return ok, m


def main():
    ap = argparse.ArgumentParser()
    # 2026-08-24 与 08-21 的上涨占比/家数完全相同, 是同一快照写了两次;
    # 08-25 也是管线追加的 —— 所以回填从 08-24 起, 而不是 08-26
    ap.add_argument('--start', default='2026-08-24')
    ap.add_argument('--end', default='2026-09-16')
    ap.add_argument('--apply', action='store_true')
    ap.add_argument('--refetch', action='store_true')
    ap.add_argument('--limit', type=int, default=None)
    ap.add_argument('--min-stocks', type=int, default=None,
                    help='覆盖 MIN_STOCKS_PER_DAY; 只用于小样本试跑')
    args = ap.parse_args()

    if not os.path.exists(EXCEL_PATH):
        print(f"[错误] 未找到 {EXCEL_PATH}")
        sys.exit(1)

    if args.min_stocks:
        global MIN_STOCKS_PER_DAY
        MIN_STOCKS_PER_DAY = args.min_stocks
    state = collect(limit=args.limit, refetch=args.refetch)
    series, totals = build_series(state)

    df = pd.read_excel(EXCEL_PATH)
    d = pd.to_datetime(df.iloc[:, COL_DATE], errors='coerce').dt.normalize()
    ok, _ = validate(series, totals, df, args.start)

    target = df.index[(d >= args.start) & (d <= args.end)]
    print(f"\n待回填区间 {args.start} ~ {args.end}: Excel 内 {len(target)} 行")
    print("回填预览:")
    for i in target:
        k = d[i].strftime('%Y-%m-%d')
        cur = pd.to_numeric(df.iloc[:, COL_RISE_DEC], errors='coerce')[i] * 100
        if k in series:
            print(f"  {k}  {cur:5.2f}%  ->  {series[k]:5.2f}%   (参与 {totals[k]} 只)")
        else:
            print(f"  {k}  {cur:5.2f}%  ->  无数据, 跳过")

    if not args.apply:
        print("\n[未写入] 这是校验模式。确认无误后加 --apply 执行写入。")
        return
    if not ok:
        print("\n[拒绝写入] 口径校验未通过。")
        sys.exit(2)

    n = 0
    for i in target:
        k = d[i].strftime('%Y-%m-%d')
        if k not in series:
            continue
        ratio = series[k] / 100.0
        df.iat[i, COL_RISE_DEC] = ratio                      # 小数
        df.iat[i, COL_RISE_PCT] = series[k]                  # 百分数
        df.iat[i, COL_UP_CNT] = int(round(ratio * totals[k]))  # 上涨家数
        df.iat[i, COL_MEMBERS] = totals[k]                   # 参与个股数
        n += 1
    df.to_excel(EXCEL_PATH, index=False)
    print(f"\n[OK] 已回填上涨个股占比 {n} 行, 写入 {EXCEL_PATH}")
    print("     接着运行 python update.py 重算分位数并发布。")


if __name__ == '__main__':
    main()
