# -*- coding: utf-8 -*-
"""
美股情绪指标数据抓取 —— 对应 A 股的 auto_fetch_daily.py

四个子指标 (与 A 股一一对应, 但口径按美股实际可得数据调整):
  1. 全市场成交额   <- 11 只 SPDR 板块 ETF 的成交额合计 (十亿美元)
  2. 板块集中度     <- 同一批 ETF 中成交额前 3 名占合计的比重 (%)
  3. 上涨个股占比   <- UNIVERSE 中当日收涨的比例 (%)
  4. VIX            <- ^VIX 收盘价 (计算分位时反向, 见 us_sentiment_indicator.py)

前两项来自同一次下载, 因此天然自洽 —— 与 A 股用一次申万调用同时得到
成交额集中度和换手率是同一个思路。

输出: us_market_data.csv (date, dollar_volume_b, top3_share_pct, rise_pct, vix)

设计原则 (沿用 A 股管线的教训):
- 缺数据一律留空, 绝不写占位常量
- 每次请求都有超时与重试, 不依赖库自带的默认值
- 增量抓取: 已有的日期不重复下载
"""

import os
import sys
import time
import contextlib

import numpy as np
import pandas as pd
import requests

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(BASE_DIR, 'us_market_data.csv')

FETCH_TIMEOUT = (10, 30)   # (连接, 读取) 秒
MAX_RETRY = 3

# SPDR 11 个 GICS 板块 ETF —— 完整切分标普 500 的行业结构
SECTOR_ETFS = ['XLK', 'XLF', 'XLV', 'XLY', 'XLP', 'XLE', 'XLI', 'XLB', 'XLRE', 'XLU', 'XLC']
MIN_SECTORS = 9            # 当日有效板块少于此数则该日留空

# 市场宽度的统计样本。默认用标普 100 成分 (大盘股), 可自行替换为更宽的列表;
# 注意这是大盘股口径, 与 A 股"全部成份股"的广度含义并不完全等同。
UNIVERSE = """
AAPL ABBV ABT ACN ADBE AIG AMD AMGN AMT AMZN AVGO AXP BA BAC BK BKNG BLK BMY BRK-B C
CAT CHTR CL CMCSA COF COP COST CRM CSCO CVS CVX DE DHR DIS DOW DUK EMR EXC F FDX GD
GE GILD GM GOOG GOOGL GS HD HON IBM INTC INTU JNJ JPM KHC KO LIN LLY LMT LOW MA MCD
MDLZ MDT MET META MMM MO MRK MS MSFT NEE NFLX NKE NVDA ORCL PEP PFE PG PM PYPL QCOM
RTX SBUX SCHW SO SPG T TGT TMO TMUS TSLA TXN UNH UNP UPS USB V VZ WFC WMT XOM
""".split()
MIN_UNIVERSE = 60          # 当日有效样本少于此数则宽度留空

VIX_SYMBOL = '^VIX'


@contextlib.contextmanager
def _request_timeout(timeout):
    """
    给底层那些不带 timeout 的 requests 调用注入默认超时。

    与 auto_fetch_daily._request_timeout 同理: socket.setdefaulttimeout 拦不住,
    因为 urllib3 会在连接后无条件执行 sock.settimeout(self.timeout), 而 requests
    未传 timeout 时该值为 None, 等于把 socket 设回永久阻塞。只有拦在 requests
    这一层才有效。退出时还原, 不外溢。
    """
    orig = requests.Session.request

    def patched(self, *args, **kwargs):
        if kwargs.get('timeout') is None:
            kwargs['timeout'] = timeout
        return orig(self, *args, **kwargs)

    requests.Session.request = patched
    try:
        yield
    finally:
        requests.Session.request = orig


def _download(tickers, start, end):
    """批量下载日线, 带重试。返回 yfinance 的多层列 DataFrame, 失败返回 None。"""
    import yfinance as yf

    last = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            with _request_timeout(FETCH_TIMEOUT):
                df = yf.download(tickers, start=start, end=end, progress=False,
                                 auto_adjust=False, threads=False)
            if df is not None and len(df) > 0:
                return df
            last = RuntimeError('返回空数据')
        except Exception as e:
            last = e
        if attempt < MAX_RETRY:
            wait = 2 ** attempt
            print(f"     [重试 {attempt}/{MAX_RETRY}] {type(last).__name__}: {last} —— {wait}s 后重试")
            time.sleep(wait)
    print(f"     [WARN] 下载失败: {last}")
    return None


def _field(df, name, tickers):
    """从 yfinance 的多层列里取出某个字段, 返回 DataFrame[日期 x 代码]。"""
    if isinstance(df.columns, pd.MultiIndex):
        if name not in df.columns.get_level_values(0):
            return pd.DataFrame()
        sub = df[name]
    else:
        # 单只代码时 yfinance 返回单层列
        if name not in df.columns:
            return pd.DataFrame()
        sub = df[[name]]
        sub.columns = [tickers[0] if isinstance(tickers, list) else tickers]
    return sub


def fetch_sector_metrics(start, end):
    """板块 ETF -> 每日成交额合计(十亿美元) 与 前3占比(%)。"""
    print(f">> [1/3] 下载 {len(SECTOR_ETFS)} 只板块 ETF ({start} ~ {end}) ...")
    df = _download(SECTOR_ETFS, start, end)
    if df is None:
        return pd.DataFrame()

    close, vol = _field(df, 'Close', SECTOR_ETFS), _field(df, 'Volume', SECTOR_ETFS)
    if close.empty or vol.empty:
        print("     [WARN] 缺少 Close/Volume 字段")
        return pd.DataFrame()

    dollar = (close * vol).dropna(how='all')          # 成交额 = 收盘价 x 成交量
    rows = []
    for d, r in dollar.iterrows():
        vals = r.dropna()
        vals = vals[vals > 0]
        if len(vals) < MIN_SECTORS:
            continue
        total = float(vals.sum())
        top3 = float(vals.sort_values(ascending=False).head(3).sum())
        rows.append({'date': pd.Timestamp(d).strftime('%Y-%m-%d'),
                     'dollar_volume_b': total / 1e9,
                     'top3_share_pct': top3 / total * 100})
    print(f"     [OK] {len(rows)} 个交易日")
    return pd.DataFrame(rows)


def fetch_breadth(start, end):
    """UNIVERSE -> 每日收涨比例(%)。"""
    print(f">> [2/3] 下载 {len(UNIVERSE)} 只成分股用于市场宽度 ...")
    df = _download(UNIVERSE, start, end)
    if df is None:
        return pd.DataFrame()

    close = _field(df, 'Close', UNIVERSE)
    if close.empty:
        print("     [WARN] 缺少 Close 字段")
        return pd.DataFrame()

    chg = close.pct_change()
    rows = []
    for d, r in chg.iterrows():
        vals = r.dropna()
        if len(vals) < MIN_UNIVERSE:
            continue
        rows.append({'date': pd.Timestamp(d).strftime('%Y-%m-%d'),
                     'rise_pct': float((vals > 0).sum()) / len(vals) * 100})
    print(f"     [OK] {len(rows)} 个交易日")
    return pd.DataFrame(rows)


def fetch_vix(start, end):
    """^VIX 收盘价。"""
    print(">> [3/3] 下载 VIX ...")
    df = _download(VIX_SYMBOL, start, end)
    if df is None:
        return pd.DataFrame()
    close = _field(df, 'Close', [VIX_SYMBOL])
    if close.empty:
        print("     [WARN] 缺少 Close 字段")
        return pd.DataFrame()
    ser = close.iloc[:, 0].dropna()
    print(f"     [OK] {len(ser)} 个交易日")
    return pd.DataFrame({'date': [pd.Timestamp(d).strftime('%Y-%m-%d') for d in ser.index],
                         'vix': ser.values})


def main():
    # 滚动 252 日分位需要至少两年历史才有意义, 首次抓取多取一些
    default_start = (pd.Timestamp.today() - pd.Timedelta(days=365 * 4)).strftime('%Y-%m-%d')
    start = sys.argv[1] if len(sys.argv) > 1 else default_start
    end = (pd.Timestamp.today() + pd.Timedelta(days=1)).strftime('%Y-%m-%d')

    print("=" * 56)
    print(">> 美股情绪指标数据抓取")
    print("=" * 56)

    sectors, breadth, vix = fetch_sector_metrics(start, end), fetch_breadth(start, end), fetch_vix(start, end)
    if sectors.empty and breadth.empty and vix.empty:
        print("\n[错误] 三项数据全部抓取失败, 不写入任何内容")
        sys.exit(1)

    out = None
    for part in (sectors, breadth, vix):
        if part.empty:
            continue
        out = part if out is None else out.merge(part, on='date', how='outer')
    out = out.sort_values('date').reset_index(drop=True)

    # 补齐缺失列, 保持结构稳定 (缺的就是缺的, 留空)
    for c in ['dollar_volume_b', 'top3_share_pct', 'rise_pct', 'vix']:
        if c not in out.columns:
            out[c] = np.nan
    out = out[['date', 'dollar_volume_b', 'top3_share_pct', 'rise_pct', 'vix']]

    # 与已有数据合并: 新抓到的值覆盖旧值, 旧有而本次没抓到的保留
    if os.path.exists(CSV_PATH):
        old = pd.read_csv(CSV_PATH)
        merged = pd.concat([old, out], ignore_index=True)
        merged = merged.groupby('date', as_index=False).last()
        out = merged.sort_values('date').reset_index(drop=True)

    out.to_csv(CSV_PATH, index=False)
    print(f"\n[OK] 写入 {CSV_PATH}: {len(out)} 个交易日 ({out['date'].iloc[0]} ~ {out['date'].iloc[-1]})")
    for c in ['dollar_volume_b', 'top3_share_pct', 'rise_pct', 'vix']:
        n = out[c].notna().sum()
        print(f"     {c:18s} 有效 {n:5d} / {len(out)}")
    print("\n下一步: python us_sentiment_indicator.py")


if __name__ == '__main__':
    main()
