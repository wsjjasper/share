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

# 确保控制台输出 UTF-8, 避免 Windows GBK 下打印 • 等字符直接抛 UnicodeEncodeError
# (A 股的 auto_fetch_daily.py / update.py 都有这段, 这两个脚本原来漏了)
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

FETCH_TIMEOUT = (10, 30)   # (连接, 读取) 秒
MAX_RETRY = 3

# 市场样本: 标普 100 成分 (大盘股)。成交额、板块集中度、上涨占比三项全部由这一份
# 数据算出 —— 与 A 股用一次申万调用同时得到集中度与换手率是同一思路: 同源即自洽。
#
# 早先板块集中度用的是 11 只 SPDR 板块 ETF 的成交额。那个口径有两个毛病:
#   1. 只有 11 个分母, 前 3 占比数学下限就是 3/11 = 27.3%, 879 天全部挤在
#      35.5~60.1% 之间(标准差 3.47pp)。这么窄的分布喂进滚动分位, 分位就是在放大
#      噪声 —— 101/879 天分位单日跳动超过 50 分。
#   2. ETF 自身成交量反映的是机构对冲与申赎流, 不是板块内个股的活跃程度。
# 改成按 GICS 板块汇总成分股成交额后, 衡量的是真实个股交易的集中度。
UNIVERSE_SECTORS = {
    '信息技术': ['AAPL', 'ACN', 'ADBE', 'AMD', 'AVGO', 'CRM', 'CSCO', 'IBM', 'INTC',
                 'INTU', 'MSFT', 'NVDA', 'ORCL', 'QCOM', 'TXN'],
    # 2023 年 GICS 调整后 V / MA / PYPL 由信息技术划入金融
    '金融': ['AIG', 'AXP', 'BAC', 'BLK', 'BRK-B', 'C', 'COF', 'GS', 'JPM', 'MA',
             'MET', 'MS', 'PYPL', 'SCHW', 'USB', 'V', 'WFC'],
    '医疗保健': ['ABBV', 'ABT', 'AMGN', 'BMY', 'CVS', 'DHR', 'GILD', 'JNJ', 'LLY',
                 'MDT', 'MRK', 'PFE', 'TMO', 'UNH'],
    '可选消费': ['AMZN', 'BKNG', 'F', 'GM', 'HD', 'LOW', 'MCD', 'NKE', 'SBUX', 'TGT', 'TSLA'],
    '日常消费': ['CL', 'COST', 'KHC', 'KO', 'MDLZ', 'MO', 'PEP', 'PG', 'PM', 'WMT'],
    '能源': ['COP', 'CVX', 'XOM'],
    '工业': ['BA', 'CAT', 'DE', 'EMR', 'FDX', 'GD', 'GE', 'HON', 'LMT', 'MMM',
             'RTX', 'UNP', 'UPS'],
    '原材料': ['DOW', 'LIN'],
    '房地产': ['AMT', 'SPG'],
    '公用事业': ['DUK', 'EXC', 'NEE', 'SO'],
    '通信服务': ['CHTR', 'CMCSA', 'DIS', 'GOOG', 'GOOGL', 'META', 'NFLX', 'T', 'TMUS'],
}
SECTOR_OF = {t: sec for sec, tics in UNIVERSE_SECTORS.items() for t in tics}
UNIVERSE = sorted(SECTOR_OF)

MIN_SECTORS = 9            # 当日有效板块少于此数则集中度留空 (共 11 个)
MIN_UNIVERSE = 60          # 当日有效样本少于此数则该项留空

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


def fetch_equity_metrics(start, end):
    """
    一次下载 UNIVERSE 的收盘价与成交量, 同时得出三项指标。

    同源的好处与 A 股一致: 成交额与集中度出自同一份数据, 不会互相打架。

      dollar_volume_b  成分股成交额合计 (十亿美元)
      top3_share_pct   按 GICS 板块汇总后, 成交额前 3 个板块占合计的比重 (%)
      rise_pct         当日收涨比例 (%)

    缺数据一律留空, 不写占位值。
    """
    print(f">> [1/2] 下载 {len(UNIVERSE)} 只成分股 ({start} ~ {end}) ...")
    df = _download(UNIVERSE, start, end)
    if df is None:
        return pd.DataFrame()

    close, vol = _field(df, 'Close', UNIVERSE), _field(df, 'Volume', UNIVERSE)
    if close.empty:
        print("     [WARN] 缺少 Close 字段")
        return pd.DataFrame()

    dollar = (close * vol) if not vol.empty else pd.DataFrame()
    # fill_method=None: 默认的 'pad' 会把停牌日前向填充成前一日收盘, 涨跌幅变成 0 ——
    # 该股仍计入分母却算作"没涨", 系统性压低宽度。置 None 后停牌日得到 NaN, 被一并排除。
    chg = close.pct_change(fill_method=None)

    rows = []
    for d in close.index:
        rec = {'date': pd.Timestamp(d).strftime('%Y-%m-%d')}

        if not dollar.empty:
            dv = dollar.loc[d].dropna()
            dv = dv[dv > 0]
            if len(dv) >= MIN_UNIVERSE:
                rec['dollar_volume_b'] = float(dv.sum()) / 1e9
                by_sector = {}
                for tic, val in dv.items():
                    sec = SECTOR_OF.get(tic)
                    if sec:
                        by_sector[sec] = by_sector.get(sec, 0.0) + float(val)
                total = sum(by_sector.values())
                if len(by_sector) >= MIN_SECTORS and total > 0:
                    top3 = sum(sorted(by_sector.values(), reverse=True)[:3])
                    rec['top3_share_pct'] = top3 / total * 100

        c = chg.loc[d].dropna()
        if len(c) >= MIN_UNIVERSE:
            rec['rise_pct'] = float((c > 0).sum()) / len(c) * 100

        if len(rec) > 1:          # 只有日期就不要这一行
            rows.append(rec)

    out = pd.DataFrame(rows)
    got = {c: int(out[c].notna().sum()) for c in
           ('dollar_volume_b', 'top3_share_pct', 'rise_pct') if c in out.columns}
    print(f"     [OK] {len(out)} 个交易日 {got}")
    return out

def fetch_vix(start, end):
    """^VIX 收盘价。"""
    print(">> [2/2] 下载 VIX ...")
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
    # 历史长度直接决定检验的可信度: 60 日视角下, 4 年只有约 13 个独立观测,
    # 任何结论都是噪声。拉到 15 年约 50 个, 仍不算多但已能做判断。
    # 注意 UNIVERSE 是"今天的"标普100成分, 越往前回溯幸存者偏差越强 ——
    # 早年不存在的票(META 2012、TSLA 2010、ABBV 2013、PYPL 2015...)那几天没有数据,
    # 由 MIN_UNIVERSE 挡掉; 但"当年有、今天没了"的票永远不会出现在样本里。
    # 用它做长期回测要知道这一点, validate_signal.py 的输出会提示。
    default_start = (pd.Timestamp.today() - pd.Timedelta(days=365 * 15)).strftime('%Y-%m-%d')
    start = sys.argv[1] if len(sys.argv) > 1 else default_start
    end = (pd.Timestamp.today() + pd.Timedelta(days=1)).strftime('%Y-%m-%d')

    print("=" * 56)
    print(">> 美股情绪指标数据抓取")
    print("=" * 56)

    equity, vix = fetch_equity_metrics(start, end), fetch_vix(start, end)
    if equity.empty and vix.empty:
        print("\n[错误] 三项数据全部抓取失败, 不写入任何内容")
        sys.exit(1)

    out = None
    for part in (equity, vix):
        if part.empty:
            continue
        out = part if out is None else out.merge(part, on='date', how='outer')
    out = out.sort_values('date').reset_index(drop=True)

    # 丢掉"今天"这一行: 美东当日收盘前 yfinance 会返回一根**未完成**的盘中 K 线,
    # 价格是此刻的价格而不是收盘价。2026-09-17 就这样进过两次结果文件(宽度 60.4%),
    # 而那一场当时还没开盘结束。MIN_SUBS 只在个股整体缺数据时才挡得住, 挡不住
    # "有数但是盘中数"。少一天没关系 —— 下一轮 CI 自然会把它补上。
    # 按美东收盘时间判断, 而不是简单地丢掉"今天" —— 收盘后那一行是完整的,
    # 无条件丢弃会白白少一个交易日。16:00 收盘, 留 15 分钟给数据落库。
    now_et = pd.Timestamp.now(tz='America/New_York')
    if now_et.hour * 60 + now_et.minute < 16 * 60 + 15:
        cutoff = now_et.strftime('%Y-%m-%d')
        n_before = len(out)
        out = out[out['date'] < cutoff].reset_index(drop=True)
        if len(out) < n_before:
            print(f"     [跳过] 美东 {now_et:%H:%M} 尚未收盘, 丢弃 {n_before - len(out)} 行未完成K线")

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
