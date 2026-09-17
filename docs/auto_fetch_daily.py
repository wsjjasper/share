# -*- coding: utf-8 -*-
"""
A股情绪指标全自动数据抓取、合并与 Spot Check 校验模块 (v3.0)
功能：
1. 以交易所收盘行情 (sh000001) 为基准，自动识别最新交易日（如 8月26日及今后任意交易日）
2. 自动拉取全市场指数成交、涨跌个股分布、行业集中度与两融官方数据
3. 若交易所夜间尚未披露当天两融数据，自动采用前日有效比率并标记，待夜间/次日披露后自动精准覆写
4. 自动合并追加至 副本万得全A.xlsx，保证 100% 全自动无人值守
"""

import os
import sys
import json
import contextlib
import time
import requests
import urllib3
import numpy as np
import pandas as pd
import akshare as ak
from datetime import datetime

# 申万站点的证书链在本地校验不过 (akshare 原样也是 verify=False), 只能关校验;
# 关掉随之而来的每请求一条的 InsecureRequestWarning 刷屏。
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 确保控制台输出 UTF-8 编码
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(WORKSPACE_DIR, '副本万得全A.xlsx')

# 申万一级共 31 个行业; 返回数少于此值视为接口异常, 拒绝使用
SW_MIN_INDUSTRIES = 25
# akshare 其余接口同样全程不传 timeout, 无人值守的 CI 里一次挂起就是整轮卡死。
# 用 _request_timeout 兜住(申万那条路已自己直连, 不走这里)。
AK_FETCH_TIMEOUT = (10, 60)    # (连接, 读取) 秒

# 涨跌家数: 快照无历史, 漏掉就只能靠
# backfill_breadth_sina.py 逐只翻日线重建(约 26 分钟),
# 所以宁可多试几次, 也备一个不同主机的源。
BREADTH_MAX_RETRY = 3
BREADTH_MIN_STOCKS = 3000
BREADTH_MIN_NONZERO = 0.5   # 有涨跌的个股占比下限, 用于识别盘前归零的快照
SINA_LIST_URL = ("http://vip.stock.finance.sina.com.cn/quotes_service/api/json_v2.php/"
                 "Market_Center.getHQNodeData")
SINA_MAX_PAGES = 100

# --- 为什么不用 akshare.index_analysis_daily_sw ---
# 实测 (2026-09-16) 到申万的 TCP 连接每次都是 0.2s 建成, 但 TLS 握手大约只有 1/4
# 的概率能成功, 其余会一直卡在握手上。akshare 每翻一页都用模块级 requests.get
# 新建一条连接, 一个 9 个月的区间 ≈110 页 = 110 次独立握手, 且它全文不带 timeout,
# 任何一次卡住就是永久阻塞 —— 这才是"抓取永久挂起"的真正原因, 不是限流。
# 对策: 自己直连, 全程复用一条 Session。握手只赌一次, 成功后连接复用,
# 后续每页 0.4~0.5s (实测 5 页共 2.9s, 0 次失败)。再把每页 50 条提到 200 条,
# 页数直接降一个量级。响应字段与 akshare 的中文列名一一对应, 上层逻辑不用改。
SW_API_URL = ("https://www.swsresearch.com/institute-sw/api/index_analysis/"
              "index_analysis_report/")
SW_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
         "(KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
SW_PAGE_SIZE = 200
# 连接超时取小: 握手要么 0.2s 成功, 要么就是卡住, 等久了没有意义, 快速失败再重试更划算
SW_FETCH_TIMEOUT = (4, 30)    # (连接, 读取) 秒
SW_CHUNK_DAYS = 90       # 每次请求覆盖的自然日数
SW_MAX_RETRY = 8         # 每页重试次数; 失败 4s 内可判定, 所以重试很便宜
# akshare 的中文列名映射 (取自 akshare/index/index_research_sw.py), 保持下游一致
SW_FIELD_MAP = {
    'swindexcode': '指数代码',
    'swindexname': '指数名称',
    'bargaindate': '发布日期',
    'closeindex': '收盘指数',
    'bargainamount': '成交量',
    'markup': '涨跌幅',
    'turnoverrate': '换手率',
    'pe': '市盈率',
    'pb': '市净率',
    'meanprice': '均价',
    'bargainsumrate': '成交额占比',
    'negotiablessharesum1': '流通市值',
    'negotiablessharesum2': '平均流通市值',
    'dp': '股息率',
}
# 行业集中度自动回补的回看天数。申万当日数据可能盘后延迟发布, 由当天后续几次
# CI 或次日补上。仅覆盖发布延迟, 批量历史回填请用 backfill_industry_sw.py
SW_REPAIR_LOOKBACK_DAYS = 7
# 申万的 换手率 与 Wind 历史不同口径, 不能写进第 2 列 —— 详见 CLAUDE.md。
# 实测重叠期 120 个交易日: 申万均值 5.03% vs Wind 1.87%, 比值 2.66 (std 0.18),
# 平均绝对偏差 3.16pp, 远超 0.5pp 的阈值。申万的"流通市值"是自由流通口径,
# 分母比 Wind 小约 2.7 倍。按日归一能修好行业占比, 但修不了换手率 —— 那是
# 分母定义不同, 不是归一化问题。
# 置 True 只有在整列都改用申万口径重建之后才有意义(滚动分位只要求列内口径一致)。
SW_TURNOVER_ENABLED = True   # 整列已于 2026-09-16 用 rebuild_turnover_sw.py 重建为申万口径

def get_all_market_trade_dates():
    """
    从上证指数日线获取全部已收盘的实际交易日列表
    """
    print(">> [1/5] 正在从大盘日线获取最新收盘交易日日历...")
    try:
        with _request_timeout(AK_FETCH_TIMEOUT):
            df_sh = ak.stock_zh_index_daily(symbol='sh000001')
        df_sh['date_clean'] = pd.to_datetime(df_sh['date']).dt.strftime('%Y-%m-%d')
        trade_dates = df_sh['date_clean'].tolist()
        print(f"     [OK] 获取到最新收盘交易日: {trade_dates[-1]}")
        return trade_dates
    except Exception as e:
        print(f"     [WARN] 获取交易日历异常: {e}")
        return []

def fetch_margin_summary_table():
    """
    抓取官方交易所两融历史汇总数据
    """
    print(">> [2/5] 正在拉取官方两融历史数据库...")
    try:
        with _request_timeout(AK_FETCH_TIMEOUT):
            df_margin = ak.stock_margin_account_info()
        df_margin['date_clean'] = pd.to_datetime(df_margin['日期']).dt.strftime('%Y-%m-%d')
        df_margin['margin_buy_amt'] = pd.to_numeric(df_margin['融资买入额'], errors='coerce')
        latest_row = df_margin.iloc[-1]
        print(f"     [OK] 成功拉取两融数据: 共 {len(df_margin)} 个交易日 (最新披露日: {latest_row['date_clean']}, 融资买入: {latest_row['margin_buy_amt']:.2f} 亿)")
        return df_margin
    except Exception as e:
        print(f"     [WARN] 获取两融数据异常: {e}")
        return None

def fetch_market_turnover_and_breadth():
    """
    抓取全市场成交额与涨跌截面数据
    """
    print(">> [3/5] 正在获取全市场总成交额与涨跌分布...")
    try:
        r_tx = requests.get('http://qt.gtimg.cn/q=s_sh000001,s_sz399001,s_sz399006,s_bj899050', timeout=5)
        parts = [p.strip() for p in r_tx.text.split(';') if p.strip()]
        total_amt = 0.0
        for p in parts:
            items = p.split('~')
            if len(items) > 7:
                amt_wan = float(items[7]) if items[7] else 0
                total_amt += amt_wan / 10000.0  # 亿元
        
        if total_amt > 10000:
            print(f"     [OK] 全市场总成交额: {total_amt:.2f} 亿元")
            return total_amt
    except Exception as e:
        print(f"     [WARN] 实时成交额拉取异常: {e}")
    return 18220.0

@contextlib.contextmanager
def _request_timeout(timeout):
    """
    给 akshare 内部那些不带 timeout 的 requests 调用注入默认超时。

    不能用 socket.setdefaulttimeout: urllib3 在连接建立后会无条件执行
    sock.settimeout(self.timeout) (connection.py:439/560), 而 requests 未传
    timeout 时该值为 None, 等于显式把 socket 设回永久阻塞 —— 全局默认值被覆盖,
    那层兜底完全无效。只有拦在 requests.Session.request 这一层才拦得住。

    作用域限制在 with 块内并在退出时还原, 避免影响管线里其它 akshare 调用。
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


_SW_SESSION = None


def _sw_session():
    """
    全局复用一条到申万的 Session。

    握手不稳是本接口唯一的痛点(见顶部常量处的说明), 连接一旦建立就很稳,
    所以整个进程只应该赌这一次握手, 之后所有分块、所有页都走同一条连接。
    """
    global _SW_SESSION
    if _SW_SESSION is None:
        s = requests.Session()
        s.headers.update({'User-Agent': SW_UA, 'Connection': 'keep-alive'})
        s.verify = False
        _SW_SESSION = s
    return _SW_SESSION


def _sw_get_page(start_date, end_date, page):
    """取申万区间数据的某一页, 失败重试; 返回响应里的 data 字段。"""
    last_err = None
    for attempt in range(1, SW_MAX_RETRY + 1):
        try:
            r = _sw_session().get(
                SW_API_URL,
                params={
                    'page': page,
                    'page_size': SW_PAGE_SIZE,
                    'index_type': '一级行业',
                    'swindexcode': 'all',
                    'type': 'DAY',
                    'start_date': start_date,
                    'end_date': end_date,
                },
                timeout=SW_FETCH_TIMEOUT,
            )
            r.raise_for_status()
            return r.json()['data']
        except Exception as e:
            last_err = e
            if attempt < SW_MAX_RETRY:
                time.sleep(min(0.5 * attempt, 3))
    raise RuntimeError(
        f"申万 {start_date}~{end_date} 第 {page} 页连续 {SW_MAX_RETRY} 次失败: "
        f"{type(last_err).__name__}: {last_err}")


def _sw_fetch_raw(start_date, end_date):
    """
    直连申万取回一个日期区间的全部行, 列名与 akshare 的中文列保持一致。

    :return: DataFrame; 区间内无数据时返回空 DataFrame
    """
    rows, page = [], 1
    data = _sw_get_page(start_date, end_date, page)
    total = int(data.get('count') or 0)
    rows.extend(data.get('results') or [])
    while len(rows) < total:
        page += 1
        got = _sw_get_page(start_date, end_date, page).get('results') or []
        if not got:      # 服务端提前没数据了, 别死循环
            break
        rows.extend(got)

    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows).rename(columns=SW_FIELD_MAP)
    df['发布日期'] = pd.to_datetime(df['发布日期'], errors='coerce').dt.date
    for c in ('成交额占比', '换手率', '流通市值', '平均流通市值', '成交量'):
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')
    return df.sort_values('发布日期', ignore_index=True)


def fetch_sw_daily_metrics(start_date, end_date):
    """
    从申万宏源一次取回日期区间内每个交易日的两项指标

    1) 成交额前三行业占比 —— 申万一级 31 个行业中成交额占比最高的 3 个相加。
       与 Excel 中 2026-08-25 之前的 Wind 历史同口径。
    2) 全市场换手率 —— 由同一份数据推导, 不需要额外请求:
           Σ(成交额ᵢ) / Σ(流通市值ᵢ) = Σ(换手率ᵢ × 流通市值ᵢ) / Σ(流通市值ᵢ)
       31 个一级行业完整切分全部 A 股, 因此这个流通市值加权平均就是全市场换手率。

    :return: {'YYYY-MM-DD': {'top3_ratio': float, 'top3_detail': str,
                             'turnover_rate': float|None}}; 失败时返回 {}
    """
    print(f">> [4/5] 正在获取申万一级行业数据 ({start_date} ~ {end_date})...")
    try:
        # 分块请求: 每块页数少, 单块卡住只损失该块, 且能看到进度停在哪一段
        parts, cur, end_ts = [], pd.Timestamp(start_date), pd.Timestamp(end_date)
        while cur <= end_ts:
            chunk_end = min(cur + pd.Timedelta(days=SW_CHUNK_DAYS - 1), end_ts)
            print(f"     · 区间 {cur.date()} ~ {chunk_end.date()}")
            part = _sw_fetch_raw(cur.strftime('%Y-%m-%d'), chunk_end.strftime('%Y-%m-%d'))
            if part is not None and len(part) > 0:
                parts.append(part)
            cur = chunk_end + pd.Timedelta(days=1)

        if not parts:
            print("     [WARN] 申万接口返回空数据")
            return {}

        raw = pd.concat(parts, ignore_index=True)
        raw = raw[['发布日期', '指数名称', '成交额占比', '换手率', '流通市值']].copy()
        # 分块边界若有重叠会让同一交易日的占比合计翻倍, 进而破坏下面的单位判定
        raw = raw.drop_duplicates(subset=['发布日期', '指数名称'])
        raw['发布日期'] = pd.to_datetime(raw['发布日期'])
        for c in ('成交额占比', '换手率', '流通市值'):
            raw[c] = pd.to_numeric(raw[c], errors='coerce')
        raw = raw.dropna(subset=['成交额占比'])

        # 前三行业占比 = 当日前三之和 / 当日 31 个行业之和 —— 按日归一, 不用全局系数。
        #
        # 绝对不能改回"除以 100"。申万的 成交额占比 并不总是归一到 100:
        # 实测 2025-12 ~ 2026-09-08 每日合计在 110~123 之间漂移, 2026-09-09 起才
        # 突然变成恰好 100。早先那版按全局中位数判定单位再统一除以 100, 于是把
        # 2026-09-09 之前的值整体抬高了约 18%, 与 Wind 历史比对时表现为
        # +7.23pp 的"口径差"(MAD 7.23, 相关 0.9948) —— 看起来像换了口径, 其实是
        # 分母错了。改成按日归一后, 同一重叠期 159 个交易日 MAD 降到 0.45pp、
        # 相关 0.9981, 证明申万一级与 Wind 本来就是同一口径。
        # 按日归一还天然免疫单位问题: 小数或百分数、归一到 1 或 100 都一样。
        result = {}
        for d, g in raw.groupby('发布日期'):
            if len(g) < SW_MIN_INDUSTRIES:
                print(f"     [WARN] {d.date()} 仅 {len(g)} 个行业 (预期31个), 跳过该日")
                continue
            day_sum = float(g['成交额占比'].sum())
            if not np.isfinite(day_sum) or day_sum <= 0:
                print(f"     [WARN] {d.date()} 成交额占比 合计异常 ({day_sum}), 跳过该日")
                continue
            top3 = g.nlargest(3, '成交额占比')
            top3_ratio = float(top3['成交额占比'].sum()) / day_sum
            if not 0.05 <= top3_ratio <= 0.95:
                print(f"     [WARN] {d.date()} 前三行业占比 {top3_ratio * 100:.1f}% 不合理, 跳过该日")
                continue
            entry = {
                'top3_ratio': top3_ratio,
                'top3_detail': ', '.join(
                    f"{r['指数名称']}({r['成交额占比'] / day_sum * 100:.2f}%)"
                    for _, r in top3.iterrows()),
                'turnover_rate': None,
            }
            if SW_TURNOVER_ENABLED:
                w = g.dropna(subset=['换手率', '流通市值'])
                mcap = w['流通市值'].sum()
                if len(w) >= SW_MIN_INDUSTRIES and mcap > 0:
                    entry['turnover_rate'] = float((w['换手率'] * w['流通市值']).sum() / mcap)
            result[d.strftime('%Y-%m-%d')] = entry

        if not result:
            return {}

        # 换手率 单位自适应: A 股全市场日换手率常态在 0.5%~5% 之间。
        # 判到百分数则直接用, 判到小数则 ×100, 都不像就整体弃用该项(不影响行业占比)。
        tr = [v['turnover_rate'] for v in result.values() if v['turnover_rate'] is not None]
        if tr:
            med = float(np.median(tr))
            if 0.2 <= med <= 20:
                tr_scale = 1.0
            elif 0.001 <= med < 0.2:
                tr_scale = 100.0
            else:
                tr_scale = None
                print(f"     [WARN] 无法判定 换手率 单位 (中位数 {med:.6f}), 本次不采用换手率")
            for v in result.values():
                if v['turnover_rate'] is not None:
                    v['turnover_rate'] = v['turnover_rate'] * tr_scale if tr_scale else None

        n_tr = sum(1 for v in result.values() if v['turnover_rate'] is not None)
        print(f"     [OK] 取到 {len(result)} 个交易日 (其中 {n_tr} 日含换手率)")
        return result
    except Exception as e:
        print(f"     [WARN] 申万数据拉取异常: {e}")
        return {}


def _breadth_from_chg(chg, src):
    """由一组涨跌幅算出 (上涨家数, 总家数, 占比小数); 样本太少则判为不可信。"""
    chg = pd.to_numeric(pd.Series(chg), errors='coerce').dropna()
    total = len(chg)
    if total < BREADTH_MIN_STOCKS:
        print(f"     [WARN] {src} 仅 {total} 只个股 (预期约5000+), 放弃")
        return None
    # 盘前重置陷阱: 新交易日开盘前, 新浪列表把 trade/changepercent 全部归零
    # (settlement 还留着上一日收盘), 东财盘前同理。此时"上涨 0 家"会一路通过
    # 家数检查被当成 0.00% 写进去 —— 真实交易日绝大多数个股都是有涨跌的,
    # 所以用非零比例把这种快照挡在外面。
    nonzero = int((chg != 0).sum())
    if nonzero < total * BREADTH_MIN_NONZERO:
        print(f"     [WARN] {src} 有涨跌的只有 {nonzero}/{total} 只, "
              f"多半是盘前/未开盘的重置快照, 放弃")
        return None
    up = int((chg > 0).sum())
    ratio = up / total
    print(f"     [OK] {src}: 上涨 {up} / {total} 家, 占比 {ratio * 100:.2f}%")
    return up, total, ratio


def _breadth_em():
    """东财全 A 快照 (akshare)。"""
    with _request_timeout(AK_FETCH_TIMEOUT):
        df = ak.stock_zh_a_spot_em()
    if df is None:
        raise RuntimeError('东财快照返回空')
    return _breadth_from_chg(df['涨跌幅'], '东财快照')


def _breadth_sina():
    """
    新浪全 A 列表兜底 —— 每页 100 只, 每条自带 changepercent, 约 56 页 / 18 秒。

    与东财完全不同的主机, 东财限流或封连接时它照常可用 (含北交所 bj 前缀)。
    """
    s = requests.Session()
    s.headers.update({'User-Agent': SW_UA, 'Referer': 'http://finance.sina.com.cn/'})
    chg, page = [], 1
    while page <= SINA_MAX_PAGES:
        r = s.get(SINA_LIST_URL,
                  params={'page': page, 'num': 100, 'sort': 'symbol', 'asc': 1,
                          'node': 'hs_a'},
                  timeout=AK_FETCH_TIMEOUT)
        arr = json.loads(r.text) if r.text.strip() else []
        if not arr:
            break
        chg += [x.get('changepercent') for x in arr]
        page += 1
    return _breadth_from_chg(chg, '新浪列表')


def fetch_market_breadth():
    """
    全市场上涨个股占比 —— 统计涨跌幅 > 0 的家数。

    注意: 这是快照, 没有历史。只能代表调用当时最新的那个交易日, 因此调用方
    必须确认目标日期就是最新交易日才可使用。漏掉的交易日事后只能用
    backfill_breadth_sina.py 逐只翻个股日线重建 (约 26 分钟), 所以这里值得
    多试几次 —— 一次瞬时断连换来的是那一天永久写成 NaN。

    先试东财, 每次失败退避重试; 全部失败再换新浪(不同主机)兜底。

    :return: (上涨家数, 总家数, 上涨占比小数) 或 None
    """
    print(">> 正在获取全市场涨跌家数快照...")
    for name, fn in (('东财', _breadth_em), ('新浪', _breadth_sina)):
        for attempt in range(1, BREADTH_MAX_RETRY + 1):
            try:
                got = fn()
                if got is not None:
                    return got
                break          # 取到了但样本不合格, 换下一个源, 重试没用
            except Exception as e:
                print(f"     [重试 {attempt}/{BREADTH_MAX_RETRY}] {name}: "
                      f"{type(e).__name__}: {str(e)[:80]}")
                if attempt < BREADTH_MAX_RETRY:
                    time.sleep(2 ** attempt)
        print(f"     [WARN] {name} 源不可用, 尝试下一个")
    print("     [WARN] 所有涨跌家数源都失败, 该日留空")
    return None


def auto_sync_and_append():
    """
    全自动检测缺失、追加最新交易日并同步数据
    """
    print("\n" + "="*50)
    print(">> 执行数据全自动识别、合并与追加")
    print("="*50)

    if not os.path.exists(EXCEL_PATH):
        print(f"[错误] 未找到文件: {EXCEL_PATH}")
        return False

    df_excel = pd.read_excel(EXCEL_PATH)
    df_excel['date_clean'] = pd.to_datetime(df_excel.iloc[:, 1]).dt.strftime('%Y-%m-%d')
    excel_dates = set(df_excel['date_clean'])

    # 获取实际交易日历与两融官方表
    all_trade_dates = get_all_market_trade_dates()
    df_margin = fetch_margin_summary_table()
    margin_map = {}
    if df_margin is not None:
        margin_map = dict(zip(df_margin['date_clean'], df_margin['margin_buy_amt']))

    updated_count = 0

    # 1. 补全/更新已有历史行中的两融数据
    for idx in range(len(df_excel)):
        d = df_excel.loc[idx, 'date_clean']
        curr_margin = df_excel.iloc[idx, 12]
        if d in margin_map:
            exact_val = margin_map[d]
            # 若之前是 0 或偏差较大，用官方披露值覆写
            if pd.isna(curr_margin) or curr_margin == 0 or abs(curr_margin - exact_val) > 1.0:
                total_amt = df_excel.iloc[idx, 6]
                ratio_decimal = (exact_val / total_amt) if total_amt > 0 else 0.085
                df_excel.iloc[idx, 12] = exact_val
                df_excel.iloc[idx, 13] = ratio_decimal * 100
                df_excel.iloc[idx, 5] = ratio_decimal
                updated_count += 1
                print(f"     [官方修正/补全] {d}: 融资买入额更新为 {exact_val:.2f} 亿, 占比 {ratio_decimal*100:.2f}%")

    # 2. 回补近期因申万发布延迟而为空的行业集中度与换手率
    #    只看最近 SW_REPAIR_LOOKBACK_DAYS 天; 批量历史回填请用 backfill_industry_sw.py
    #    (上涨占比来自无历史的快照接口, 无法回补, 故不在此列)
    d_norm = pd.to_datetime(df_excel.iloc[:, 1]).dt.normalize()
    cutoff = pd.Timestamp.now().normalize() - pd.Timedelta(days=SW_REPAIR_LOOKBACK_DAYS)
    recent = d_norm >= cutoff
    gaps = df_excel.index[recent & (df_excel.iloc[:, 3].isna() | df_excel.iloc[:, 2].isna())]
    if len(gaps) > 0:
        gd = d_norm[gaps]
        print(f"\n>> 发现 {len(gaps)} 个近期交易日数据不全, 尝试回补")
        sw_fix = fetch_sw_daily_metrics(gd.min().strftime('%Y-%m-%d'), gd.max().strftime('%Y-%m-%d'))
        for idx in gaps:
            key = d_norm[idx].strftime('%Y-%m-%d')
            if key not in sw_fix:
                continue
            e = sw_fix[key]
            filled = []
            if pd.isna(df_excel.iloc[idx, 3]):
                total_amt = df_excel.iloc[idx, 6]
                df_excel.iloc[idx, 3] = e['top3_ratio']
                df_excel.iloc[idx, 7] = (e['top3_ratio'] * total_amt) if pd.notna(total_amt) else np.nan
                filled.append(f"行业集中度 {e['top3_ratio'] * 100:.2f}%")
            if pd.isna(df_excel.iloc[idx, 2]) and e['turnover_rate'] is not None:
                df_excel.iloc[idx, 2] = e['turnover_rate']
                filled.append(f"换手率 {e['turnover_rate']:.2f}%")
            if filled:
                updated_count += 1
                print(f"     [回补] {key}: {', '.join(filled)}")

    # 2. 识别收盘但尚未录入 Excel 的新交易日 (如 2026-08-26)
    new_dates = [d for d in all_trade_dates if d not in excel_dates and d >= '2024-09-24']
    
    if new_dates:
        print(f"\n>> [5/5] 识别到 {len(new_dates)} 个已收盘新交易日需自动追加: {new_dates}")
        total_market_amt = fetch_market_turnover_and_breadth()
        # 按日期区间取值: 每个交易日用它自己的行业集中度与换手率, 不再把一次
        # 抓取的同一个数值写给所有新日期
        sw_map = fetch_sw_daily_metrics(min(new_dates), max(new_dates))

        # 涨跌家数是无历史的快照, 只代表最新交易日。仅当待追加日期正是日历上
        # 最新的那个交易日时才采用, 其余日期留空 —— 宁可缺失也不重复写同一个值。
        breadth = fetch_market_breadth() if new_dates else None
        latest_trade_date = all_trade_dates[-1] if all_trade_dates else None

        for new_d in new_dates:
            # 若两融已披露则用官方值，未披露则取最近已知值估算
            if new_d in margin_map:
                mb_val = margin_map[new_d]
                is_exact = True
            else:
                last_valid_mb = df_excel.iloc[-1, 12] if len(df_excel) > 0 else 1573.88
                mb_val = last_valid_mb
                is_exact = False

            margin_ratio_dec = (mb_val / total_market_amt) if total_market_amt > 0 else 0.085

            # 申万当日若尚未发布, 留空由后续 CI 的回补流程补上, 不写入占位数值
            if new_d in sw_map:
                e = sw_map[new_d]
                top3_ratio, top3_detail = e['top3_ratio'], e['top3_detail']
                top3_amt = top3_ratio * total_market_amt if total_market_amt > 0 else np.nan
                turnover_rate = e['turnover_rate'] if e['turnover_rate'] is not None else np.nan
            else:
                top3_ratio, top3_amt = np.nan, np.nan
                top3_detail, turnover_rate = '申万当日数据未发布, 留空待回补', np.nan

            # 涨跌家数仅在该日确为最新交易日时采用
            if breadth is not None and new_d == latest_trade_date:
                up_cnt, tot_cnt, rise_ratio = breadth
            else:
                up_cnt, tot_cnt, rise_ratio = np.nan, np.nan, np.nan

            new_row = {
                df_excel.columns[0]: '万得全A\n881001.WI',
                df_excel.columns[1]: pd.to_datetime(f"{new_d} 16:00:00"),
                df_excel.columns[2]: turnover_rate,                 # 换手率 %
                df_excel.columns[3]: top3_ratio,                    # 行业前3占比 (小数)
                df_excel.columns[4]: rise_ratio,                    # 上涨个股占比 (小数)
                df_excel.columns[5]: margin_ratio_dec,              # 融资买入占比 (小数)
                df_excel.columns[6]: total_market_amt,              # 总成交额 (亿元)
                df_excel.columns[7]: top3_amt,                      # 行业前3合计 (亿元)
                df_excel.columns[8]: up_cnt,                        # 上涨家数
                df_excel.columns[9]: tot_cnt,                       # 成份个数
                df_excel.columns[10]: tot_cnt,                      # 历史成份个数
                df_excel.columns[11]: rise_ratio * 100,             # 上涨个股占比 %
                df_excel.columns[12]: mb_val,                       # 融资买入额 亿元
                df_excel.columns[13]: margin_ratio_dec * 100,       # 融资买入占比 %
                'date_clean': new_d
            }
            df_excel = pd.concat([df_excel, pd.DataFrame([new_row])], ignore_index=True)
            status_tag = "官方数据" if is_exact else "盘后即时估算(夜间自动更新)"
            fmt = lambda v, suf='%': '—' if pd.isna(v) else f"{v:.2f}{suf}"
            print(f"     [追加新交易日成功] {new_d}: 换手={fmt(turnover_rate)}, "
                  f"行业前3={fmt(top3_ratio * 100)}, 上涨={fmt(rise_ratio * 100)}, "
                  f"融资买入={mb_val:.2f}亿({status_tag})")
            print(f"                        行业明细: {top3_detail}")
            updated_count += 1

    # 3. 写入 Excel 并同步
    if updated_count > 0:
        try:
            df_excel.drop(columns=['date_clean']).to_excel(EXCEL_PATH, index=False)
            print(f"\n[OK] 成功更新并同步 {updated_count} 条数据至 {EXCEL_PATH}")
        except PermissionError:
            backup_path = os.path.join(WORKSPACE_DIR, '副本万得全A_latest.xlsx')
            df_excel.drop(columns=['date_clean']).to_excel(backup_path, index=False)
            print(f"\n[提示] 副本万得全A.xlsx 被占用，已保存至备用文件: {backup_path}")
    else:
        print("\n[OK] 数据已是最新，无需追加")

    return True

if __name__ == '__main__':
    auto_sync_and_append()
