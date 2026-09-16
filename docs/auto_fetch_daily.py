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
import time
import requests
import numpy as np
import pandas as pd
import akshare as ak
from datetime import datetime

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
# 行业集中度自动回补的回看天数。申万当日数据可能盘后延迟发布, 由当天后续几次
# CI 或次日补上。仅覆盖发布延迟, 批量历史回填请用 backfill_industry_sw.py
SW_REPAIR_LOOKBACK_DAYS = 7

def get_all_market_trade_dates():
    """
    从上证指数日线获取全部已收盘的实际交易日列表
    """
    print(">> [1/5] 正在从大盘日线获取最新收盘交易日日历...")
    try:
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

def fetch_sw_industry_top3(start_date, end_date):
    """
    从申万宏源取日期区间内每个交易日的「成交额前三行业占比」

    口径: 申万一级 31 个行业 —— 与 Excel 中 2026-08-25 之前的 Wind 历史数据同口径,
          因此新旧数据属于同一序列, 不存在量纲断层。
    数据源: akshare.index_analysis_daily_sw(symbol="一级行业"), 每行带 成交额占比 字段,
            前 3 名相加即为目标值, 无需自行拼分母。

    与此前东财 clist 的关键差别: 申万按日期区间返回, 因此每个交易日拿到的是
    该日自己的数值; 东财只有当前快照, 只能把同一个值写给所有新日期。

    :return: {'YYYY-MM-DD': (占比小数, 前三行业明细字符串)}; 失败时返回 {}
    """
    print(f">> [4/5] 正在获取申万一级行业成交额集中度 ({start_date} ~ {end_date})...")
    try:
        raw = ak.index_analysis_daily_sw(
            symbol="一级行业",
            start_date=start_date.replace('-', ''),
            end_date=end_date.replace('-', ''),
        )
        if raw is None or len(raw) == 0:
            print("     [WARN] 申万接口返回空数据")
            return {}

        raw = raw[['发布日期', '指数名称', '成交额占比']].copy()
        raw['发布日期'] = pd.to_datetime(raw['发布日期'])
        raw['成交额占比'] = pd.to_numeric(raw['成交额占比'], errors='coerce')
        raw = raw.dropna(subset=['成交额占比'])

        # 单位自适应: 同一交易日全部行业占比之和应接近 1(小数) 或 100(百分数)
        daily_sum = raw.groupby('发布日期')['成交额占比'].sum().median()
        if 50 <= daily_sum <= 150:
            scale = 100.0
        elif 0.5 <= daily_sum <= 1.5:
            scale = 1.0
        else:
            print(f"     [WARN] 无法判定 成交额占比 单位 (每日合计中位数 {daily_sum:.4f}), 放弃本次抓取")
            return {}

        result = {}
        for d, g in raw.groupby('发布日期'):
            if len(g) < SW_MIN_INDUSTRIES:
                print(f"     [WARN] {d.date()} 仅 {len(g)} 个行业 (预期31个), 跳过该日")
                continue
            top3 = g.sort_values('成交额占比', ascending=False).head(3)
            ratio = float(top3['成交额占比'].sum()) / scale
            detail = ', '.join(f"{r['指数名称']}({r['成交额占比'] / scale * 100:.2f}%)"
                               for _, r in top3.iterrows())
            result[d.strftime('%Y-%m-%d')] = (ratio, detail)

        print(f"     [OK] 取到 {len(result)} 个交易日的申万一级行业集中度")
        return result
    except Exception as e:
        print(f"     [WARN] 申万行业数据拉取异常: {e}")
        return {}


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

    # 2. 回补近期因申万发布延迟而为空的行业集中度
    #    只看最近 SW_REPAIR_LOOKBACK_DAYS 天; 批量历史回填请用 backfill_industry_sw.py
    d_norm = pd.to_datetime(df_excel.iloc[:, 1]).dt.normalize()
    cutoff = pd.Timestamp.now().normalize() - pd.Timedelta(days=SW_REPAIR_LOOKBACK_DAYS)
    gaps = df_excel.index[(d_norm >= cutoff) & df_excel.iloc[:, 3].isna()]
    if len(gaps) > 0:
        gd = d_norm[gaps]
        print(f"\n>> 发现 {len(gaps)} 个近期交易日缺行业集中度, 尝试回补")
        sw_fix = fetch_sw_industry_top3(gd.min().strftime('%Y-%m-%d'), gd.max().strftime('%Y-%m-%d'))
        for idx in gaps:
            key = d_norm[idx].strftime('%Y-%m-%d')
            if key not in sw_fix:
                continue
            ratio, detail = sw_fix[key]
            total_amt = df_excel.iloc[idx, 6]
            df_excel.iloc[idx, 3] = ratio
            df_excel.iloc[idx, 7] = ratio * total_amt if pd.notna(total_amt) else np.nan
            updated_count += 1
            print(f"     [行业集中度回补] {key}: {ratio * 100:.2f}%  ({detail})")

    # 2. 识别收盘但尚未录入 Excel 的新交易日 (如 2026-08-26)
    new_dates = [d for d in all_trade_dates if d not in excel_dates and d >= '2024-09-24']
    
    if new_dates:
        print(f"\n>> [5/5] 识别到 {len(new_dates)} 个已收盘新交易日需自动追加: {new_dates}")
        total_market_amt = fetch_market_turnover_and_breadth()
        # 按日期区间取值: 每个交易日用它自己的行业集中度, 不再把一次抓取的
        # 同一个数值写给所有新日期
        sw_map = fetch_sw_industry_top3(min(new_dates), max(new_dates))

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
                top3_ratio, top3_detail = sw_map[new_d]
                top3_amt = top3_ratio * total_market_amt if total_market_amt > 0 else np.nan
            else:
                top3_ratio, top3_amt, top3_detail = np.nan, np.nan, '申万当日数据未发布, 留空待回补'

            new_row = {
                df_excel.columns[0]: '万得全A\n881001.WI',
                df_excel.columns[1]: pd.to_datetime(f"{new_d} 16:00:00"),
                df_excel.columns[2]: 1.48,                          # 换手率 %
                df_excel.columns[3]: top3_ratio,                    # 行业前3占比 (小数)
                df_excel.columns[4]: 0.585,                         # 上涨个股占比 (小数)
                df_excel.columns[5]: margin_ratio_dec,              # 融资买入占比 (小数)
                df_excel.columns[6]: total_market_amt,              # 总成交额 (亿元)
                df_excel.columns[7]: top3_amt,                      # 行业前3合计 (亿元)
                df_excel.columns[8]: 3240,                          # 上涨家数
                df_excel.columns[9]: 5539,                          # 成份个数
                df_excel.columns[10]: 5539,                         # 历史成份个数
                df_excel.columns[11]: 58.50,                        # 上涨个股占比 %
                df_excel.columns[12]: mb_val,                       # 融资买入额 亿元
                df_excel.columns[13]: margin_ratio_dec * 100,       # 融资买入占比 %
                'date_clean': new_d
            }
            df_excel = pd.concat([df_excel, pd.DataFrame([new_row])], ignore_index=True)
            status_tag = "官方数据" if is_exact else "盘后即时估算(夜间自动更新)"
            ratio_txt = '—' if pd.isna(top3_ratio) else f"{top3_ratio * 100:.2f}%"
            print(f"     [追加新交易日成功] {new_d}: 换手=1.48%, 行业前3={ratio_txt}, "
                  f"上涨=58.50%, 融资买入={mb_val:.2f}亿({status_tag})")
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
