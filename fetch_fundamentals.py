# -*- coding: utf-8 -*-
"""
基本面: 拉取同花顺主要指标报表并合并

    python fetch_fundamentals.py            # 按季度规则增量拉取 + 重新合并
    python fetch_fundamentals.py --force    # 忽略规则, 全部重新拉取

输入  fundamentals_codes.txt   股票代码列表, 每行一个 (# 开头为注释, 逗号/空格分隔也可)
缓存  fundamentals/raw/{code}.csv  每只股票最近一次成功下载的原始报表 (科目 × 报告期)
状态  fundamentals_state.json  每只股票的名称、已拉取最新报告期、拉取/检查时间
输出  基本面_合并.xlsx
        - 「最新一期」: 每只股票 1 行, 列为该股最新报告期的各项指标
        - 「全部报告期」: 每只股票 1 行, 列为 "报告期 指标" (行转列), 最新报告期在前

拉取规则 (财报按季度发布):
    应有报告期 = 严格早于今天的最近一个季末 (3-31 / 6-30 / 9-30 / 12-31)。
    已拉取的最新报告期 >= 应有报告期 → 不拉取。
    例: 已有 2026-09-30, 今天 2026-10-31 → 跳过; 到 2027-01-01 应有 2026-12-31 → 重新拉取。
    应有但还没披露时 (年报最晚 4 月底才出), 每个自然日最多重试一次, 避免 CI 一天三跑反复请求。

下载失败或返回空报表时保留旧缓存, 绝不用空数据覆盖已有结果。
"""

import argparse
import io
import json
import os
import sys
import time
from datetime import date, datetime

import pandas as pd
import requests

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CODES_FILE = os.path.join(BASE_DIR, 'fundamentals_codes.txt')
RAW_DIR = os.path.join(BASE_DIR, 'fundamentals', 'raw')
STATE_FILE = os.path.join(BASE_DIR, 'fundamentals_state.json')
OUTPUT_FILE = os.path.join(BASE_DIR, '基本面_合并.xlsx')

REPORT_URL = 'https://basic.10jqka.com.cn/api/stock/export.php?export=main&type=report&code={code}'
NAME_URL = 'https://qt.gtimg.cn/q={symbols}'
HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                         '(KHTML, like Gecko) Chrome/124.0 Safari/537.36',
           'Referer': 'https://basic.10jqka.com.cn/'}
TIMEOUT = (5, 30)
MAX_RETRY = 3
REQUEST_GAP = 0.5   # 两只股票之间的间隔(秒), 对同花顺客气一点

PERIOD_HEADER = '科目\\时间'


# ---------------------------------------------------------------- 报告期规则
def expected_period(today: date) -> str:
    """严格早于 today 的最近一个季末日期 (YYYY-MM-DD)。"""
    y = today.year
    for m, d in ((12, 31), (9, 30), (6, 30), (3, 31)):
        q = date(y, m, d)
        if q < today:
            return q.isoformat()
    return date(y - 1, 12, 31).isoformat()


def needs_fetch(info: dict, expected: str, today: date, force: bool) -> tuple[bool, str]:
    if force:
        return True, '强制拉取'
    latest = info.get('latest_period')
    if not latest or not os.path.exists(raw_path(info['code'])):
        return True, '尚无缓存'
    if latest >= expected:
        return False, f'已有 {latest}, 应有 {expected}'
    if info.get('last_checked') == today.isoformat():
        return False, f'应有 {expected} 但今日已检查过 (仍为 {latest})'
    return True, f'已有 {latest} < 应有 {expected}'


# ---------------------------------------------------------------- 输入输出
def read_codes() -> list[str]:
    if not os.path.exists(CODES_FILE):
        return []
    codes = []
    with open(CODES_FILE, encoding='utf-8') as f:
        for line in f:
            line = line.split('#', 1)[0]
            for tok in line.replace(',', ' ').replace('，', ' ').split():
                tok = tok.strip()
                if tok.isdigit() and len(tok) == 6 and tok not in codes:
                    codes.append(tok)
                elif tok:
                    print(f'[警告] 忽略无法识别的代码: {tok}')
    return codes


def raw_path(code: str) -> str:
    return os.path.join(RAW_DIR, f'{code}.csv')


def load_state() -> dict:
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_state(state: dict) -> None:
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2, sort_keys=True)


def load_raw(code: str) -> pd.DataFrame | None:
    p = raw_path(code)
    if not os.path.exists(p):
        return None
    return pd.read_csv(p, index_col=0, dtype=str, encoding='utf-8-sig')


# ---------------------------------------------------------------- 抓取
def fetch_report(code: str) -> pd.DataFrame:
    """下载并解析一只股票的报表, 返回 index=科目, columns=报告期(新→旧) 的 DataFrame。"""
    last_err = None
    for attempt in range(1, MAX_RETRY + 1):
        try:
            r = requests.get(REPORT_URL.format(code=code), headers=HEADERS, timeout=TIMEOUT)
            r.raise_for_status()
            if not r.content:
                # 同花顺对不存在的代码返回 200 + 空 body, 重试无意义
                raise ValueError('返回空报表 (代码不存在或无数据)')
            # 同花顺生成的 xls 复合文档目录有瑕疵, xlrd 默认会拒绝, 必须忽略
            df = pd.read_excel(io.BytesIO(r.content), header=None, dtype=str,
                               engine='xlrd',
                               engine_kwargs={'ignore_workbook_corruption': True})
            return parse_report(df)
        except ValueError as e:
            raise RuntimeError(str(e)) from e
        except Exception as e:  # 网络错误: 退避重试
            last_err = e
            if attempt < MAX_RETRY:
                time.sleep(2 ** attempt)
    raise RuntimeError(f'下载失败 ({MAX_RETRY} 次): {last_err}')


def parse_report(df: pd.DataFrame) -> pd.DataFrame:
    # 找到 "科目\时间" 所在的表头行, 不写死行号
    hdr = None
    for i in range(min(len(df), 10)):
        if str(df.iat[i, 0]).strip() == PERIOD_HEADER:
            hdr = i
            break
    if hdr is None:
        raise ValueError(f'报表格式不符: 未找到表头 "{PERIOD_HEADER}"')
    periods = [str(p).strip()[:10] for p in df.iloc[hdr, 1:]]
    body = df.iloc[hdr + 1:].copy()
    body = body[body.iloc[:, 0].notna()]
    out = pd.DataFrame(body.iloc[:, 1:].values,
                       index=[str(x).strip() for x in body.iloc[:, 0]],
                       columns=periods)
    out = out.loc[:, [c for c in out.columns if c and c != 'nan']]
    out.index.name = '科目'
    valid = [c for c in out.columns if len(c) == 10 and c[4] == '-' and c[7] == '-']
    if not valid or out.empty:
        raise ValueError('报表格式不符: 没有报告期列')
    out = out[sorted(valid, reverse=True)]
    return out


def market_prefix(code: str) -> str:
    if code.startswith(('6', '9')) and not code.startswith('92'):
        return 'sh'
    if code.startswith(('4', '8', '92')):
        return 'bj'
    return 'sz'


def fetch_names(codes: list[str]) -> dict:
    """腾讯行情取股票简称; 失败不影响主流程。"""
    names = {}
    if not codes:
        return names
    try:
        symbols = ','.join(market_prefix(c) + c for c in codes)
        r = requests.get(NAME_URL.format(symbols=symbols), headers=HEADERS, timeout=TIMEOUT)
        r.raise_for_status()
        for line in r.content.decode('gbk', errors='replace').split(';'):
            if '="' not in line:
                continue
            parts = line.split('="', 1)[1].split('~')
            if len(parts) > 2 and parts[2] in codes and parts[1]:
                names[parts[2]] = parts[1]
    except Exception as e:
        print(f'[提示] 股票简称获取失败, 沿用旧名称: {e}')
    return names


# ---------------------------------------------------------------- 合并
def to_number(v):
    """'160222200' → 数值; '100.20%' 与 '--' 之类保持原样。"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    if s in ('', 'nan', 'False', '--'):
        return None
    try:
        f = float(s.replace(',', ''))
        return int(f) if f.is_integer() and abs(f) < 1e15 else f
    except ValueError:
        return s


def build_merged(codes: list[str], state: dict):
    latest_rows, wide_rows = [], []
    metric_order = []
    all_periods = set()
    for code in codes:
        raw = load_raw(code)
        if raw is None or raw.empty:
            continue
        for m in raw.index:
            if m not in metric_order:
                metric_order.append(m)
        all_periods.update(raw.columns)
        info = state.get(code, {})
        base = {'代码': code, '名称': info.get('name', ''), '最新报告期': raw.columns[0]}

        latest = dict(base)
        for m in raw.index:
            latest[m] = to_number(raw.at[m, raw.columns[0]])
        latest_rows.append(latest)

        wide = dict(base)
        for p in raw.columns:
            for m in raw.index:
                wide[(p, m)] = to_number(raw.at[m, p])
        wide_rows.append(wide)

    if not latest_rows:
        return None, None

    head = ['代码', '名称', '最新报告期']
    df_latest = pd.DataFrame(latest_rows).reindex(columns=head + metric_order)

    periods = sorted(all_periods, reverse=True)
    wide_cols = [(p, m) for p in periods for m in metric_order]
    df_wide = pd.DataFrame(wide_rows).reindex(columns=head + wide_cols)
    df_wide.columns = head + [f'{p} {m}' for p, m in wide_cols]
    # 某一期所有股票都没有的指标列整列丢掉 (如个别指标仅在年报出现)
    keep = head + [c for c in df_wide.columns[len(head):] if df_wide[c].notna().any()]
    return df_latest, df_wide[keep]


def write_excel(df_latest: pd.DataFrame, df_wide: pd.DataFrame) -> bool:
    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as w:
            df_latest.to_excel(w, sheet_name='最新一期', index=False)
            df_wide.to_excel(w, sheet_name='全部报告期', index=False)
            for ws in w.book.worksheets:
                ws.freeze_panes = 'D2'
                ws.column_dimensions['A'].width = 9
                ws.column_dimensions['B'].width = 11
                ws.column_dimensions['C'].width = 12
        return True
    except PermissionError:
        # 不像其它脚本那样改写到 *_latest.xlsx —— 那会让页面下载到旧文件而无人察觉
        print(f'[错误] {OUTPUT_FILE} 被占用 (是否在 Excel 中打开?), 本次未写入合并结果')
        return False


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(description='拉取同花顺主要指标报表并合并')
    ap.add_argument('--force', action='store_true', help='忽略季度规则, 全部重新拉取')
    args = ap.parse_args()

    codes = read_codes()
    if not codes:
        print(f'[提示] {os.path.basename(CODES_FILE)} 中没有股票代码, 跳过基本面更新')
        return
    os.makedirs(RAW_DIR, exist_ok=True)

    today = date.today()
    expected = expected_period(today)
    state = load_state()
    print(f'股票 {len(codes)} 只 · 今天 {today} · 应有报告期 {expected}')

    names = fetch_names(codes)
    fetched = failed = 0
    for i, code in enumerate(codes):
        info = state.setdefault(code, {})
        info['code'] = code
        if names.get(code):
            info['name'] = names[code]
        do, why = needs_fetch(info, expected, today, args.force)
        label = f"{code} {info.get('name', '')}".strip()
        if not do:
            print(f'  跳过 {label}: {why}')
            continue
        if fetched or failed:
            time.sleep(REQUEST_GAP)
        info['last_checked'] = today.isoformat()
        try:
            df = fetch_report(code)
        except Exception as e:
            failed += 1
            print(f'  [失败] {label}: {e} (保留旧缓存)')
            continue
        df.to_csv(raw_path(code), encoding='utf-8-sig')
        prev = info.get('latest_period')
        info['latest_period'] = df.columns[0]
        info['fetched_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        fetched += 1
        note = '' if prev != df.columns[0] else ' (尚未披露新一期)'
        print(f'  拉取 {label}: {why} → 最新 {df.columns[0]}, {df.shape[1]} 期{note}')

    save_state(state)

    df_latest, df_wide = build_merged(codes, state)
    if df_latest is None:
        print('[提示] 没有任何可合并的报表')
        return
    if write_excel(df_latest, df_wide):
        print(f'合并完成: {os.path.basename(OUTPUT_FILE)} · {len(df_latest)} 只股票 · '
              f'全部报告期 {df_wide.shape[1] - 3} 列 · 本次拉取 {fetched}, 失败 {failed}')


if __name__ == '__main__':
    main()
