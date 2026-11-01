# -*- coding: utf-8 -*-
"""
基本面: 拉取同花顺主要指标报表, 填入用户模板 基本面_模板.xlsx, 输出 基本面_合并.xlsx

    python fetch_fundamentals.py            # 按季度规则增量拉取 + 重新生成合并文件
    python fetch_fundamentals.py --force    # 忽略规则, 全部重新拉取

模板 基本面_模板.xlsx (用户提供, 本脚本只读不写)
    - 每个 sheet 是一个公司分类 (光学 / LED / 有色 / 军工 …), A 列代码, B 列名称, 第 3 行起每行一只股票。
      股票列表就取自这里 —— 想增删股票, 改模板即可。
    - 第 1 行是指标名, 第 2 行是期间 (2026Q3 这样的季度, 或 2019 这样的年份)。
    - 需要填数的是 METRICS 里的 5 个季度数据块, 由第 1 行指标名 + 第 2 行季度标签识别,
      不写死列号。其余列 (公式、年度数据、备注) 一律不动。

生成规则
    - 只填**空**单元格; 模板里已有的数值保持原样 (会与同花顺比对并打印差异, 但不覆盖)。
    - 同花顺有、模板数据块里还没有的新季度 (如 2027Q1), 在该数据块末尾**插入新列**。插列时
      全表公式、条件格式范围、列宽都会跟着平移, 效果与在 Excel 里手工插列相同 ——
      引用 2026Q3 的公式插列后仍指向 2026Q3。
    - 第 1 行连续重复的指标名合并为一个单元格。
    - 模板里的线程批注 openpyxl 无法保存, 转为普通批注, 文字内容保留。
    - 百分比按模板口径存为小数 (39.46% → 0.3946), 每股收益存原值。

拉取规则 (财报按季度发布):
    应有报告期 = 严格早于今天的最近一个季末。已拉取的最新报告期 >= 应有报告期 → 不拉取。
    例: 已有 2026-09-30, 今天 2026-10-31 → 跳过; 到 2027-01-01 应有 2026-12-31 → 重新拉取。
    应有但还没披露时, 每个自然日最多重试一次, 避免 CI 一天三跑反复请求。

缓存  fundamentals/raw/{code}.csv  每只股票最近一次成功下载的原始报表 (科目 × 报告期)
状态  fundamentals_state.json      已拉取最新报告期、拉取/检查时间
下载失败或返回空报表时保留旧缓存, 绝不用空数据覆盖已有结果。
"""

import argparse
import html
import io
import json
import os
import re
import sys
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import copy
from datetime import date, datetime

import openpyxl
import pandas as pd
import requests
from openpyxl.comments import Comment
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.styles import Alignment
from openpyxl.utils import column_index_from_string, get_column_letter

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, '基本面_模板.xlsx')
RAW_DIR = os.path.join(BASE_DIR, 'fundamentals', 'raw')
STATE_FILE = os.path.join(BASE_DIR, 'fundamentals_state.json')
OUTPUT_FILE = os.path.join(BASE_DIR, '基本面_合并.xlsx')

REPORT_URL = 'https://basic.10jqka.com.cn/api/stock/export.php?export=main&type=report&code={code}'
HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                         '(KHTML, like Gecko) Chrome/124.0 Safari/537.36',
           'Referer': 'https://basic.10jqka.com.cn/'}
TIMEOUT = (5, 30)
MAX_RETRY = 3
REQUEST_GAP = 0.5   # 每个线程两次请求之间的间隔(秒), 对同花顺客气一点
MAX_WORKERS = 4     # 并发下载数; 实测单只约 2-3 秒, 4 线程 155 只约 2 分钟

PERIOD_HEADER = '科目\\时间'
HEADER_ROW, PERIOD_ROW, FIRST_STOCK_ROW = 1, 2, 3

# 模板数据块名 (第 1 行) → (同花顺科目名, 是否为百分比)
# 实测模板已有值 = 同花顺值 / 100, 净资产收益率用的是非摊薄口径
METRICS = {
    '营业总收入同比增长率': ('营业总收入同比增长率', True),
    '扣非净利润同比增长率': ('扣非净利润同比增长率', True),
    '基本每股收益':         ('基本每股收益(元)',     False),
    '净资产收益率':         ('净资产收益率',         True),
    '销售净利率':           ('销售净利率',           True),
}

# 模板里写错的代码 → 正确代码 (模板本身不改)
CODE_FIXES = {'68851': '688510'}   # 航亚科技, 模板少了一位

QUARTER_RE = re.compile(r'^(\d{4})Q([1-4])$')


# ---------------------------------------------------------------- 季度工具
def period_to_quarter(p: str) -> str:
    """'2026-09-30' → '2026Q3'"""
    return f'{p[:4]}Q{(int(p[5:7]) - 1) // 3 + 1}'


def quarter_key(q: str) -> int:
    m = QUARTER_RE.match(q)
    return int(m.group(1)) * 4 + int(m.group(2)) - 1


def key_to_quarter(k: int) -> str:
    return f'{k // 4}Q{k % 4 + 1}'


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


# ---------------------------------------------------------------- 模板读取
def normalize_code(v) -> str | None:
    if v is None:
        return None
    s = str(v).strip()
    if isinstance(v, (int, float)) or (s.isdigit() and len(s) < 6 and s not in CODE_FIXES):
        s = s.split('.')[0].zfill(6)
    s = CODE_FIXES.get(s, s)
    return s if re.fullmatch(r'\d{6}', s) else None


def read_template_stocks() -> list[dict]:
    """模板中每一个股票行: code, name, category(=sheet 名), row。同一代码可出现多行。"""
    if not os.path.exists(TEMPLATE_FILE):
        return []
    wb = openpyxl.load_workbook(TEMPLATE_FILE, read_only=True)
    out = []
    for ws in wb.worksheets:
        for r, row in enumerate(ws.iter_rows(min_row=FIRST_STOCK_ROW, max_col=2, values_only=True),
                                start=FIRST_STOCK_ROW):
            raw, name = row[0], row[1]
            if raw is None:
                continue
            code = normalize_code(raw)
            if code is None:
                print(f'[警告] {ws.title}!A{r} 代码无法识别: {raw!r}, 跳过')
                continue
            out.append({'code': code, 'name': (name or '').strip(), 'category': ws.title, 'row': r})
    wb.close()
    return out


def stock_summary() -> list[dict]:
    """页面用: 每个代码一条, 按模板顺序; 分类按出现的 sheet 合并。"""
    state = load_state()
    seen = {}
    for s in read_template_stocks():
        e = seen.get(s['code'])
        if e is None:
            seen[s['code']] = {'code': s['code'], 'name': s['name'], 'category': s['category'],
                               'latest_quarter': None}
            lp = state.get(s['code'], {}).get('latest_period')
            seen[s['code']]['latest_quarter'] = period_to_quarter(lp) if lp else None
        elif s['category'] not in e['category'].split(' / '):
            e['category'] += ' / ' + s['category']
    return list(seen.values())


# ---------------------------------------------------------------- 缓存与状态
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
            df = pd.read_excel(io.BytesIO(r.content), header=None, dtype=str, engine='xlrd',
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
                       index=[str(x).strip() for x in body.iloc[:, 0]], columns=periods)
    out.index.name = '科目'
    valid = [c for c in out.columns if re.fullmatch(r'\d{4}-\d{2}-\d{2}', c)]
    if not valid or out.empty:
        raise ValueError('报表格式不符: 没有报告期列')
    return out[sorted(valid, reverse=True)]


def parse_value(v, is_pct: bool):
    """'39.46%' → 0.3946 (百分比口径); '0.3841' → 0.3841; 空/'--'/False → None"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip().replace(',', '')
    if s in ('', 'nan', 'False', '--', '-'):
        return None
    pct = s.endswith('%')
    try:
        x = float(s.rstrip('%'))
    except ValueError:
        return None
    if is_pct:
        return round(x / 100, 8)
    return x / 100 if pct else x


def stock_quarter_values(code: str) -> dict:
    """{模板指标名: {'2026Q3': 值}}"""
    raw = load_raw(code)
    out = {m: {} for m in METRICS}
    if raw is None:
        return out
    for m, (src, is_pct) in METRICS.items():
        if src not in raw.index:
            continue
        for p in raw.columns:
            if p[5:] not in ('03-31', '06-30', '09-30', '12-31'):
                continue
            val = parse_value(raw.at[src, p], is_pct)
            if val is not None:
                out[m][period_to_quarter(p)] = val
    return out


# ---------------------------------------------------------------- 插列 (等价于 Excel 手工插列)
_CELL_RE = re.compile(r'^(\$?)([A-Z]{1,3})(\$?)(\d+)$')
_COL_RE = re.compile(r'^(\$?)([A-Z]{1,3})$')


def _shift_part(part: str, idx: int, n: int) -> str:
    m = _CELL_RE.match(part)
    if m:
        c = column_index_from_string(m.group(2))
        if c >= idx:
            c += n
        return f'{m.group(1)}{get_column_letter(c)}{m.group(3)}{m.group(4)}'
    m = _COL_RE.match(part)
    if m:
        c = column_index_from_string(m.group(2))
        if c >= idx:
            c += n
        return f'{m.group(1)}{get_column_letter(c)}'
    return part


def _shift_ref(ref: str, idx: int, n: int, sheet: str) -> str:
    if '!' in ref:
        prefix, body = ref.rsplit('!', 1)
        if prefix.strip("'") != sheet:
            return ref
        return prefix + '!' + ':'.join(_shift_part(p, idx, n) for p in body.split(':'))
    return ':'.join(_shift_part(p, idx, n) for p in ref.split(':'))


def shift_formula(formula: str, idx: int, n: int, sheet: str) -> str:
    tok = Tokenizer(formula)
    for t in tok.items:
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            t.value = _shift_ref(t.value, idx, n, sheet)
    return '=' + ''.join(t.value for t in tok.items)


def insert_columns(ws, idx: int, n: int, style_src_col: int) -> None:
    """在 idx 处插入 n 列, 平移公式/条件格式/列宽, 新列样式取自 style_src_col (插列前的列号)。"""
    if ws.merged_cells.ranges:
        raise RuntimeError(f'{ws.title}: 插列前不应存在合并单元格')
    ws.insert_cols(idx, n)
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith('='):
                c.value = shift_formula(c.value, idx, n, ws.title)
    src = style_src_col if style_src_col < idx else style_src_col + n
    for r in range(1, ws.max_row + 1):
        s = ws.cell(r, src)
        if s.has_style:
            for k in range(n):
                ws.cell(r, idx + k)._style = copy(s._style)
    # 列宽: 模板按区间 (min~max) 存。区间整体在插入点右侧 → 平移; 插入点落在区间内部 →
    # 区间扩展 (与 Excel 一致); 插入点紧贴区间右端 → 新列沿用左侧列宽, 单独成区间。
    spans = []
    for d in ws.column_dimensions.values():
        lo = d.min or column_index_from_string(d.index)
        hi = d.max or lo
        spans.append((lo, hi, d))
    new_dims = []
    covered = False
    for lo, hi, d in spans:
        if lo >= idx:
            lo, hi = lo + n, hi + n
        elif hi >= idx:
            hi, covered = hi + n, True
        new_dims.append((lo, min(hi, 16384), d))
    if not covered:
        w = next((d.width for lo, hi, d in spans if lo <= style_src_col <= hi and d.width), None)
        if w:
            new_dims.append((idx, idx + n - 1, None, w))
    ws.column_dimensions.clear()
    for item in new_dims:
        lo, hi, d = item[:3]
        letter = get_column_letter(lo)
        nd = copy(d) if d is not None else ws.column_dimensions[letter]
        if d is None:
            nd.width = item[3]
        nd.index, nd.min, nd.max = letter, lo, hi
        ws.column_dimensions[letter] = nd
    # 条件格式
    old = ws.conditional_formatting
    new = ConditionalFormattingList()
    for cf in old:
        sq = ' '.join(_shift_ref(str(rg), idx, n, ws.title) for rg in cf.sqref.ranges)
        for rule in cf.rules:
            new.add(sq, rule)
    ws.conditional_formatting = new


# ---------------------------------------------------------------- 模板数据块
def find_blocks(ws) -> list[dict]:
    """返回 METRICS 中各季度数据块: name, cols=[(列号, '2026Q3'), ...] (从左到右)。"""
    blocks, cur, name = [], None, None
    for c in range(1, ws.max_column + 1):
        h = ws.cell(HEADER_ROW, c).value
        p = ws.cell(PERIOD_ROW, c).value
        if h is not None and str(h).strip():
            name = str(h).strip()
        is_q = isinstance(p, str) and QUARTER_RE.match(p.strip())
        if cur and (not is_q or name != cur['name']):
            blocks.append(cur)
            cur = None
        if is_q and name in METRICS:
            if cur is None:
                cur = {'name': name, 'cols': []}
            cur['cols'].append((c, p.strip()))
        if p is None and h is None:
            name = None
    if cur:
        blocks.append(cur)
    return blocks


def extend_blocks(ws, latest_key: int) -> list[str]:
    """把每个数据块补到 latest_key 季度; 从右往左处理, 插列不影响左边数据块的列号。"""
    added = []
    for b in sorted(find_blocks(ws), key=lambda b: -b['cols'][-1][0]):
        last_col, last_q = b['cols'][-1]
        missing = [key_to_quarter(k) for k in range(quarter_key(last_q) + 1, latest_key + 1)]
        if not missing:
            continue
        idx = last_col + 1
        insert_columns(ws, idx, len(missing), style_src_col=last_col)
        for k, q in enumerate(missing):
            ws.cell(HEADER_ROW, idx + k).value = b['name']
            ws.cell(PERIOD_ROW, idx + k).value = q
        added.append(f"{b['name']} +{','.join(missing)}")
    return added


def merge_repeated_headers(ws) -> int:
    merged, c, maxc = 0, 1, ws.max_column
    while c <= maxc:
        v = ws.cell(HEADER_ROW, c).value
        e = c
        while v is not None and e + 1 <= maxc and ws.cell(HEADER_ROW, e + 1).value == v:
            e += 1
        if e > c:
            ws.merge_cells(start_row=HEADER_ROW, start_column=c, end_row=HEADER_ROW, end_column=e)
            ws.cell(HEADER_ROW, c).alignment = Alignment(horizontal='center', vertical='center')
            merged += 1
        c = e + 1
    return merged


def restore_threaded_comments(wb, path: str) -> int:
    """线程批注 openpyxl 存不下, 只剩一段占位提示; 从原文件取出真实内容写成普通批注。"""
    restored = 0
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        wb_rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        wb_xml = z.read('xl/workbook.xml').decode('utf-8')
        rid_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', wb_rels))
        rid_target.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="(rId\d+)"', wb_rels)})
        for sheet_name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb_xml):
            target = rid_target.get(rid, '').lstrip('/')
            target = target if target.startswith('xl/') else 'xl/' + target
            rels = target.replace('worksheets/', 'worksheets/_rels/') + '.rels'
            if rels not in names:
                continue
            m = re.search(r'Target="\.\./threadedComments/([^"]+)"', z.read(rels).decode('utf-8'))
            if not m:
                continue
            xml = z.read('xl/threadedComments/' + m.group(1)).decode('utf-8')
            texts = {}
            for ref, text in re.findall(r'<threadedComment [^>]*ref="(\w+)"[^>]*>\s*<text>(.*?)</text>',
                                        xml, re.S):
                texts.setdefault(ref, []).append(html.unescape(text))
            ws = wb[sheet_name]
            for ref, parts in texts.items():
                ws[ref].comment = Comment('\n'.join(parts), '批注')
                restored += 1
    return restored


# ---------------------------------------------------------------- 生成合并文件
def build_workbook(stocks: list[dict]) -> bool:
    wb = openpyxl.load_workbook(TEMPLATE_FILE)
    n_comments = restore_threaded_comments(wb, TEMPLATE_FILE)
    values = {s['code']: stock_quarter_values(s['code']) for s in stocks}
    by_sheet = {}
    for s in stocks:
        by_sheet.setdefault(s['category'], []).append(s)

    filled = kept = 0
    diffs = []
    for ws in wb.worksheets:
        rows = by_sheet.get(ws.title, [])
        keys = [quarter_key(q) for s in rows for m in METRICS for q in values[s['code']][m]]
        if keys:
            for a in extend_blocks(ws, max(keys)):
                print(f'  {ws.title}: 新增季度列 {a}')
        blocks = find_blocks(ws)
        for s in rows:
            vals = values[s['code']]
            for b in blocks:
                is_pct = METRICS[b['name']][1]
                for col, q in b['cols']:
                    v = vals[b['name']].get(q)
                    if v is None:
                        continue
                    cell = ws.cell(s['row'], col)
                    if cell.value is None or (isinstance(cell.value, str) and not cell.value.strip()):
                        cell.value = v
                        filled += 1
                    else:
                        kept += 1
                        if isinstance(cell.value, (int, float)):
                            tol = 0.00051 if is_pct else 0.0051
                            if abs(cell.value - v) > tol:
                                diffs.append((ws.title, cell.coordinate, s['name'], b['name'], q,
                                              cell.value, v))
        merge_repeated_headers(ws)

    wb.calculation.fullCalcOnLoad = True   # openpyxl 不保留公式缓存值, 让 Excel 打开时重算
    try:
        wb.save(OUTPUT_FILE)
    except PermissionError:
        print(f'[错误] {OUTPUT_FILE} 被占用 (是否在 Excel 中打开?), 本次未写入合并结果')
        return False
    print(f'合并完成: {os.path.basename(OUTPUT_FILE)} · 新填 {filled} 格 · '
          f'模板已有值保留 {kept} 格 (与同花顺不一致 {len(diffs)} 格) · 线程批注转普通批注 {n_comments} 条')
    for d in diffs[:15]:
        print(f'    差异 {d[0]}!{d[1]} {d[2]} {d[3]} {d[4]}: 模板 {d[5]} vs 同花顺 {d[6]}')
    return True


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser(description='拉取同花顺主要指标报表并填入模板')
    ap.add_argument('--force', action='store_true', help='忽略季度规则, 全部重新拉取')
    args = ap.parse_args()

    stocks = read_template_stocks()
    if not stocks:
        print(f'[提示] 未找到 {os.path.basename(TEMPLATE_FILE)} 或其中没有股票, 跳过基本面更新')
        return
    codes = list(dict.fromkeys(s['code'] for s in stocks))
    names = {s['code']: s['name'] for s in stocks}
    os.makedirs(RAW_DIR, exist_ok=True)

    today = date.today()
    expected = expected_period(today)
    state = load_state()
    print(f'股票 {len(codes)} 只 · 今天 {today} · 应有报告期 {expected}')

    todo = []
    for code in codes:
        info = state.setdefault(code, {})
        info['code'] = code
        info.pop('name', None)
        do, why = needs_fetch(info, expected, today, args.force)
        if do:
            todo.append((code, why))

    def worker(code):
        df = fetch_report(code)          # 失败抛异常, 由主线程记录
        time.sleep(REQUEST_GAP)
        return df

    # 下载并发, 但写缓存与状态只在主线程里做, 且每完成一只就落盘一次:
    # 进程中途被杀掉, 已完成的股票下次不会重拉; 失败的股票不记 last_checked, 下次照常重试。
    fetched = failed = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(worker, code): (code, why) for code, why in todo}
        for fut in as_completed(futures):
            code, why = futures[fut]
            info = state[code]
            label = f'{code} {names.get(code, "")}'.strip()
            try:
                df = fut.result()
            except Exception as e:
                failed += 1
                print(f'  [失败] {label}: {e} (保留旧缓存, 下次运行重试)')
                continue
            df.to_csv(raw_path(code), encoding='utf-8-sig')
            prev = info.get('latest_period')
            info['latest_period'] = df.columns[0]
            info['last_checked'] = today.isoformat()
            info['fetched_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            fetched += 1
            save_state(state)
            note = ' (尚未披露新一期)' if prev == df.columns[0] else ''
            print(f'  [{fetched + failed}/{len(todo)}] 拉取 {label}: {why} → 最新 {df.columns[0]}, '
                  f'{df.shape[1]} 期{note}', flush=True)
    # 模板里已删除的股票不再保留状态
    for code in [c for c in state if c not in codes]:
        del state[code]
    save_state(state)
    print(f'本次拉取 {fetched}, 失败 {failed}, 跳过 {len(codes) - fetched - failed}')

    build_workbook(stocks)


if __name__ == '__main__':
    main()
