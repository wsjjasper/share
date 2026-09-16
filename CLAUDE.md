# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A self-updating A-share (China equity) market sentiment indicator. Four daily micro-structure
sub-indicators from Wind All-A (万得全A, 881001.WI) — turnover rate (换手率), top-3 industry
share of traded value (成交额前三行业占比), share of advancing stocks (上涨个股占比), and margin
buying as a share of turnover (融资买入额占比) — are each converted to a rolling 252-trading-day
percentile rank, then equal-weighted into a 0–100 composite. Output range starts 2024-09-24.
The result is published as a static dashboard on GitHub Pages (https://wsjjasper.github.io/share/).

There is no test suite, no `requirements.txt`, and no build step. The "build" is running the pipeline.

## Commands

```bash
# Full pipeline: fetch → compute → chart → generate HTML → sync docs/ → git commit & push
python update.py

# Individual stages (each is standalone and writes its own outputs)
python auto_fetch_daily.py     # append/repair rows in 副本万得全A.xlsx
python sentiment_indicator.py  # 情绪指标_结果.xlsx + 情绪指标_图表.png
python generate_html.py        # index.html AND docs/index.html
```

Dependencies (matching CI): `pip install pandas numpy akshare requests matplotlib openpyxl`.
Chinese fonts must be present or the chart renders with tofu boxes — CI installs
`fonts-wqy-microhei fonts-noto-cjk`; `sentiment_indicator.py` tries SimHei / Microsoft YaHei /
WenQuanYi Micro Hei / Noto Sans CJK SC in that order.

`一键更新.bat` and `setup_docs.py` both hardcode `D:\vscode_workspace\情绪指标` and only work on
the author's Windows machine. `setup_docs.py` is one-off scaffolding that regenerates both
READMEs and `.gitignore` from inline strings — do not run it; edit the README files directly.

## Pipeline architecture

`update.py` is the orchestrator and the only thing CI invokes. It shells out to the three stage
scripts via `subprocess`, then does two things worth knowing about:

1. **Copies ~10 files from the repo root into `docs/`.** The root files are the working copies;
   `docs/` is the GitHub Pages publish directory and is a byte-identical mirror. Always edit at
   the root and let `update.py` (or `generate_html.py`, which writes both `index.html` paths)
   propagate. Editing anything under `docs/` directly will be silently overwritten.
   Despite what the root README says, root `index.html` is a full copy of the dashboard, not a redirect.
2. **Runs `git add . && git commit && git push` itself** (to `main` and `main:master`). The
   GitHub Actions workflow then pushes again. This is why the history has two identical commits
   per trading day. Any change to `update.py` should preserve that the script is expected to
   commit — but be aware that running it locally will commit your working tree.

`.github/workflows/daily_update.yml` runs `update.py` at 08:00, 12:00 and 15:00 UTC Mon–Fri
(Beijing 16:00 / 20:00 / 23:00) — post-close, after official margin data publishes, and a final
night check. The three runs exist so that estimated margin figures get overwritten with official
ones later the same day.

## Data model: positional columns

`副本万得全A.xlsx` is a raw Wind export whose headers contain embedded newlines and unit
annotations (see `col_info.txt` for the exact 14 columns). **Both `auto_fetch_daily.py` and
`sentiment_indicator.py` address columns by position, never by name.** `sentiment_indicator.py`
blindly reassigns `df.columns` to a 14-element list; `auto_fetch_daily.py` writes via
`df_excel.iloc[idx, N]`. Inserting, removing or reordering a column in the Excel silently
corrupts every downstream number. The positions that matter:

| idx | meaning |
| --- | --- |
| 1 | date (timestamped 16:00:00) |
| 2 | turnover rate, already % |
| 3 | top-3 industry share, **decimal 0–1** |
| 4 | advancing-stock share, **decimal 0–1** |
| 5 | margin-buy share, **decimal 0–1** |
| 6 | total market turnover, 亿元 |
| 12 | margin buy amount, 亿元 |
| 13 | margin-buy share, % |

Columns 3/4/5 are multiplied by 100 in `sentiment_indicator.py`; 2 is not. Margin data does not
exist before 2010-03-31, and a margin ratio of exactly `0` is coerced to `NaN` (treated as
"not yet disclosed"), so the composite is a mean over 3 sub-indicators on those days rather than 4.

## Percentile mechanics

`rolling_percentile_rank` uses a mid-rank formula — `(below + 0.5*equal) / n * 100` — over a
252-row window with `min_periods=126`, and returns `NaN` if fewer than half the window values are
valid. Because the window is in *rows*, it depends on the full history back to 2000 being present
in the Excel; the 2024-09-24 start date is only an output filter applied after the percentiles are
computed.

## Data fetching and its estimates

`auto_fetch_daily.py` has no Wind license. It reconstructs rows from free sources:
akshare (`stock_zh_index_daily` for the trading calendar, `stock_margin_account_info` for official
margin data), Tencent `qt.gtimg.cn` for aggregate turnover, and Eastmoney `push2delay` for
industry concentration. Two things follow:

- **New rows are partly hardcoded placeholders.** When appending a fresh trading day it writes
  literal constants for turnover rate (1.48), advancing-stock share (0.585 / 58.50) and advancing
  count (3240) because no free source is wired up for them. Only turnover, industry concentration
  and margin figures are actually fetched. Each fetcher also has a hardcoded fallback return
  (18220.0 亿 turnover; 0.435 / 7950.0 industry) used when the HTTP call fails, so a network
  failure produces plausible-looking wrong data rather than an error.
- **Self-repair pass.** Before appending, it walks every existing row and overwrites columns
  5/12/13 from the official akshare margin table whenever the stored value is NaN, 0, or differs
  by more than 1.0 亿. This is what makes the same-day estimate converge to the official number
  across the three daily CI runs.

The Eastmoney URL deliberately uses `t:1` (31 primary industries). Switching to `t:2` yields ~100
sub-sectors and halves the top-3 ratio, silently breaking comparability with the Wind history.

## `_latest.xlsx` fallback trap

Every Excel write is wrapped in a `PermissionError` handler (the author keeps the files open in
Excel on Windows) that diverts the output to a `*_latest.xlsx` sibling. Readers prefer the
*primary* file and only fall back to `_latest` if the primary is missing. So when a fallback
fires, the next stage keeps reading stale data and nothing warns you. `情绪指标_结果_latest.xlsx`
is committed at the repo root as a leftover of exactly this; it is not read by anything as long as
`情绪指标_结果.xlsx` exists.

## Editing the dashboard

`generate_html.py` is a ~1000-line Python file that is almost entirely one raw string
(`html_content = r"""..."""`) holding the complete Tailwind + ECharts page. Data is injected by
two literal `str.replace` calls at the bottom, on the sentinels `DATA_JSON_PLACEHOLDER` and
`LATEST_DATE`. Consequences: HTML/CSS/JS edits happen inside that Python string; the page ships
no external data file (`data_records.json` at the repo root and in `docs/` is a stale artifact of
`setup_docs.py`, not a runtime input); and the browser-side `rawData` array uses English keys
(`turnover`, `top3_ind`, `rise_pct`, `margin_pct`, `pct_*`, `composite_sentiment`) assigned
positionally in `generate_html.py`, which is a third place where column order is load-bearing.

## Thresholds

`> 80` overheated (过热), `20–80` neutral (中性), `< 20` oversold (过冷). These appear in the
matplotlib chart, the HTML dashboard and `update.py`'s console summary — change all three together.
