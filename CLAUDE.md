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

`auto_fetch_daily.py` has no Wind license. It reconstructs rows from free sources, all via
akshare except one: `stock_zh_index_daily` for the trading calendar, `stock_margin_account_info`
for official margin data, `index_analysis_daily_sw` for industry concentration, and Tencent
`qt.gtimg.cn` for aggregate turnover. Two things follow:

- **Missing inputs are left `NaN`, never filled with a constant.** Every sub-indicator on a new
  row is now fetched. The one hardcoded fallback left is `fetch_market_turnover_and_breadth`,
  which returns 18220.0 亿 when its HTTP call fails — a network failure there still produces
  plausible-looking wrong data rather than an error. (Rows appended before 2026-09-16 carry
  literal constants — turnover rate 1.48, advancing share 0.585 / 58.50, advancing count 3240 —
  from when no source was wired up; `backfill_industry_sw.py` replaces the turnover ones.)
- **Self-repair pass.** Before appending, it walks every existing row and overwrites columns
  5/12/13 from the official akshare margin table whenever the stored value is NaN, 0, or differs
  by more than 1.0 亿. This is what makes the same-day estimate converge to the official number
  across the three daily CI runs.

Two of the four sub-indicators come from 申万宏源 through a single call.
`fetch_sw_daily_metrics(start, end)` wraps akshare's
`index_analysis_daily_sw(symbol="一级行业", start_date, end_date)`, which returns one row per
industry per day over a date range, and derives both:

- **Industry concentration** — the response carries a `成交额占比` column, so the top-3 share is a
  sum of three numbers; no denominator to assemble and no per-board request fan-out.
- **Whole-market turnover rate** — the same rows carry `换手率` and `流通市值`, and the 31 primary
  industries partition the whole A-share market, so
  `Σ(换手率ᵢ × 流通市值ᵢ) / Σ(流通市值ᵢ)` is algebraically `Σ(成交额ᵢ) / Σ(流通市值ᵢ)`, the
  market-wide turnover rate. It costs no extra request and inherits the same history and repair
  path as the industry column. Its unit is auto-detected too: a median outside 0.001–20 is
  rejected, and a median below 0.2 is treated as a decimal and scaled by 100.

**The 申万 fetch does not go through akshare.** `index_analysis_daily_sw` used to hang forever.
Measured root cause (2026-09-16): TCP to `www.swsresearch.com` connects in 0.2s every time, but the
**TLS handshake succeeds only about 1 time in 4** — the rest stall indefinitely. akshare opens a
*new connection per page* via module-level `requests.get` and passes no `timeout` anywhere, so a
9-month range is ~110 independent handshakes and any one of them blocks the whole script with no
error. It was never rate limiting.

`_sw_fetch_raw` therefore calls the endpoint directly (`SW_API_URL`), reusing one module-level
keep-alive `requests.Session` (`_sw_session`) for the entire process: the handshake is gambled on
**once**, and every page afterwards rides the established connection at 0.4–0.5s. `SW_PAGE_SIZE` is
200 rather than akshare's 50, cutting page count by 4×. `_sw_get_page` retries `SW_MAX_RETRY` (8)
times per page with a short `SW_FETCH_TIMEOUT` ((4, 30)) — the connect timeout is deliberately small
because a handshake either lands in 0.2s or never, so failing fast and re-rolling is cheaper than
waiting. Measured effect: a 12-trading-day range went from hanging forever to **1.3s**, and 5
consecutive pages ran with 0 failures. `verify=False` is required (the cert chain does not validate
locally, same as akshare), so `urllib3.disable_warnings` is called at import.

Responses are renamed to akshare's Chinese column names via `SW_FIELD_MAP` (copied from
`akshare/index/index_research_sw.py`), so everything downstream is unchanged. `fetch_sw_daily_metrics`
still splits the range into `SW_CHUNK_DAYS` (90) day chunks for progress visibility, concatenating
and de-duplicating on `(发布日期, 指数名称)`: an overlap would double a day's `成交额占比` sum and
break the unit auto-detection below.

`_request_timeout` remains, but now guards the *other* akshare calls (`stock_zh_index_daily`,
`stock_margin_account_info`, `stock_zh_a_spot_em`) with `AK_FETCH_TIMEOUT` ((10, 60)) — they have the
same no-timeout flaw and the pipeline runs unattended in CI. It patches `requests.Session.request` to
fill in a timeout when a caller passes none and restores it on exit. **`socket.setdefaulttimeout` does
not work for this** and was removed: urllib3 runs `sock.settimeout(self.timeout)` unconditionally
after connecting (`connection.py:439/560`), and with requests passing `timeout=None` that call
explicitly returns the socket to blocking mode, overriding the global default. Only a patch at the
requests layer holds.

Three properties matter:

- **It is NOT the same caliber as the Wind history.** This was assumed and is now measured false.
  Over the 120-trading-day overlap (2026-03-04 ~ 2026-08-25), 申万 runs systematically high on both
  derived columns, so its values must **not** be spliced into the pre-2026-08-26 rows:

  | column | 申万 mean | Wind mean | bias | 申万/Wind | corr |
  | --- | --- | --- | --- | --- | --- |
  | 行业集中度 | 51.88% | 43.84% | +8.04pp | 1.19–1.23× | 0.9918 |
  | 换手率 | 5.03% | 1.87% | +3.16pp | 2.57–2.81× | 0.8654 |

  The correlations are high — it tracks the same signal — but the level is on a different basis
  (申万's turnover rate is almost certainly free-float-denominated, and its `成交额占比` is
  normalized within 申万 coverage rather than over the whole market). Because each sub-indicator is
  a **rolling 252-day percentile of its own column**, a level shift is not a cosmetic offset: it
  pegs the new days at the top of their own window. Measured on the 6 rows that reached production,
  the industry percentile was inflated by **+35 to +61 points**; writing 申万 turnover unadjusted
  produces a percentile of **99.8**. Reconstructing a Wind-caliber figure from 申万 absolutes does
  not work either — `成交量 × 均价` overshoots Wind's market turnover by ~1.4× because `均价` is an
  index-level mean price.
- **It is queried by date range**, so each trading day gets its own figure. The previous Eastmoney
  `clist` endpoint returned only a current snapshot, and the fetch sat outside the date loop, so
  one value was written to every date appended in that run (2026-08-27 and 08-28 in the committed
  history are identical for exactly this reason).
- **It never substitutes a constant.** If 申万 has not published yet, or returns fewer than
  `SW_MIN_INDUSTRIES` (25) rows for a day, or its `成交额占比` sums to neither ~1 nor ~100 per day
  (the unit auto-detection), the fetch returns nothing and the row is left `NaN`. A repair pass in
  `auto_sync_and_append` then refills any `NaN` industry or turnover-rate cell dated within
  `SW_REPAIR_LOOKBACK_DAYS` (7) on a later run — the same shape as the margin self-repair, and the
  reason the three daily CI runs matter for these columns too.

**Advancing-stock share has no historical source.** `fetch_market_breadth` counts
`涨跌幅 > 0` over `stock_zh_a_spot_em`, a whole-market snapshot with no date parameter — nothing in
akshare returns market breadth for a past date. Two consequences: the value is only written when
the date being appended is the newest closed trading day (`all_trade_dates[-1]`), so a run
catching up on several days leaves the older ones `NaN` rather than repeating one figure across
them; and the repair pass cannot fill it, so a trading day missed by all three CI runs stays
`NaN` permanently. The fetch also rejects a snapshot carrying fewer than 3000 stocks.

**Caliber history of this column.** Rows through 2026-08-25 are Wind's own industry figures
(mean 36.76%). 2026-08-26 to 2026-09-16 were fetched from Eastmoney `fs=m:90+t:1`, the
**geographic** board pool (31 provinces) rather than industries — the province count coincides
exactly with 申万一级's 31 sectors and the ratio (40–43%) sat inside the Wind range (25–52%), so
the series looked continuous while measuring geography. Those 16 rows have been voided to `NaN`.
Eastmoney's pools, for reference if anyone reaches for them again: `t:1` geographic, `t:2`
industry (~86 boards, top-3 ≈ 20%, far finer than 申万一级), `t:3` concept.

## Backfilling history: `backfill_industry_sw.py`

A one-off script for filling stretches of the industry-concentration and turnover-rate columns
older than the `SW_REPAIR_LOOKBACK_DAYS` window the daily pipeline covers. It imports
`fetch_sw_daily_metrics` from `auto_fetch_daily.py` rather than keeping its own copy, so both read
the same source through the same code.

It refuses to write unless **both** columns agree with the Wind history: per-column thresholds
live in the `CHECKS` list at the top of the file (industry ≤ 3 percentage points mean absolute
deviation, turnover rate ≤ 0.5, correlation ≥ 0.80 for both), and any column failing rejects the
whole write. The comparison window covers only rows dated **before** `--start`: inside the
backfill range the Excel holds either voided `NaN`s or the 1.48 turnover placeholder, and letting
those into the window both breaks a legitimate check and could pass a bad one when a placeholder
happens to sit near the real value. Run it with no flags to validate and preview; `--apply`
writes. Column 7 (前三行业合计) is derived as ratio × column 6, since the source gives only the
share.

Run its validation before trusting the daily path: the pipeline's repair pass fills recent `NaN`
industry cells automatically and has no validation gate of its own.



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
