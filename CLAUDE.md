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

# Repair scripts for historical gaps the daily path cannot reach. All three default to a
# dry run that validates and previews; --apply writes. Each has its own section below.
python backfill_industry_sw.py   # 行业集中度 (col 3/7) — 申万, seconds
python backfill_breadth_sina.py  # 上涨个股占比 (col 4/8/9/11) — 逐只日线, ~26 min
python rebuild_turnover_sw.py    # 换手率 (col 2) — 整列重建, --start 2015-01-01
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

- **It is the same caliber as the Wind history — but only if you normalize per day.**
  申万's `成交额占比` does **not** reliably sum to 100. Measured across 175 trading days: the
  31 shares sum to **110–123** from 2025-12 through 2026-09-08, then switch to exactly **100**
  from **2026-09-09** onward. The original code picked one `ratio_scale` for the whole fetch from
  the median daily sum and divided everything by 100, so every pre-09-09 value came out ~18% too
  high. Against Wind that looked like a caliber difference (+7.23pp bias) when it was just a wrong
  denominator. `fetch_sw_daily_metrics` now computes `top3 / (that day's own sum)`, which is
  immune to both the regime change and the decimal-vs-percent question:

  | 算法 | 申万 mean | Wind mean | bias | MAD | corr | gate |
  | --- | --- | --- | --- | --- | --- | --- |
  | 全局 ÷100 (旧) | 49.08% | 41.85% | +7.23pp | 7.23 | 0.9948 | 未通过 |
  | 按日归一 (现) | 42.27% | 41.85% | +0.42pp | **0.45** | **0.9981** | **通过** |

  Rows written on/after 2026-09-09 are unaffected (their day sum was already 100), so the fix
  changes no value that ever reached production. **Never reintroduce a global `÷100`.**

- **The turnover rate derived from it is a different caliber from Wind's — deliberately.** 申万's
  `流通市值` is a **free-float** figure, so the derived rate runs ~2.5x Wind's old column. That is
  not a bug to be corrected: as of 2026-09-16 the whole 换手率 column from 2015-01-01 onward has
  been **rebuilt** in this caliber by `rebuild_turnover_sw.py`, because Wind's original column
  could not be reproduced from any free source (see that section). `SW_TURNOVER_ENABLED` is `True`
  so the daily pipeline keeps writing the same caliber. Its unit is auto-detected: a median outside
  0.001–20 is rejected, and a median below 0.2 is treated as a decimal and scaled by 100.

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

**Advancing-stock share is snapshot-only on the daily path.** No free API returns market breadth
for a past date, so `fetch_market_breadth` counts `涨跌幅 > 0` over a whole-market snapshot. It is
only written when the date being appended is the newest closed trading day (`all_trade_dates[-1]`),
so a run catching up on several days leaves the older ones `NaN` rather than repeating one figure.
A day missed by all three CI runs is no longer lost, though — `backfill_breadth_sina.py` can
reconstruct it (see below).

Three defenses, all added after real failures:

- **Retries.** It used to return `None` on the first exception. Because the snapshot has no
  history, one transient disconnect meant that day was `NaN` forever. Now `BREADTH_MAX_RETRY` (3)
  attempts with exponential backoff.
- **A second host.** When Eastmoney rate-limits (it will — it started refusing connections outright
  after a burst of testing), `_breadth_sina` falls back to the 新浪 list endpoint, which carries
  `changepercent` per stock across ~56 pages of 100 in about 18s and covers 北交所 too.
- **A zero-snapshot guard, which matters more than it looks.** Before the next session opens, the
  新浪 list resets every `trade` and `changepercent` to `0.000` while `settlement` still holds the
  prior close; Eastmoney behaves the same way pre-open. That snapshot has ~5500 rows, so it sails
  past the `BREADTH_MIN_STOCKS` (3000) check and yields **0 advancing = 0.00% breadth**, which is
  a plausible-looking number that is catastrophically wrong. `BREADTH_MIN_NONZERO` (0.5) requires
  that at least half the stocks show a non-zero move; a real trading day is near 100%. This hole
  existed in the original Eastmoney-only code too and simply had not been hit yet.

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

Each column is gated **independently**: per-column thresholds live in the `CHECKS` list at the top
of the file (industry ≤ 3 percentage points mean absolute deviation, turnover rate ≤ 0.5,
correlation ≥ 0.80 for both). A column that passes is written; a column that fails is left
untouched, and the run only aborts if *nothing* passes. It used to be all-or-nothing, which meant
the permanently-failing turnover column blocked the industry backfill from ever landing. `validate()`
returns a `{key: bool}` map; note `ok` is wrapped in `bool()` because `mad`/`corr` are numpy scalars
and `np.True_ is True` is `False`, which silently broke the summary line. The comparison window covers only rows dated **before** `--start`: inside the
backfill range the Excel holds either voided `NaN`s or the 1.48 turnover placeholder, and letting
those into the window both breaks a legitimate check and could pass a bad one when a placeholder
happens to sit near the real value. Run it with no flags to validate and preview; `--apply`
writes. Column 7 (前三行业合计) is derived as ratio × column 6, since the source gives only the
share.

Run its validation before trusting the daily path: the pipeline's repair pass fills recent `NaN`
industry cells automatically and has no validation gate of its own.

As of 2026-09-16 the 16 rows 2026-08-26 ~ 2026-09-16 have had their industry column backfilled and
validated (MAD 0.47pp, corr 0.9968). The turnover column in those same rows still holds the literal
placeholder `1.48` — the caliber gate rejects 申万 for it, so the backfill skipped it. Careful:
`1.48` also occurs legitimately in 45 rows back to 2002, so anything that voids the placeholder must
select by **date range**, not by value. The 上涨个股占比 column in those rows likewise still holds
the placeholder `0.585`, which no source can repair (see the breadth note above).



## Rebuilding breadth: `backfill_breadth_sina.py`

Reconstructs 上涨个股占比 for past dates by counting, stock by stock, how many closed above their
previous close. It exists because nothing returns breadth by date: the exchanges' own summaries
carry turnover, market cap and turnover rate but no advance/decline counts; every date-parameterized
akshare function is either unrelated or a 涨停板 pool; 乐咕's 赚钱效应 is a snapshot whose page
structure changed (the function raises); and guessing at 乐咕 JSON API paths returns 404.

The cost is driven by the **number of stocks, not the number of dates** — each request returns a
whole date range — so `DATALEN` (60) is nearly free and is set to leave ~44 trading days outside the
backfill range to validate against. The universe comes from the 新浪 list endpoint (~5560 including
北交所) and the bars from 新浪 `CN_MarketData.getKLineData` at ~0.28s each, so a full pass is about
26 minutes. Use 新浪, not Eastmoney: Eastmoney throttles a burst of this size within a few hundred
requests and then closes connections. Only per-date counters are cached (`_breadth_cache.tmp`,
gitignored), never the bars, and the processed-symbol list makes the run resumable.

Measured against the Wind rows: MAD **0.09pp**, max deviation 0.46pp, correlation **1.0000** over 41
trading days. It reproduces the column almost exactly.

**The validation window must exclude the backfill range** (`--start`), the same rule
`backfill_industry_sw.py` follows. Leaving it in cost real accuracy here: 2026-08-24 and 2026-08-25
are themselves bad rows, and with them inside the window MAD read 1.22pp and correlation 0.9742 —
a flawless reconstruction looked merely adequate.

**How those two rows were caught, and why it generalizes.** 2026-08-21 and 2026-08-24 carry
*identical* values in three columns at once — 上涨占比 45.134501, 上涨家数 2500, 成份数 5539 — which is
the same "one snapshot written to several dates" bug already recorded for 2026-08-27/08-28. Hence
the default `--start` is 2026-08-24, not 2026-08-26, and 18 rows were rewritten rather than 16.
**Duplicate values across adjacent dates are the signature of this whole class of bug**; when a
backfill disagrees with stored data, check whether the stored row is a copy of its neighbour before
assuming the new source is wrong.

## Rebuilding the turnover column: `rebuild_turnover_sw.py`

A one-off script that overwrote column 2 for every row from 2015-01-01 on. It exists because
**Wind's original 换手率 column cannot be reconstructed from any free source.** Measured over 51
trading days against the exchanges' own published summaries (`stock_sse_deal_daily(date)` for 沪
and `stock_szse_summary(date)` for 深, which publish 成交金额 and 流通市值 by date):

| pair | corr |
| --- | --- |
| Wind 成交额 vs exchange 成交额 | **1.0000** |
| exchange 换手率 vs exchange 成交额 | 0.9876 |
| Wind 换手率 vs exchange 换手率 | 0.6770 |
| Wind 换手率 vs **Wind's own 成交额** | **0.6810** |

The 成交额 columns agree perfectly, so the disagreement is entirely in the denominator — and Wind's
implied denominator moves a median of **4.05% per day** where the real float market cap moves 1.26%.
The column carries ~12% idiosyncratic variation relative to any real turnover rate, so no
`成交额 ÷ 市值` construction reproduces it. Do not spend time looking again; 乐咕
(`stock_a_congestion_lg` is 拥挤度 with NaN recent values, `stock_market_activity_legu` is broken)
and 申万市场表征 were checked too.

The way out is that the composite consumes **percentile ranks within each column's own rolling
252-day window**, so a column only needs to be self-consistent — it does not need to match Wind.
The rebuilt series is a genuine turnover rate: per-year correlation with 成交额 of 0.9327 (worst
year 0.8232) versus 0.7166 for the column it replaced.

Range matters: this endpoint returns **nothing usable before 2014** (rows exist but carry fewer
than `SW_MIN_INDUSTRIES`), 87% of 2014, and 98.8–100% from 2015. A full-history run therefore fails
its own coverage gate at 47% — use `--start 2015-01-01`, which scores 99.40%. The splice at
2015-01-01 is harmless because the earliest output (2024-09-24) looks back only 252 rows, to
2023-09-08. Gap days are written as `NaN` rather than left at the old caliber: a mixed-caliber
column is worse than a missing cell.

Its gate is self-consistency, not agreement with Wind: coverage ≥ 95%, every value within
0.2–20%, per-year correlation with 成交额 ≥ 0.85, and implied-denominator median daily move ≤ 3%.
**Correlation must be computed per year and then medianed** — 成交额 is a level and 换手率 a ratio
whose denominator grew two orders of magnitude, so a single correlation across all history is
meaningless (Wind's own column scores 0.169 across 26 years but 0.72 within a year). Results are
cached in `_sw_turnover_cache.tmp` (gitignored via `*.tmp`) so validate and `--apply` do not refetch;
`--refetch` forces a re-pull.
## The US market tab

The dashboard carries two markets in one page. **There is only one set of DOM elements**: the tab
switch swaps the underlying dataset and the labels, it does not duplicate the markup. The four
sub-indicator keys are deliberately identical across markets (`turnover`, `top3_ind`, `rise_pct`,
`margin_pct` and their `pct_*` twins), so `updateDashboard`, the table and the chart builders never
branch on market — what a key *means* comes from the `MARKETS` config in the page, which supplies
per-market titles, card labels, chart series names, axis names and the raw-value unit suffix
(`%` for A-shares, none for the US since its values carry mixed units).

`generate_html.py` injects two arrays into `const DATASETS`. The US one is read from
`us_sentiment_result.csv` when that file exists and is `[]` otherwise, which makes the tab render an
explicit "尚无数据" empty state — `switchMarket` hides every `main > section` and shows the
placeholder. Nothing is ever filled with example or estimated numbers.

**The two injection sentinels must not be substrings of one another.** The original pair was
`DATA_JSON_PLACEHOLDER` and `US_DATA_JSON_PLACEHOLDER`; the first `str.replace` rewrote the tail of
the second, leaving `us: US_[{…A-share data…}]` — wrong data *and* a ReferenceError that killed the
whole script. The US slot is now `US_DATA_JSON_SLOT`.

`populateTable()` clears `tbody` before filling it; without that, every tab switch appended another
15 rows.

**`sec.hidden = true` is not enough to hide a section.** Tailwind's display utilities beat the UA
stylesheet's `[hidden] { display: none }`, so the 「四大微观情绪子指标」section — the one carrying
`class="grid ..."` — kept rendering on an empty US tab with the attribute set (`hidden` true,
`display: grid`). The other six sections have no display class, which is why only that one leaked.
`switchMarket` now sets `sec.style.display` as well; an inline style outranks the class.

**Everything market-specific has to come from `MARKETS`, not the static markup.** Hiding a section
only helps while a market is empty; once its data lands the section renders again, A-share wording
and all. `MARKETS` therefore also carries `docTitle`, `footerSystem`, `footerSource` and a `method`
block (intro prose, the composite formula, and the four sub-indicator table rows), and
`switchMarket` rebuilds the 子指标 table, the 明细表 header row (from `subs[i].raw` and
`subs[i].label`) and the raw-series chart title (from `axisLeft`/`axisRight`). The footer sits
*outside* `main`, so it is never covered by the section hiding and has to be switched explicitly.
Two sentences that named 换手率 in passing were reworded to be market-neutral instead of being
duplicated per market. The check that matters: on the US tab, no visible text may mention 万得全A,
881001, 换手率, 融资买入 or 行业.

### US pipeline

`us_fetch_daily.py` → `us_market_data.csv` → `us_sentiment_indicator.py` → `us_sentiment_result.csv`.
Needs `yfinance` on top of the A-share dependencies. **It is not wired into `update.py` or CI**, and
CI does not install `yfinance`. Run the two scripts by hand when the US numbers should move.

**The US stages now run in `update.py`, deliberately without a return-code check.** A-shares are the
main line; the US tab is an addition riding on Yahoo, and one network blip there must not fail the
whole daily run. On failure the step warns, skips the second script and continues — `generate_html.py`
then republishes the page from the committed CSV, i.e. the previous day's US numbers. CI installs
`yfinance` alongside the A-share dependencies, and `us_sentiment_result.csv` is in the `docs/` sync
list because the page's download button links to it.

**Both CSVs are committed, and that is load-bearing.** `generate_html.py` injects `[]` when
`us_sentiment_result.csv` is absent, and CI regenerates the page three times every trading day — so
an uncommitted CSV means every CI run silently republishes the dashboard with an empty US tab, which
is exactly what happened on the first attempt. The A-share data files are committed for the same
reason. The US CSVs are not in `.gitignore`, so `update.py`'s `git add .` keeps them.

First live run (2026-09-17) fetched 1005 trading days back to 2022-09-19 and produced 880 rows of
output (the first 252 are consumed by the rolling window). Two fixes came out of it:

- `close.pct_change()` defaulted to `fill_method='pad'`, which forward-fills a halted day to the
  prior close: the stock then scores a 0% change, stays in the denominator and counts as *not*
  advancing, biasing breadth down. It is now `pct_change(fill_method=None)` so halted days drop out
  of numerator and denominator alike — the same rule `backfill_breadth_sina.py` follows. It also
  silences a pandas deprecation that would eventually have become an error.
- Neither US script had the Windows UTF-8 stdout guard the A-share scripts carry, so
  `us_sentiment_indicator.py` wrote its CSV and *then* died with `UnicodeEncodeError` on the `•` in
  its summary — exit code 1 with the data already on disk. Wiring that into `update.py`, which
  aborts on a non-zero return, would have failed the whole run for a cosmetic print.

`UNIVERSE` still lists `BK`, which Yahoo 404s. 101 of 102 resolve, well above `MIN_UNIVERSE` (60),
so the run proceeds — but it is a stale ticker, not a transient error.

**Holiday rows, and why `MIN_SUBS` exists.** On US market holidays yfinance still returns a `^VIX`
row while the sector ETFs and the constituent list return nothing, so the outer merge invents a row
carrying VIX alone. `out[pct_cols].mean(axis=1)` skips NaN, so that row produced a "composite" that
was really just the VIX percentile wearing a four-indicator label — 2026-05-25 (Memorial Day) scored
64.48 and 2026-09-07 (Labor Day) 85.91 that way. `us_sentiment_indicator.py` now requires
`MIN_SUBS` (3) of the four percentiles before it will emit a composite, and prints the dates it
drops. 3 rather than 4 matches A-shares, where margin data legitimately does not exist before
2010-03-31. Output went 880 → 878 rows. Note both dates *are* valid A-share sessions, so they
still appear in the `cn` dataset — finding a date in the page is not evidence it survived here.

**⚠ 板块集中度 is the weakest of the four and is knowingly shipped as-is.** It is the top-3 share of
just 11 SPDR ETFs, which bounds it from below at 3/11 = 27.3% and squeezes the whole 879-day history
into 35.5–60.1% with a standard deviation of 3.47pp. Feed a distribution that tight into a rolling
percentile and the percentile amplifies noise: a 10pp move in the raw value sweeps the entire
historical range, so 2026-09-15 → 09-16 went 42.89% → 53.29% and the percentile went 14.6 → 96.2.
101 of 879 days move the percentile by more than 50 points. Compare the coefficient of variation
(std ÷ mean) across the four:

| sub-indicator | CV | median daily percentile move | days moving >50 |
| --- | --- | --- | --- |
| 板块集中度 | **0.076** | 18.7 | 101 |
| 成交额 | 0.355 | 14.3 | 42 |
| 上涨占比 | 0.388 | 26.4 | 179 |
| VIX | 0.243 | 5.0 | 1 |

上涨占比 also jumps often, but that is genuine — breadth really does swing from 8% to 94%, exactly as
it does for A-shares. 板块集中度 is different: the series barely moves, so roughly a quarter of the US
composite is close to noise. The A-share analogue avoids this by partitioning the market into 31
申万一级 industries (floor 3/31 = 9.7%, observed 25–52%). Two further caveats on the same block of
data: ETF trading volume reflects institutional hedging and creation/redemption flows rather than
how actively a sector's underlying stocks trade, and the 11 ETFs sum to a median of 10.9 十亿美元
against a real US tape of roughly 400–600 十亿, so the card labelled 「全市场成交额」 is measuring
about 2% of the market. The fix, when someone wants it, is to aggregate the already-downloaded
`UNIVERSE` constituents' dollar volume by GICS sector instead of using ETF volume.

The four dimensions, and why they differ from the A-share set:

| A-share | US | source |
| --- | --- | --- |
| 换手率 | total dollar volume (十亿美元) | the 11 SPDR sector ETFs, summed |
| 成交额前三行业占比 | top-3 sector share (%) | same ETFs — one download feeds both, so they are self-consistent |
| 上涨个股占比 | advancing share (%) | `UNIVERSE` in `us_fetch_daily.py`, an S&P 100 list — a **mega-cap** proxy, narrower than the A-share metric |
| 融资买入额占比 | VIX | `^VIX`; FINRA margin debt is monthly, so there is no daily equivalent |

**VIX enters inverted.** High VIX means fear means low sentiment, the opposite direction from the
other three, so `us_sentiment_indicator.py` stores `100 - percentile(VIX)` in `pct_margin_pct` while
`margin_pct` keeps the raw VIX for display. Anything comparing the two markets' composites has to
account for that substitution — they are not the same index computed on different data.

`sentiment_core.py` holds `rolling_percentile_rank`; both `sentiment_indicator.py` and
`us_sentiment_indicator.py` import it so the percentile math cannot drift between markets.
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
