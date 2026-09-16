# HANDOVER — 申万数据回填未完成

写给接手的本地 session。上一段工作在**没有外网**的沙箱里进行：申万
(`www.swsresearch.com`)、东财 (`push2delay/push2his.eastmoney.com`)、乐咕
(`legulegu.com`) 全部被出口策略拦截，返回 `000`。因此**所有代码只经过构造数据的
离线验证，真实接口一次都没跑通过**。你有网，第一件事是验证这些假设。

先读 `CLAUDE.md`，它有完整的管线结构、列位置约定和历史坑位。本文件只讲未完成的事。

---

## 1. 当前卡点：申万抓取永久挂起

`python backfill_industry_sw.py --apply` 卡在抓取阶段，长时间无响应。

### 根因（已确认）

`akshare.index_analysis_daily_sw` 内部按每页 50 条**串行翻页**，9 个月区间约 109 次
连续请求，而 `akshare/index/index_research_sw.py` 全文 `timeout=` 出现 **0 次** —— 每个
`requests.get` 都不带超时。服务端一限流或挂起就永久阻塞。

### 已修：超时现在真的生效了

最初 (`46ffd67`) 我用 `socket.setdefaulttimeout(30)` 兜底，**那是无效的**，已删除。
原因：urllib3 在连接建立后无条件执行

```
urllib3/connection.py:439  self.sock.settimeout(self.timeout)
urllib3/connection.py:560  self.sock.settimeout(self.timeout)
```

`requests.get()` 不传 timeout 时 `self.timeout` 就是 `None`，于是 `sock.settimeout(None)`
**显式把 socket 设回永久阻塞，覆盖掉全局默认值**。（环境：urllib3 2.6.3 / requests 2.33.1）

现在改成 `_request_timeout` 上下文管理器，拦在 `requests.Session.request` 这一层注入
`SW_FETCH_TIMEOUT = (10, 30)`，退出时还原，作用域不外溢到管线里其它 akshare 调用。

**这是本次唯一走真实 requests 栈验证过的代码**（在 `Session.send` 记录实际收到的 timeout）：
无 timeout 的调用被注入 `(10,30)`、显式传 `timeout=5` 的不被覆盖、正常与异常路径都正确还原。
分块 (45天/块) 与重试 (3次, 2s/4s 退避) 保留不变。

### 重跑后会看到什么

超时只是把"永久挂起"变成"**报错退出**"，不等于抓取就能成功。两种可能：

- **跑通** → 之前只是偶发卡在某一页，问题解决
- **每块重试 3 次后全部超时** → 申万在限流或拒绝连接，走下面的出路

后者才是真没解决。**别把"现在会报错了"当成修好了。**

### 若仍失败，两条出路

**A. 绕开 akshare，直连申万**（推荐）

端点和字段我已从 akshare 源码读出，不用猜：

```
GET https://www.swsresearch.com/institute-sw/api/index_analysis/index_analysis_report/
params: page=1, page_size=50, index_type=一级行业, swindexcode=all, type=DAY,
        start_date=YYYY-MM-DD, end_date=YYYY-MM-DD
headers: User-Agent: Mozilla/5.0 ...
verify=False          # akshare 原样如此
```

响应 `data.count` 是总条数，`data.results` 是当页数组。字段改名映射：

| 原始字段 | 含义 |
|---|---|
| `bargaindate` | 发布日期 |
| `swindexname` | 指数名称 |
| `bargainsumrate` | 成交额占比 |
| `turnoverrate` | 换手率 |
| `negotiablessharesum1` | 流通市值 |

自己控制 timeout、重试、间隔。**重点：把 `page_size` 从 50 调到 200**，页数直接降一个量级
——如果瓶颈是请求次数触发限流，这是最直接的解法。写好后替换
`auto_fetch_daily._sw_fetch_raw` 即可，上层的分块、去重、单位判定、校验全都不用动。

**B. 缩小区间硬扛**：`--validate-days 20` + `SW_CHUNK_DAYS=15`。只降低概率，不解决问题。

### 先做的诊断（30 秒）

单次最小请求，带 timeout，看是真挂还是慢：

```python
import requests, json
r = requests.get(
    "https://www.swsresearch.com/institute-sw/api/index_analysis/index_analysis_report/",
    params={"page":1,"page_size":50,"index_type":"一级行业","swindexcode":"all",
            "type":"DAY","start_date":"2026-09-15","end_date":"2026-09-16"},
    headers={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                          "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"},
    verify=False, timeout=(10,30))
print(r.status_code, r.json()["data"]["count"])
print(json.dumps(r.json()["data"]["results"][0], ensure_ascii=False, indent=2))
```

超时 → 网络/限流；秒回 → 问题只在 akshare 的翻页循环，选 A 或 B。

---

## 2. 待完成的任务

### 2.1 回填 16 行（主线）

`副本万得全A.xlsx` 中 2026-08-26 ~ 2026-09-16：

| 日期 | 换手率 | 行业占比% | 上涨占比 |
|---|---|---|---|
|2026-08-26|1.48|**空**|0.58|
|2026-08-27|1.48|**空**|0.58|
|2026-08-28|1.48|**空**|0.58|
|2026-08-31|1.48|**空**|0.58|
|2026-09-01|1.48|**空**|0.58|
|2026-09-02|1.48|**空**|0.58|
|2026-09-03|1.48|**空**|0.58|
|2026-09-04|1.48|**空**|0.58|
|2026-09-07|1.48|**空**|0.58|
|2026-09-08|1.48|**空**|0.58|
|2026-09-09|1.48|39.19|0.58|
|2026-09-10|1.48|38.33|0.58|
|2026-09-11|1.48|41.04|0.58|
|2026-09-14|1.48|42.14|0.58|
|2026-09-15|1.48|43.35|0.58|
|2026-09-16|1.48|46.92|0.58|

- **换手率 16 行全是占位常量 `1.48`**，不是真实数据
- 行业占比后 6 行已由日常管线的回补流程填上（`SW_REPAIR_LOOKBACK_DAYS=7` 窗口内），
  **这 6 个值是申万源唯一一次真实产出**，都落在 Wind 历史区间内，是正面信号
- 上涨占比 16 行全是占位 `0.585`，**且无法回填**（见 2.3）

目标：`python backfill_industry_sw.py --apply` 跑通，两列 16 行写满，再 `python update.py`。

注意：放宽 `SW_REPAIR_LOOKBACK_DAYS` **解决不了换手率**——回补只填 `NaN`，而 1.48 是实数。
必须走回填脚本，它对目标区间是无条件覆写。

### 2.2 口径校验（第一次真实对账）

回填脚本写入前强制比对 Excel 里的 Wind 真实历史，任一列不过就整体拒写（退出码 2）：

| 列 | 平均绝对偏差阈值 | 相关系数 | 对照基准 |
|---|---|---|---|
| 行业集中度 | ≤ 3.0 个百分点 | ≥ 0.80 | Wind 25.32~51.64%，均值 36.76% |
| 换手率 | ≤ 0.5 个百分点 | ≥ 0.80 | Wind 1.15~4.21%，均值 1.88% |

**换手率这列是更硬的证据**：它有全部历史可比，行业集中度只有 465 天重叠。

校验窗口只取 `--start` 之前的行（这是个修过的 bug：早先把待回填区间也算进去，
里面的 1.48 占位值把相关系数压到 0.45，导致正常口径被误判为不通过）。

若行业集中度算出 **~20%** 而非 25~52%，说明拿到的是东财细分口径不是申万一级，**回滚**。

### 2.3 上涨占比无解（已确认，别再找了）

akshare 里**没有任何带日期参数的市场宽度接口**。我查过：
`stock_market_activity_legu`（乐咕赚钱效应）、`stock_board_industry_summary_ths`、
`stock_zh_a_spot_em` 全是无参数快照。

现状：`fetch_market_breadth()` 用 `stock_zh_a_spot_em` 数涨跌幅>0 的家数，**只在待追加
日期正好是最新交易日时才写入**，补录多日时较早日期留空。历史的 16 行永久无法回填。

如果要解决，方向是找别的数据源（同花顺/通达信/自建每日快照归档），超出本次范围。

---

## 3. 未经真实验证的假设清单

按风险排序。全部只做过构造数据的离线测试。

1. **`成交额占比` 的单位**：代码按"每日 31 行业合计接近 100 还是 1"自动判定，两者都不像就中止。
   真实响应到底是哪种，未知。
2. **换手率推导**：`Σ(换手率ᵢ×流通市值ᵢ)/Σ(流通市值ᵢ) = Σ成交额ᵢ/Σ流通市值ᵢ`。
   数学成立的前提是申万的 `换手率ᵢ` 确实等于 `成交额ᵢ/流通市值ᵢ`。**没验证过**。
   如果申万用的是自由流通市值或别的分母，结果会系统性偏离——校验门槛应该能拦住。
3. **`index_type=一级行业` 返回的就是 31 个行业**。间接证据：用户实跑时进度条显示 109 页，
   ×50 ≈ 5450 行 ÷ 约 170 交易日 ≈ 32，与 31 吻合。
4. **申万数据的发布延迟**：`SW_REPAIR_LOOKBACK_DAYS=7` 是拍的，不知道实际延迟多久。
5. **CI 环境能否连通申万**。GitHub Actions runner 在境外，访问境内金融站点可能比本地更差。
   这是个真实风险：本地跑通不等于 CI 跑通。

---

## 4. 本次已完成（供理解代码为什么长这样）

按时间顺序，全部已在 `main`：

| 提交 | 内容 |
|---|---|
| `e057747` | 新增 CLAUDE.md |
| `ccbbdee` | 行业集中度改为本地按成交额显式降序，修 top3 名称与金额错位 |
| `bd16700` | 东财 t:1→t:2（**后被 2b06f83 取代**） |
| `02e2f67` | 把 16 行地域口径数据置空 |
| `a54c05d` | 新增 `backfill_industry_sw.py` |
| `2b06f83` | 日常抓取改用申万；修"一次抓取写给多个日期"的老 bug |
| `642d1f3` | 换手率改为申万推导；上涨占比改用快照；清除全部硬编码占位常量 |
| `46ffd67` | 申万抓取分块+重试（其中 socket 超时那层无效） |
| `(最新)` | 删除无效的 socket 兜底，改为 requests 层注入超时 |

**最重要的历史教训**：原代码用东财 `fs=m:90+t:1`，注释写"一级行业31个板块"，实际是
**地域板块（31个省份）**。省份数恰好也是 31，且地域前3占比 40~43% 正好落在 Wind 行业
历史区间 25~52% 内部，所以数字一直"看起来合理"，错了很久没被发现。

**教训本身**：这个项目的数据错误不会以异常的形式出现，只会以"看起来合理的数字"出现。
所以每次换源都必须用重叠期的 Wind 历史做定量对账，不能只看数值范围像不像。

---

## 5. 仓库约定（`CLAUDE.md` 有详版，这里只列最容易踩的）

- **列按位置访问，不按列名**。改动 Excel 列序会静默污染全部下游数字。
  关键位置：1=日期 2=换手率(%) 3=行业占比(小数) 4=上涨占比(小数) 6=总成交额(亿) 7=前三行业合计(亿)
- **`docs/` 是 GitHub Pages 发布目录，是根目录的镜像**。改根目录，让 `update.py` 同步；
  直接改 `docs/` 会被静默覆盖
- **`update.py` 自己会 `git add . && commit && push`**。本地跑它会把工作区一起提交
- **`_latest.xlsx` 陷阱**：Excel 被占用时写入会转存到 `*_latest.xlsx`，而读取端优先读主文件 ——
  一旦触发，下游继续读旧数据且无任何告警
- 分支：开发在 `claude/init-1ndobi`，推 `main` 时同步推 `main:master`
