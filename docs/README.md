# A股综合情绪指标研报 (A-Share Sentiment Indicator)

本项目基于万得全A (881001.WI) 日频交易数据，构建由四大子指标等权合成的综合情绪指标体系。

## 核心指标体系
1. **换手率**: 市场交易活跃度
2. **成交额前三行业占比**: 行业结构集中度与抱团程度
3. **上涨个股占比**: 市场广度与赚钱效应
4. **融资买入额占比**: 杠杆资金进攻情绪

## 分位数与合成方法
- **分位数转换**: 滚动 252 交易日历史分位数排名（百分位 $0 \sim 100\%$）
- **等权平均**: 四大分位数等权合成综合情绪指标（$0 \sim 100\%$）
- **区间定义**: 
  - $> 80$: 过热区间
  - $20 \sim 80$: 中性区间
  - $< 20$: 过冷区间

## 访问在线报告
- 打开 `docs/index.html` 即可查看交互式研报图表与数据看板。
- GitHub Pages 地址: `https://wsjjasper.github.io/share/` (或 `https://wsjjasper.github.io/share/docs/`)

## 文件清单
- `index.html`: 交互式网页研报（集成 ECharts、KaTeX、TailwindCSS）
- `sentiment_chart.png`: 高清 Matplotlib 研报图表
- `情绪指标_结果.xlsx`: 2024-09-24 至今的分位数与综合指标清洗数据集
- `sentiment_indicator.py`: 核心计算与制图脚本
- `generate_html.py`: 自动化生成 index.html 脚本
- `副本万得全A.xlsx`: 万得全A历史原始数据集
