# A-Share Market Sentiment Indicator (A股综合情绪指标)

[![GitHub Pages](https://img.shields.io/badge/GitHub%20Pages-Live%20Report-blue?logo=github)](https://wsjjasper.github.io/share/)

基于万得全A (881001.WI) 日频数据构建的A股综合情绪指标体系（2024年9月24日至今，滚动252日分位数）。

👉 **[在线研报看板 (Live Dashboard)](https://wsjjasper.github.io/share/)** 或查看 [`docs/index.html`](docs/index.html)。

---

## 目录结构
```
├── docs/
│   ├── index.html              # 交互式网页研报 (ECharts + KaTeX + Tailwind)
│   ├── sentiment_chart.png     # 高清研报走势图
│   ├── 情绪指标_结果.xlsx       # 2024-09-24 至今完整清洗数据集
│   ├── sentiment_indicator.py  # 核心量化计算与制图 Python 脚本
│   ├── generate_html.py        # 研报网页自动生成脚本
│   ├── 副本万得全A.xlsx        # 原始万得全A行情与微观结构数据
│   └── README.md
├── index.html                  # 根目录重定向与镜像入口
├── README.md
└── .gitignore
```

## 核心方法
1. **四大微观维度**: 换手率、行业集中度、上涨占比、融资买入占比
2. **滚动历史分位数**: 计算过去 252 交易日（1年）百分位排名，消除量纲与长周期体量漂移
3. **等权合成**: 四大子指标分位数等权平均得到综合情绪指标（$0 \sim 100$）
