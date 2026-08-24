# -*- coding: utf-8 -*-
"""
Setup repository structure for docs folder
"""

import shutil
import os

base_dir = r'D:\vscode_workspace\情绪指标'
docs_dir = os.path.join(base_dir, 'docs')
os.makedirs(docs_dir, exist_ok=True)

files_to_copy_to_docs = [
    'index.html',
    'sentiment_chart.png',
    'sentiment_indicator.py',
    'generate_html.py',
    '情绪指标_结果.xlsx',
    '情绪指标_图表.png',
    '副本万得全A.xlsx',
    'col_info.txt',
    'data_records.json'
]

for fname in files_to_copy_to_docs:
    src = os.path.join(base_dir, fname)
    dst = os.path.join(docs_dir, fname)
    if os.path.exists(src):
        shutil.copy2(src, dst)
        print(f"Copied {fname} to docs/")

# Create docs/README.md
docs_readme = """# A股综合情绪指标研报 (A-Share Sentiment Indicator)

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
"""

with open(os.path.join(docs_dir, 'README.md'), 'w', encoding='utf-8') as f:
    f.write(docs_readme)
print("Created docs/README.md")

# Create root README.md
root_readme = """# A-Share Market Sentiment Indicator (A股综合情绪指标)

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
"""

with open(os.path.join(base_dir, 'README.md'), 'w', encoding='utf-8') as f:
    f.write(root_readme)
print("Created root README.md")

# Create .gitignore
gitignore_content = """__pycache__/
*.pyc
.system_generated/
"""

with open(os.path.join(base_dir, '.gitignore'), 'w', encoding='utf-8') as f:
    f.write(gitignore_content)
print("Created .gitignore")
