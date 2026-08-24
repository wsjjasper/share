# -*- coding: utf-8 -*-
"""
Generate index.html from sentiment analysis results with exquisitely styled human-readable formulas, interactive zoom & linked ECharts
"""

import pandas as pd
import json
import os
import shutil

# Read excel results (with fallback)
excel_path = r'D:\vscode_workspace\情绪指标\情绪指标_结果.xlsx'
if not os.path.exists(excel_path):
    excel_path = r'D:\vscode_workspace\情绪指标\情绪指标_结果_latest.xlsx'

df = pd.read_excel(excel_path)
cols = [
    'date', 'turnover', 'top3_ind', 'rise_pct', 'margin_pct',
    'pct_turnover', 'pct_top3_ind', 'pct_rise_pct', 'pct_margin_pct',
    'composite_sentiment'
]
df.columns = cols
df['date'] = pd.to_datetime(df['date']).dt.strftime('%Y-%m-%d')

data_records = []
for idx, row in df.iterrows():
    data_records.append({
        'date': row['date'],
        'turnover': None if pd.isna(row['turnover']) else round(float(row['turnover']), 2),
        'top3_ind': None if pd.isna(row['top3_ind']) else round(float(row['top3_ind']), 2),
        'rise_pct': None if pd.isna(row['rise_pct']) else round(float(row['rise_pct']), 2),
        'margin_pct': None if pd.isna(row['margin_pct']) else round(float(row['margin_pct']), 2),
        'pct_turnover': None if pd.isna(row['pct_turnover']) else round(float(row['pct_turnover']), 2),
        'pct_top3_ind': None if pd.isna(row['pct_top3_ind']) else round(float(row['pct_top3_ind']), 2),
        'pct_rise_pct': None if pd.isna(row['pct_rise_pct']) else round(float(row['pct_rise_pct']), 2),
        'pct_margin_pct': None if pd.isna(row['pct_margin_pct']) else round(float(row['pct_margin_pct']), 2),
        'composite_sentiment': None if pd.isna(row['composite_sentiment']) else round(float(row['composite_sentiment']), 2)
    })

data_json = json.dumps(data_records, ensure_ascii=False)

# Latest values
latest = data_records[-1]
prev = data_records[-2]

latest_date = latest['date']
comp_val = latest['composite_sentiment']
turnover_val = latest['turnover']
pct_turnover_val = latest['pct_turnover']
top3_val = latest['top3_ind']
pct_top3_val = latest['pct_top3_ind']
rise_val = latest['rise_pct']
pct_rise_val = latest['pct_rise_pct']
prev_margin_pct_val = prev['pct_margin_pct']

if comp_val > 80:
    comp_badge_class = "bg-red-100 text-red-700"
    comp_badge_text = "🔥 过热"
elif comp_val < 20:
    comp_badge_class = "bg-green-100 text-green-700"
    comp_badge_text = "❄️ 过冷"
elif comp_val < 50:
    comp_badge_class = "bg-blue-100 text-blue-700"
    comp_badge_text = "⚖️ 中性偏冷"
else:
    comp_badge_class = "bg-amber-100 text-amber-700"
    comp_badge_text = "⚖️ 中性偏暖"

html_content = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>A股综合情绪指标研报 | A-Share Market Sentiment Indicator</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Apache ECharts -->
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js"></script>
    <!-- KaTeX for math rendering -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css">
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js"></script>
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/contrib/auto-render.min.js"
        onload="renderMathInElement(document.body, {
            delimiters: [
                {left: '$$', right: '$$', display: true},
                {left: '$', right: '$', display: false}
            ]
        });"></script>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Microsoft YaHei', sans-serif;
        }
        code, pre {
            font-family: 'JetBrains Mono', monospace;
        }
        .glass-card {
            background: rgba(255, 255, 255, 0.95);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(226, 232, 240, 0.8);
        }
        .zoom-btn {
            transition: all 0.2s ease;
            cursor: pointer;
        }
        .zoom-btn:hover {
            transform: translateY(-1px);
        }
        .zoom-btn:active {
            transform: translateY(0px);
        }
        .formula-card {
            background: linear-gradient(145deg, #0f172a, #1e293b);
            border: 1px solid rgba(255, 255, 255, 0.1);
        }
        .fraction-box {
            display: inline-flex;
            flex-direction: column;
            vertical-align: middle;
            text-align: center;
            padding: 0 4px;
        }
        .fraction-numerator {
            border-bottom: 2px solid rgba(129, 140, 248, 0.8);
            padding-bottom: 4px;
            font-weight: 500;
        }
        .fraction-denominator {
            padding-top: 4px;
            font-weight: 500;
            color: #cbd5e1;
        }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 antialiased min-h-screen">

    <!-- Top Navigation Header -->
    <header class="sticky top-0 z-50 bg-white/90 backdrop-blur-md border-b border-slate-200 shadow-sm">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
            <div class="flex items-center space-x-3">
                <div class="w-9 h-9 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-500 flex items-center justify-center text-white font-bold shadow-md shadow-blue-500/20">
                    <svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="20" x2="18" y2="10"></line>
                        <line x1="12" y1="20" x2="12" y2="4"></line>
                        <line x1="6" y1="20" x2="6" y2="14"></line>
                    </svg>
                </div>
                <div>
                    <h1 class="text-lg font-bold text-slate-900 leading-tight">A股综合情绪指标研报</h1>
                    <p class="text-xs text-slate-500">万得全A (881001.WI) 滚动252日分位数体系</p>
                </div>
            </div>
            <div class="flex items-center space-x-4">
                <span class="inline-flex items-center px-3 py-1 rounded-full text-xs font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                    <span class="w-2 h-2 mr-1.5 rounded-full bg-blue-500 animate-pulse"></span>
                    最新交易日: LATEST_DATE
                </span>
                <a href="https://github.com/wsjjasper/share.git" target="_blank" class="text-slate-500 hover:text-slate-800 transition-colors p-2 rounded-lg hover:bg-slate-100" title="GitHub Repository">
                    <svg class="w-5 h-5 fill-current" viewBox="0 0 24 24">
                        <path fill-rule="evenodd" clip-rule="evenodd" d="M12 2C6.477 2 2 6.484 2 12.017c0 4.425 2.865 8.18 6.839 9.504.5.092.682-.217.682-.483 0-.237-.008-.868-.013-1.703-2.782.605-3.369-1.343-3.369-1.343-.454-1.158-1.11-1.466-1.11-1.466-.908-.62.069-.608.069-.608 1.003.07 1.53 1.032 1.53 1.032.892 1.53 2.341 1.088 2.91.832.092-.647.35-1.088.636-1.338-2.22-.253-4.555-1.113-4.555-4.951 0-1.093.39-1.988 1.029-2.688-.103-.253-.446-1.272.098-2.65 0 0 .84-.27 2.75 1.026A9.564 9.564 0 0112 6.844c.85.004 1.705.115 2.504.337 1.909-1.296 2.747-1.027 2.747-1.027.546 1.379.202 2.398.1 2.651.64.7 1.028 1.595 1.028 2.688 0 3.848-2.339 4.695-4.566 4.943.359.309.678.92.678 1.855 0 1.338-.012 2.419-.012 2.747 0 .268.18.58.688.482A10.019 10.019 0 0022 12.017C22 6.484 17.522 2 12 2z"/>
                    </svg>
                </a>
            </div>
        </div>
    </header>

    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">

        <!-- Real-Time Metric Cards Dashboard -->
        <section>
            <div class="flex items-center justify-between mb-4">
                <h2 class="text-xl font-bold text-slate-900 flex items-center space-x-2">
                    <span>最新情绪读数看板</span>
                    <span class="text-xs font-normal text-slate-500 bg-slate-100 px-2.5 py-0.5 rounded-full">T=LATEST_DATE</span>
                </h2>
                <div class="text-xs text-slate-500 flex items-center space-x-3">
                    <span class="flex items-center"><span class="w-2.5 h-2.5 rounded-full bg-red-500 inline-block mr-1"></span>过热 (>80)</span>
                    <span class="flex items-center"><span class="w-2.5 h-2.5 rounded-full bg-slate-400 inline-block mr-1"></span>中性 (20-80)</span>
                    <span class="flex items-center"><span class="w-2.5 h-2.5 rounded-full bg-green-500 inline-block mr-1"></span>过冷 (<20)</span>
                </div>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
                <!-- Composite Sentiment Card -->
                <div class="glass-card rounded-2xl p-5 shadow-sm border-2 border-indigo-200 bg-gradient-to-br from-indigo-50/50 to-white relative overflow-hidden">
                    <div class="absolute -right-4 -bottom-4 w-24 h-24 bg-indigo-500/10 rounded-full blur-xl pointer-events-none"></div>
                    <div class="text-xs font-semibold text-indigo-600 uppercase tracking-wider mb-1">综合情绪指标</div>
                    <div class="flex items-baseline space-x-2">
                        <span class="text-3xl font-extrabold text-indigo-900">COMP_VAL</span>
                        <span class="text-xs font-medium text-slate-500">/ 100</span>
                    </div>
                    <div class="mt-2 flex items-center justify-between">
                        <span class="text-xs px-2 py-0.5 rounded-md font-medium COMP_BADGE_CLASS">
                            COMP_BADGE_TEXT
                        </span>
                        <span class="text-xs text-slate-500">等权合成</span>
                    </div>
                    <div class="mt-3 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div class="bg-indigo-600 h-full rounded-full" style="width: COMP_VAL%"></div>
                    </div>
                </div>

                <!-- Sub 1: Turnover -->
                <div class="glass-card rounded-2xl p-5 shadow-sm">
                    <div class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">换手率分位</div>
                    <div class="flex items-baseline space-x-2">
                        <span class="text-2xl font-bold text-slate-900">PCT_TURNOVER%</span>
                        <span class="text-xs text-slate-500">(TURNOVER_VAL%)</span>
                    </div>
                    <div class="mt-2 flex items-center justify-between text-xs">
                        <span class="font-medium text-blue-600">低位极冷</span>
                        <span class="text-slate-400">活跃度低</span>
                    </div>
                    <div class="mt-3 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div class="bg-blue-500 h-full rounded-full" style="width: PCT_TURNOVER%"></div>
                    </div>
                </div>

                <!-- Sub 2: Top 3 Industries -->
                <div class="glass-card rounded-2xl p-5 shadow-sm">
                    <div class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">行业集中度分位</div>
                    <div class="flex items-baseline space-x-2">
                        <span class="text-2xl font-bold text-slate-900">PCT_TOP3%</span>
                        <span class="text-xs text-slate-500">(TOP3_VAL%)</span>
                    </div>
                    <div class="mt-2 flex items-center justify-between text-xs">
                        <span class="font-medium text-amber-600">抱团较高</span>
                        <span class="text-slate-400">前3行业成交</span>
                    </div>
                    <div class="mt-3 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div class="bg-amber-500 h-full rounded-full" style="width: PCT_TOP3%"></div>
                    </div>
                </div>

                <!-- Sub 3: Rising Stocks -->
                <div class="glass-card rounded-2xl p-5 shadow-sm">
                    <div class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">上涨个股分位</div>
                    <div class="flex items-baseline space-x-2">
                        <span class="text-2xl font-bold text-slate-900">PCT_RISE%</span>
                        <span class="text-xs text-slate-500">(RISE_VAL%)</span>
                    </div>
                    <div class="mt-2 flex items-center justify-between text-xs">
                        <span class="font-medium text-slate-600">均衡中性</span>
                        <span class="text-slate-400">赚钱效应</span>
                    </div>
                    <div class="mt-3 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div class="bg-emerald-500 h-full rounded-full" style="width: PCT_RISE%"></div>
                    </div>
                </div>

                <!-- Sub 4: Margin Buying -->
                <div class="glass-card rounded-2xl p-5 shadow-sm">
                    <div class="text-xs font-semibold text-slate-500 uppercase tracking-wider mb-1">融资买入占比分位</div>
                    <div class="flex items-baseline space-x-2">
                        <span class="text-2xl font-bold text-slate-900">PREV_MARGIN%</span>
                        <span class="text-xs text-slate-500">(前日值)</span>
                    </div>
                    <div class="mt-2 flex items-center justify-between text-xs">
                        <span class="font-medium text-slate-500">杠杆偏低</span>
                        <span class="text-slate-400">杠杆情绪</span>
                    </div>
                    <div class="mt-3 w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                        <div class="bg-rose-500 h-full rounded-full" style="width: PREV_MARGIN%"></div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Interactive Dynamic Charts (ECharts) -->
        <section class="glass-card rounded-2xl p-6 shadow-sm border border-slate-200">
            <div class="flex flex-col md:flex-row md:items-center justify-between pb-4 border-b border-slate-100 gap-2">
                <div>
                    <h2 class="text-lg font-bold text-slate-900">情绪指标交互式全景走势 (2024.09.24 - 至今)</h2>
                    <p class="text-xs text-slate-500 mt-0.5">支持滚轮缩放、底部滑动条拖拽、十字光标联动、图例筛选与快捷区间切换</p>
                </div>
                <!-- Range Zoom Buttons -->
                <div class="flex items-center space-x-2 text-xs flex-wrap gap-y-1">
                    <button id="btn-zoom-1m" onclick="zoomByMonths(1, this)" class="zoom-btn px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium">近1个月</button>
                    <button id="btn-zoom-3m" onclick="zoomByMonths(3, this)" class="zoom-btn px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium">近3个月</button>
                    <button id="btn-zoom-6m" onclick="zoomByMonths(6, this)" class="zoom-btn px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium">近6个月</button>
                    <button id="btn-zoom-1y" onclick="zoomByMonths(12, this)" class="zoom-btn px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium">近1年</button>
                    <button id="btn-zoom-all" onclick="zoomAll(this)" class="zoom-btn px-3 py-1.5 rounded-lg bg-indigo-600 text-white font-medium shadow-sm">全部区间</button>
                    <button id="btn-zoom-reset" onclick="resetZoom()" class="zoom-btn px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-medium border border-slate-200">重置缩放</button>
                </div>
            </div>

            <!-- ECharts Chart Containers -->
            <div class="space-y-6 mt-6">
                <!-- Chart 1: 4 Sub-indicators -->
                <div>
                    <h3 class="text-sm font-semibold text-slate-700 mb-2 flex items-center">
                        <span class="w-2.5 h-2.5 rounded-full bg-indigo-500 inline-block mr-2"></span>
                        四大子指标滚动252日分位数走势 (0~100%)
                    </h3>
                    <div id="chart-sub-indicators" class="w-full h-72"></div>
                </div>

                <!-- Chart 2: Composite Sentiment -->
                <div>
                    <h3 class="text-sm font-semibold text-slate-700 mb-2 flex items-center">
                        <span class="w-2.5 h-2.5 rounded-full bg-blue-600 inline-block mr-2"></span>
                        综合情绪指标历史走势 (等权合成)
                    </h3>
                    <div id="chart-composite" class="w-full h-64"></div>
                </div>

                <!-- Chart 3: Raw Values -->
                <div>
                    <h3 class="text-sm font-semibold text-slate-700 mb-2 flex items-center">
                        <span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block mr-2"></span>
                        原始指标绝对值走势 (双Y轴: 换手率/融资占比 vs 上涨家数占比/行业集中度)
                    </h3>
                    <div id="chart-raw" class="w-full h-72"></div>
                </div>
            </div>
        </section>

        <!-- Methodology & Human-Readable Framework -->
        <section class="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <!-- Left: 4 Core Indicators Table -->
            <div class="glass-card rounded-2xl p-6 shadow-sm border border-slate-200 flex flex-col justify-between">
                <div>
                    <h2 class="text-lg font-bold text-slate-900 mb-3 flex items-center">
                        <svg class="w-5 h-5 mr-2 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                        </svg>
                        四大微观情绪子指标
                    </h2>
                    <p class="text-sm text-slate-600 leading-relaxed mb-4">
                        基于 <strong>万得全A (881001.WI)</strong> 核心日频交易数据，从“交易活跃度、资金抱团集中度、赚钱效应广度、杠杆做多情绪”四个独立维度综合衡量市场水温：
                    </p>

                    <div class="overflow-x-auto rounded-xl border border-slate-200">
                        <table class="w-full text-left text-xs text-slate-600">
                            <thead class="bg-slate-50 text-slate-700 uppercase font-semibold border-b border-slate-200">
                                <tr>
                                    <th class="px-3 py-2.5">#</th>
                                    <th class="px-3 py-2.5">子指标</th>
                                    <th class="px-3 py-2.5">通俗定义与计算口径</th>
                                    <th class="px-3 py-2.5">市场微观含义</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-100 bg-white">
                                <tr class="hover:bg-slate-50">
                                    <td class="px-3 py-2.5 font-bold text-slate-900">1</td>
                                    <td class="px-3 py-2.5 font-semibold text-blue-600">换手率</td>
                                    <td class="px-3 py-2.5">当日全市场成交额 / 自由流通市值</td>
                                    <td class="px-3 py-2.5">衡量交投活跃度与筹码换手意愿</td>
                                </tr>
                                <tr class="hover:bg-slate-50">
                                    <td class="px-3 py-2.5 font-bold text-slate-900">2</td>
                                    <td class="px-3 py-2.5 font-semibold text-amber-600">前3行业占比</td>
                                    <td class="px-3 py-2.5">成交额最大的前3个行业合计 / 全市场成交额</td>
                                    <td class="px-3 py-2.5">衡量主线资金抱团集中度与分化程度</td>
                                </tr>
                                <tr class="hover:bg-slate-50">
                                    <td class="px-3 py-2.5 font-bold text-slate-900">3</td>
                                    <td class="px-3 py-2.5 font-semibold text-emerald-600">上涨个股占比</td>
                                    <td class="px-3 py-2.5">当日上涨股票数量 / 全市场交易股票总数</td>
                                    <td class="px-3 py-2.5">衡量市场广度与散户普遍赚钱效应</td>
                                </tr>
                                <tr class="hover:bg-slate-50">
                                    <td class="px-3 py-2.5 font-bold text-slate-900">4</td>
                                    <td class="px-3 py-2.5 font-semibold text-rose-600">融资买入占比</td>
                                    <td class="px-3 py-2.5">两市融资买入总金额 / 全市场总成交额</td>
                                    <td class="px-3 py-2.5">衡量高风险偏好杠杆资金的主动进攻情绪</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>

                <div class="mt-4 p-3 bg-blue-50/70 border border-blue-200/80 rounded-xl text-xs text-blue-800">
                    <strong>💡 为什么需要分位数转化？</strong>
                    直接看换手率或成交额等绝对数值，会受到市场总市值膨胀、制度变革等长期趋势影响。通过转换为“滚动252日相对历史分位数”，可以直接比较当前情绪与过去一年相比是冷是热。
                </div>
            </div>

            <!-- Right: Human Readable Math & Formulas (HTML/CSS Formatted) -->
            <div class="glass-card rounded-2xl p-6 shadow-sm border border-slate-200 flex flex-col justify-between">
                <div>
                    <h2 class="text-lg font-bold text-slate-900 mb-3 flex items-center">
                        <svg class="w-5 h-5 mr-2 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 20l4-16m2 16l4-16M6 9h14M4 15h14"></path>
                        </svg>
                        算法逻辑与通俗公式
                    </h2>

                    <div class="space-y-4 text-sm text-slate-600">
                        <!-- Step 1 Formula Card (Pure CSS Visual Equation) -->
                        <div class="formula-card text-slate-100 p-4 rounded-xl shadow-md">
                            <div class="text-indigo-400 font-semibold text-xs tracking-wider mb-2 flex items-center">
                                <span class="w-2 h-2 rounded-full bg-indigo-400 inline-block mr-2"></span>
                                第一步：单项指标 ➔ 滚动 252 日历史分位数 (0 ~ 100%)
                            </div>
                            
                            <!-- Visual Equation Box -->
                            <div class="bg-slate-950/60 border border-slate-800 rounded-lg p-3 my-2 flex items-center justify-center flex-wrap gap-2 text-xs md:text-sm text-slate-100 font-sans">
                                <span class="font-bold text-indigo-300">历史分位数 (%) =</span>
                                <div class="fraction-box">
                                    <div class="fraction-numerator text-indigo-200">
                                        过去252日中数值 &lt; 今日的天数 + 0.5 × 等于今日的天数
                                    </div>
                                    <div class="fraction-denominator">
                                        过去252个交易日有效总天数
                                    </div>
                                </div>
                                <span class="font-bold text-indigo-300">× 100%</span>
                            </div>

                            <p class="text-xs text-slate-400 mt-2">
                                📌 <strong>通俗含义</strong>：若今日换手率分位数为 <strong>85%</strong>，代表今日活跃度比过去一年 <strong>85% 的交易日都要火热</strong>；0% 表示过去一年最低，100% 表示过去一年最高。
                            </p>
                        </div>

                        <!-- Step 2 Formula Card (Pure CSS Visual Equation) -->
                        <div class="formula-card text-slate-100 p-4 rounded-xl shadow-md">
                            <div class="text-emerald-400 font-semibold text-xs tracking-wider mb-2 flex items-center">
                                <span class="w-2 h-2 rounded-full bg-emerald-400 inline-block mr-2"></span>
                                第二步：四大分位数 ➔ 等权平均合成综合指标 (0 ~ 100)
                            </div>

                            <!-- Visual Equation Box -->
                            <div class="bg-slate-950/60 border border-slate-800 rounded-lg p-3 my-2 flex items-center justify-center flex-wrap gap-2 text-xs md:text-sm text-slate-100 font-sans">
                                <span class="font-bold text-emerald-300">综合情绪指标 =</span>
                                <div class="fraction-box">
                                    <div class="fraction-numerator text-emerald-200">
                                        换手率分位 + 行业集中度分位 + 上涨个股分位 + 融资买入分位
                                    </div>
                                    <div class="fraction-denominator">
                                        4 （有效子指标个数）
                                    </div>
                                </div>
                            </div>

                            <p class="text-xs text-slate-400 mt-2">
                                📌 <strong>通俗含义</strong>：四大维度各占 <strong>25% 权重</strong> 简单平均，消除单项指标的短期噪点。（若遇融资数据尚未发布，自动对剩余 3 个有效指标取均值）。
                            </p>
                        </div>

                        <!-- Regime Cards -->
                        <div class="grid grid-cols-3 gap-2 pt-1 text-center text-xs">
                            <div class="p-2.5 rounded-xl bg-red-50 border border-red-200">
                                <div class="font-bold text-red-700">> 80 分位：过热区间</div>
                                <div class="text-red-600 mt-0.5">全面亢奋 / 警惕短线见顶回撤</div>
                            </div>
                            <div class="p-2.5 rounded-xl bg-slate-100 border border-slate-200">
                                <div class="font-bold text-slate-700">20 ~ 80 分位：中性区间</div>
                                <div class="text-slate-600 mt-0.5">情绪常态 / 结构分化轮动</div>
                            </div>
                            <div class="p-2.5 rounded-xl bg-emerald-50 border border-emerald-200">
                                <div class="font-bold text-emerald-700">< 20 分位：过冷区间</div>
                                <div class="text-emerald-600 mt-0.5">极度悲观 / 往往孕育底部反弹</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- Static Chart & Graphic Artifact -->
        <section class="glass-card rounded-2xl p-6 shadow-sm border border-slate-200">
            <h2 class="text-lg font-bold text-slate-900 mb-2 flex items-center">
                <svg class="w-5 h-5 mr-2 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"></path>
                </svg>
                静态研报图表存档 (High-Resolution Output)
            </h2>
            <p class="text-xs text-slate-500 mb-4">由 Python 自动化任务生成的高清图表输出，展示 2024 年 9 月 24 日至今历史全貌：</p>
            <div class="rounded-xl overflow-hidden border border-slate-200 shadow-inner bg-slate-900 flex items-center justify-center p-2">
                <img src="sentiment_chart.png" alt="A股综合情绪指标走势图" class="w-full h-auto rounded-lg shadow">
            </div>
        </section>

        <!-- Recent Data Table -->
        <section class="glass-card rounded-2xl p-6 shadow-sm border border-slate-200">
            <div class="flex items-center justify-between mb-4">
                <div>
                    <h2 class="text-lg font-bold text-slate-900">近15个交易日详细数据明细</h2>
                    <p class="text-xs text-slate-500">展示各单项指标与百分位排名的逐日变化</p>
                </div>
                <a href="情绪指标_结果.xlsx" download class="inline-flex items-center px-3 py-1.5 rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 font-medium text-xs border border-indigo-200 transition-colors">
                    <svg class="w-4 h-4 mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path>
                    </svg>
                    下载完整 Excel 结果 (464天)
                </a>
            </div>

            <div class="overflow-x-auto rounded-xl border border-slate-200">
                <table class="w-full text-left text-xs text-slate-600">
                    <thead class="bg-slate-50 text-slate-700 uppercase font-semibold border-b border-slate-200">
                        <tr>
                            <th class="px-3 py-2.5">交易日期</th>
                            <th class="px-3 py-2.5">换手率 (%)</th>
                            <th class="px-3 py-2.5">换手率分位</th>
                            <th class="px-3 py-2.5">前3行业占比 (%)</th>
                            <th class="px-3 py-2.5">行业集中分位</th>
                            <th class="px-3 py-2.5">上涨个股占比 (%)</th>
                            <th class="px-3 py-2.5">上涨分位</th>
                            <th class="px-3 py-2.5">融资买入占比 (%)</th>
                            <th class="px-3 py-2.5">融资分位</th>
                            <th class="px-3 py-2.5 font-bold text-indigo-700 bg-indigo-50/50">综合情绪指标</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100 bg-white font-mono" id="recent-table-body">
                        <!-- Filled by JavaScript -->
                    </tbody>
                </table>
            </div>
        </section>

        <!-- Repository & Artifacts Assets -->
        <section class="glass-card rounded-2xl p-6 shadow-sm border border-slate-200">
            <h2 class="text-lg font-bold text-slate-900 mb-3 flex items-center">
                <svg class="w-5 h-5 mr-2 text-indigo-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 4H6a2 2 0 00-2 2v12a2 2 0 002 2h12a2 2 0 002-2V6a2 2 0 00-2-2h-2m-4-1v8m0 0l3-3m-3 3L9 8m-5 5h2.586a1 1 0 01.707.293l2.414 2.414a1 1 0 00.707.293h3.172a1 1 0 00.707-.293l2.414-2.414a1 1 0 01.707-.293H20"></path>
                </svg>
                项目归档与代码文件
            </h2>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                <a href="sentiment_indicator.py" download class="p-4 rounded-xl border border-slate-200 hover:border-indigo-400 bg-slate-50/50 hover:bg-indigo-50/30 transition-all flex items-start space-x-3">
                    <div class="p-2 rounded-lg bg-blue-100 text-blue-700">
                        <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 20l4-16m4 4l4 4-4 4M6 16l-4-4 4-4"></path>
                        </svg>
                    </div>
                    <div>
                        <div class="font-bold text-slate-900">sentiment_indicator.py</div>
                        <p class="text-slate-500 mt-1">完整计算与制图脚本，支持未来数据追加与一键回测</p>
                    </div>
                </a>

                <a href="情绪指标_结果.xlsx" download class="p-4 rounded-xl border border-slate-200 hover:border-indigo-400 bg-slate-50/50 hover:bg-indigo-50/30 transition-all flex items-start space-x-3">
                    <div class="p-2 rounded-lg bg-emerald-100 text-emerald-700">
                        <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
                        </svg>
                    </div>
                    <div>
                        <div class="font-bold text-slate-900">情绪指标_结果.xlsx</div>
                        <p class="text-slate-500 mt-1">2024-09-24 至今 464 个交易日分位数与综合指标清洗数据集</p>
                    </div>
                </a>

                <a href="https://github.com/wsjjasper/share.git" target="_blank" class="p-4 rounded-xl border border-slate-200 hover:border-indigo-400 bg-slate-50/50 hover:bg-indigo-50/30 transition-all flex items-start space-x-3">
                    <div class="p-2 rounded-lg bg-purple-100 text-purple-700">
                        <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14"></path>
                        </svg>
                    </div>
                    <div>
                        <div class="font-bold text-slate-900">GitHub Repository</div>
                        <p class="text-slate-500 mt-1">https://github.com/wsjjasper/share.git</p>
                    </div>
                </a>
            </div>
        </section>
    </main>

    <footer class="bg-white border-t border-slate-200 py-6 mt-12 text-center text-xs text-slate-500">
        <div class="max-w-7xl mx-auto px-4">
            <p>A股综合情绪指标研究与量化分析系统 · 自动归档至 GitHub docs</p>
            <p class="mt-1">数据来源: Wind 万得全A (881001.WI) · 滚动 252 交易日历史分位数模型</p>
        </div>
    </footer>

    <!-- Chart & Table Initialization Script -->
    <script>
        const rawData = DATA_JSON_PLACEHOLDER;

        // Populate Table with last 15 days
        const tbody = document.getElementById('recent-table-body');
        const recentRows = rawData.slice(-15).reverse();
        recentRows.forEach(r => {
            const tr = document.createElement('tr');
            tr.className = "hover:bg-slate-50 transition-colors";
            
            const getBadge = (val) => {
                if (val === null || val === undefined) return '<span class="text-slate-400">NaN</span>';
                if (val >= 80) return `<span class="px-2 py-0.5 rounded text-xs font-semibold bg-red-100 text-red-700">${val}</span>`;
                if (val <= 20) return `<span class="px-2 py-0.5 rounded text-xs font-semibold bg-green-100 text-green-700">${val}</span>`;
                return `<span class="px-2 py-0.5 rounded text-xs text-slate-700 bg-slate-100">${val}</span>`;
            };

            const getCompositeBadge = (val) => {
                if (val === null || val === undefined) return '<span class="text-slate-400">NaN</span>';
                if (val >= 80) return `<span class="px-2.5 py-1 rounded-md text-xs font-bold bg-red-500 text-white shadow-sm">${val}</span>`;
                if (val <= 20) return `<span class="px-2.5 py-1 rounded-md text-xs font-bold bg-emerald-500 text-white shadow-sm">${val}</span>`;
                return `<span class="px-2.5 py-1 rounded-md text-xs font-bold bg-indigo-100 text-indigo-800">${val}</span>`;
            };

            tr.innerHTML = `
                <td class="px-3 py-2.5 font-sans font-medium text-slate-900">${r.date}</td>
                <td class="px-3 py-2.5">${r.turnover ?? '-'}</td>
                <td class="px-3 py-2.5">${getBadge(r.pct_turnover)}</td>
                <td class="px-3 py-2.5">${r.top3_ind ?? '-'}</td>
                <td class="px-3 py-2.5">${getBadge(r.pct_top3_ind)}</td>
                <td class="px-3 py-2.5">${r.rise_pct ?? '-'}</td>
                <td class="px-3 py-2.5">${getBadge(r.pct_rise_pct)}</td>
                <td class="px-3 py-2.5">${r.margin_pct ?? '-'}</td>
                <td class="px-3 py-2.5">${getBadge(r.pct_margin_pct)}</td>
                <td class="px-3 py-2.5 bg-indigo-50/40">${getCompositeBadge(r.composite_sentiment)}</td>
            `;
            tbody.appendChild(tr);
        });

        // Initialize ECharts instances
        const dates = rawData.map(d => d.date);
        const chartSub = echarts.init(document.getElementById('chart-sub-indicators'));
        const chartComp = echarts.init(document.getElementById('chart-composite'));
        const chartRaw = echarts.init(document.getElementById('chart-raw'));

        const markLines = {
            silent: true,
            symbol: 'none',
            data: [
                { yAxis: 80, lineStyle: { color: '#ef4444', type: 'dashed', width: 1.5 }, label: { show: true, position: 'end', formatter: '过热 (80)', color: '#ef4444' } },
                { yAxis: 50, lineStyle: { color: '#94a3b8', type: 'dotted', width: 1 }, label: { show: true, position: 'end', formatter: '中位 (50)', color: '#94a3b8' } },
                { yAxis: 20, lineStyle: { color: '#22c55e', type: 'dashed', width: 1.5 }, label: { show: true, position: 'end', formatter: '过冷 (20)', color: '#22c55e' } }
            ]
        };

        // Shared dataZoom configs for charts
        const commonDataZoom = [
            {
                type: 'inside',
                xAxisIndex: [0],
                start: 0,
                end: 100
            }
        ];

        const sliderDataZoom = [
            {
                type: 'inside',
                xAxisIndex: [0],
                start: 0,
                end: 100
            },
            {
                type: 'slider',
                xAxisIndex: [0],
                bottom: 8,
                height: 22,
                start: 0,
                end: 100,
                borderColor: '#e2e8f0',
                backgroundColor: '#f8fafc',
                fillerColor: 'rgba(79, 70, 229, 0.18)',
                handleStyle: { color: '#4f46e5', borderColor: '#4338ca' },
                moveHandleStyle: { color: '#6366f1' },
                dataBackground: {
                    lineStyle: { color: '#cbd5e1' },
                    areaStyle: { color: '#e2e8f0' }
                }
            }
        ];

        // Option 1: Sub-indicators
        const optSub = {
            tooltip: { trigger: 'axis', axisPointer: { type: 'cross' } },
            legend: { data: ['换手率分位', '前3行业占比分位', '上涨个股占比分位', '融资买入占比分位'], top: 0 },
            grid: { left: '4%', right: '5%', top: '15%', bottom: '15%' },
            xAxis: { type: 'category', data: dates, boundaryGap: false },
            yAxis: { type: 'value', min: 0, max: 100, name: '分位数 (%)' },
            dataZoom: commonDataZoom,
            series: [
                { name: '换手率分位', type: 'line', data: rawData.map(d => d.pct_turnover), smooth: true, itemStyle: { color: '#3b82f6' }, lineStyle: { width: 1.5 } },
                { name: '前3行业占比分位', type: 'line', data: rawData.map(d => d.pct_top3_ind), smooth: true, itemStyle: { color: '#f59e0b' }, lineStyle: { width: 1.5 } },
                { name: '上涨个股占比分位', type: 'line', data: rawData.map(d => d.pct_rise_pct), smooth: true, itemStyle: { color: '#10b981' }, lineStyle: { width: 1.5 } },
                { name: '融资买入占比分位', type: 'line', data: rawData.map(d => d.pct_margin_pct), smooth: true, itemStyle: { color: '#ef4444' }, lineStyle: { width: 1.5 } },
                { type: 'line', markLine: markLines }
            ]
        };

        // Option 2: Composite
        const optComp = {
            tooltip: { trigger: 'axis', axisPointer: { type: 'cross' } },
            legend: { data: ['综合情绪指标'], top: 0 },
            grid: { left: '4%', right: '5%', top: '15%', bottom: '15%' },
            xAxis: { type: 'category', data: dates, boundaryGap: false },
            yAxis: { type: 'value', min: 0, max: 100, name: '综合情绪 (%)' },
            dataZoom: commonDataZoom,
            series: [
                {
                    name: '综合情绪指标',
                    type: 'line',
                    data: rawData.map(d => d.composite_sentiment),
                    smooth: true,
                    itemStyle: { color: '#4f46e5' },
                    lineStyle: { width: 2.5 },
                    areaStyle: {
                        color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                            { offset: 0, color: 'rgba(79, 70, 229, 0.45)' },
                            { offset: 0.8, color: 'rgba(79, 70, 229, 0.05)' },
                            { offset: 1, color: 'rgba(79, 70, 229, 0.0)' }
                        ])
                    },
                    markLine: markLines
                }
            ]
        };

        // Option 3: Raw Values (Dual Y-Axis)
        const optRaw = {
            tooltip: { trigger: 'axis', axisPointer: { type: 'cross' } },
            legend: { data: ['换手率(%)', '融资买入额占比(%)', '上涨个股占比(%)', '前3行业占比(%)'], top: 0 },
            grid: { left: '4%', right: '5%', top: '15%', bottom: '25%' },
            xAxis: { type: 'category', data: dates, boundaryGap: false },
            yAxis: [
                { type: 'value', name: '换手/融资 (%)', position: 'left' },
                { type: 'value', name: '上涨/集中度 (%)', position: 'right' }
            ],
            dataZoom: sliderDataZoom,
            series: [
                { name: '换手率(%)', type: 'line', yAxisIndex: 0, data: rawData.map(d => d.turnover), smooth: true, itemStyle: { color: '#3b82f6' } },
                { name: '融资买入额占比(%)', type: 'line', yAxisIndex: 0, data: rawData.map(d => d.margin_pct), smooth: true, itemStyle: { color: '#f97316' } },
                { name: '上涨个股占比(%)', type: 'line', yAxisIndex: 1, data: rawData.map(d => d.rise_pct), smooth: true, itemStyle: { color: '#10b981' } },
                { name: '前3行业占比(%)', type: 'line', yAxisIndex: 1, data: rawData.map(d => d.top3_ind), smooth: true, itemStyle: { color: '#ef4444' } }
            ]
        };

        chartSub.setOption(optSub);
        chartComp.setOption(optComp);
        chartRaw.setOption(optRaw);

        // Connect charts for unified cursor/tooltip interaction
        echarts.connect([chartSub, chartComp, chartRaw]);

        // Synchronize dataZoom across all 3 charts when scrolling/dragging
        let isSyncing = false;
        function syncZoom(sourceChart) {
            sourceChart.on('dataZoom', function (params) {
                if (isSyncing) return;
                isSyncing = true;
                let start, end;
                if (params.batch && params.batch[0]) {
                    start = params.batch[0].start;
                    end = params.batch[0].end;
                } else {
                    start = params.start;
                    end = params.end;
                }
                if (start !== undefined && end !== undefined) {
                    [chartSub, chartComp, chartRaw].forEach(target => {
                        if (target !== sourceChart) {
                            target.dispatchAction({
                                type: 'dataZoom',
                                start: start,
                                end: end
                            });
                        }
                    });
                }
                setTimeout(() => { isSyncing = false; }, 50);
            });
        }
        syncZoom(chartSub);
        syncZoom(chartComp);
        syncZoom(chartRaw);

        window.addEventListener('resize', () => {
            chartSub.resize();
            chartComp.resize();
            chartRaw.resize();
        });

        // Zoom helper functions with active button state
        function setZoomPercent(startPct, btnEl) {
            const zoomPayload = { start: startPct, end: 100 };
            
            // Dispatch dataZoom action to all charts
            [chartSub, chartComp, chartRaw].forEach(chart => {
                chart.dispatchAction({
                    type: 'dataZoom',
                    start: startPct,
                    end: 100
                });
            });

            // Update button styles
            document.querySelectorAll('.zoom-btn').forEach(b => {
                if (b.id !== 'btn-zoom-reset') {
                    b.classList.remove('bg-indigo-600', 'text-white', 'shadow-sm');
                    b.classList.add('bg-slate-100', 'text-slate-700');
                }
            });
            if (btnEl && btnEl.id !== 'btn-zoom-reset') {
                btnEl.classList.remove('bg-slate-100', 'text-slate-700');
                btnEl.classList.add('bg-indigo-600', 'text-white', 'shadow-sm');
            }
        }

        function zoomByMonths(months, btnEl) {
            const latestDateObj = new Date(dates[dates.length - 1]);
            const targetDateObj = new Date(latestDateObj);
            targetDateObj.setMonth(targetDateObj.getMonth() - months);
            const targetStr = targetDateObj.toISOString().slice(0, 10);
            
            let idx = dates.findIndex(d => d >= targetStr);
            if (idx === -1) idx = 0;
            const startPct = Math.max(0, Math.min(99, (idx / (dates.length - 1)) * 100));
            setZoomPercent(startPct, btnEl);
        }

        function zoomAll(btnEl) {
            setZoomPercent(0, btnEl);
        }

        function resetZoom() {
            setZoomPercent(0, document.getElementById('btn-zoom-all'));
        }
    </script>
</body>
</html>
"""

# Replace placeholders
html_content = html_content.replace('LATEST_DATE', str(latest_date))
html_content = html_content.replace('COMP_VAL', str(comp_val))
html_content = html_content.replace('COMP_BADGE_CLASS', comp_badge_class)
html_content = html_content.replace('COMP_BADGE_TEXT', comp_badge_text)
html_content = html_content.replace('PCT_TURNOVER', str(pct_turnover_val))
html_content = html_content.replace('TURNOVER_VAL', str(turnover_val))
html_content = html_content.replace('PCT_TOP3', str(pct_top3_val))
html_content = html_content.replace('TOP3_VAL', str(top3_val))
html_content = html_content.replace('PCT_RISE', str(pct_rise_val))
html_content = html_content.replace('RISE_VAL', str(rise_val))
html_content = html_content.replace('PREV_MARGIN', str(prev_margin_pct_val))
html_content = html_content.replace('DATA_JSON_PLACEHOLDER', data_json)

# Write to root index.html and docs/index.html
output_html_root = r'D:\vscode_workspace\情绪指标\index.html'
output_html_docs = r'D:\vscode_workspace\情绪指标\docs\index.html'

with open(output_html_root, 'w', encoding='utf-8') as f:
    f.write(html_content)

with open(output_html_docs, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"Generated index.html successfully with visual fraction cards!")
