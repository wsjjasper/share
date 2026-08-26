# -*- coding: utf-8 -*-
"""
情绪指标构建脚本
- 4个子指标: 换手率、成交额前3行业占比、上涨个股占比、融资买入额占比
- 滚动252个交易日的分位数排名 (百分位)
- 4个指标等权平均 → 综合情绪指标
- 输出时间范围: 2024-09-24 至今
"""

import pandas as pd
import numpy as np
import os
import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ========================
# 1. 读取数据
# ========================
raw_path = os.path.join(BASE_DIR, '副本万得全A.xlsx')
if not os.path.exists(raw_path):
    raw_path = os.path.join(BASE_DIR, '副本万得全A_latest.xlsx')
    if not os.path.exists(raw_path):
        print(f"错误: 未找到数据文件 {raw_path}")
        sys.exit(1)

df = pd.read_excel(raw_path)

# 重命名列 (按位置)
df.columns = [
    '指数名称', '日期',
    '换手率', '成交额前三行业占比', '上涨个股占比_raw', '融资买入额占比_raw',
    '总成交额', '成交额前三行业合计',
    '上涨家数', '成份个数_最新', '成份个数_历史',
    '上涨个股占比_pct', '融资买入额', '融资买入额占比_pct'
]

# 日期处理
df['日期'] = pd.to_datetime(df['日期'])
df = df.sort_values('日期').reset_index(drop=True)

# ========================
# 2. 提取4个子指标
# ========================
# 换手率 — 已经是百分比
df['ind_换手率'] = df['换手率']

# 成交额前3行业占比 — 原始是0~1的比例, 转为百分比
df['ind_成交额前三行业占比'] = df['成交额前三行业占比'] * 100

# 上涨个股占比 — 原始是0~1的比例, 转为百分比
df['ind_上涨个股占比'] = df['上涨个股占比_raw'] * 100

# 融资买入额占总成交额比例 — 原始是0~1的比例, 转为百分比
# 注: 2010-03-31之前无数据
df['ind_融资买入额占比'] = df['融资买入额占比_raw'] * 100
df.loc[df['ind_融资买入额占比'] == 0, 'ind_融资买入额占比'] = np.nan

# ========================
# 3. 滚动252日分位数排名 (向量化)
# ========================
indicators = ['ind_换手率', 'ind_成交额前三行业占比', 'ind_上涨个股占比', 'ind_融资买入额占比']
window = 252


def rolling_percentile_rank(series, window):
    """
    对每个时点, 计算当前值在过去window个交易日内的百分位排名 (0~100)
    使用 pandas rolling + apply
    """
    def percentile_rank(arr):
        """当前值(arr末尾)在窗口中的百分位"""
        valid = arr[~np.isnan(arr)]
        if len(valid) < window * 0.5:
            return np.nan
        current = arr[-1]
        if np.isnan(current):
            return np.nan
        # 百分位 = (小于当前值的个数 + 0.5 * 等于当前值的个数) / 总个数 * 100
        below = np.sum(valid < current)
        equal = np.sum(valid == current)
        rank = (below + 0.5 * equal) / len(valid) * 100
        return rank

    return series.rolling(window=window, min_periods=int(window * 0.5)).apply(
        percentile_rank, raw=True
    )


print("正在计算滚动252日分位数...")
for ind in indicators:
    col_name = f'pct_{ind}'
    print(f"  处理: {ind}")
    df[col_name] = rolling_percentile_rank(df[ind], window)

# ========================
# 4. 等权平均 → 综合情绪指标
# ========================
pct_cols = [f'pct_{ind}' for ind in indicators]
df['综合情绪指标'] = df[pct_cols].mean(axis=1)

# ========================
# 5. 筛选输出范围: 2024-09-24 至今
# ========================
start_date = pd.Timestamp('2024-09-24')
output = df[df['日期'] >= start_date].copy()

# 输出结果
output_cols = ['日期', 'ind_换手率', 'ind_成交额前三行业占比', 'ind_上涨个股占比', 'ind_融资买入额占比',
               'pct_ind_换手率', 'pct_ind_成交额前三行业占比', 'pct_ind_上涨个股占比', 'pct_ind_融资买入额占比',
               '综合情绪指标']
output = output[output_cols]

# 列名简化
output.columns = ['日期',
                   '换手率(%)', '成交额前三行业占比(%)', '上涨个股占比(%)', '融资买入额占比(%)',
                   '换手率_分位', '成交额前三行业占比_分位', '上涨个股占比_分位', '融资买入额占比_分位',
                   '综合情绪指标']

# 保存Excel
output_path = os.path.join(BASE_DIR, '情绪指标_结果.xlsx')
try:
    output.to_excel(output_path, index=False, float_format='%.2f')
    print(f"\n结果已保存至: {output_path}")
except PermissionError:
    print(f"\n[警告] {output_path} 正在被 Excel 打开占用，正在尝试写入备份文件...")
    backup_path = os.path.join(BASE_DIR, '情绪指标_结果_latest.xlsx')
    output.to_excel(backup_path, index=False, float_format='%.2f')
    print(f"已保存至备份文件: {backup_path}")

print(f"数据行数: {len(output)}")
print(f"日期范围: {output['日期'].min()} ~ {output['日期'].max()}")

# 输出最近几行数据
print("\n最近10个交易日的情绪指标:")
recent = output.tail(10).copy()
recent['日期'] = recent['日期'].dt.strftime('%Y-%m-%d')
for _, row in recent.iterrows():
    print(f"  {row['日期']}  换手率={row['换手率(%)']:.2f}% (分位{row['换手率_分位']:.1f})  "
          f"行业集中={row['成交额前三行业占比(%)']:.1f}% (分位{row['成交额前三行业占比_分位']:.1f})  "
          f"上涨占比={row['上涨个股占比(%)']:.1f}% (分位{row['上涨个股占比_分位']:.1f})  "
          f"融资占比={row['融资买入额占比(%)']:.1f}% (分位{row['融资买入额占比_分位']:.1f})  "
          f"→ 综合={row['综合情绪指标']:.1f}")

# ========================
# 6. 可视化
# ========================
fig, axes = plt.subplots(3, 1, figsize=(16, 14), sharex=True)

# 字体设置 (兼容 Windows / Linux 云端运行)
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'WenQuanYi Micro Hei', 'Noto Sans CJK SC', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

dates = output['日期']

# 上图: 4个子指标的分位数
ax1 = axes[0]
ax1.plot(dates, output['换手率_分位'], label='换手率分位', linewidth=1.2, alpha=0.8)
ax1.plot(dates, output['成交额前三行业占比_分位'], label='成交额前三行业占比分位', linewidth=1.2, alpha=0.8)
ax1.plot(dates, output['上涨个股占比_分位'], label='上涨个股占比分位', linewidth=1.2, alpha=0.8)
ax1.plot(dates, output['融资买入额占比_分位'], label='融资买入额占比分位', linewidth=1.2, alpha=0.8)
ax1.axhline(y=80, color='red', linestyle='--', alpha=0.5, label='过热阈值(80)')
ax1.axhline(y=20, color='green', linestyle='--', alpha=0.5, label='过冷阈值(20)')
ax1.set_ylabel('分位数 (%)')
ax1.set_title('4个子指标 - 滚动252日分位数')
ax1.legend(loc='upper left', fontsize=8)
ax1.set_ylim(0, 100)
ax1.grid(True, alpha=0.3)

# 中图: 综合情绪指标
ax2 = axes[1]
ax2.fill_between(dates, output['综合情绪指标'], alpha=0.3, color='steelblue')
ax2.plot(dates, output['综合情绪指标'], color='steelblue', linewidth=2, label='综合情绪指标')
ax2.axhline(y=80, color='red', linestyle='--', alpha=0.5, label='过热(80)')
ax2.axhline(y=50, color='gray', linestyle='--', alpha=0.3, label='中位(50)')
ax2.axhline(y=20, color='green', linestyle='--', alpha=0.5, label='过冷(20)')
ax2.set_ylabel('综合情绪指标 (%)')
ax2.set_title('综合情绪指标 (4个子指标等权平均)')
ax2.legend(loc='upper left', fontsize=8)
ax2.set_ylim(0, 100)
ax2.grid(True, alpha=0.3)

# 下图: 4个原始指标
ax3 = axes[2]
ax3_twin = ax3.twinx()
ax3.plot(dates, output['换手率(%)'], label='换手率(%)', color='tab:blue', linewidth=1, alpha=0.7)
ax3.plot(dates, output['融资买入额占比(%)'], label='融资买入额占比(%)', color='tab:orange', linewidth=1, alpha=0.7)
ax3_twin.plot(dates, output['上涨个股占比(%)'], label='上涨个股占比(%)', color='tab:green', linewidth=1, alpha=0.7)
ax3_twin.plot(dates, output['成交额前三行业占比(%)'], label='成交额前三行业占比(%)', color='tab:red', linewidth=1, alpha=0.7)
ax3.set_ylabel('换手率 / 融资占比 (%)')
ax3_twin.set_ylabel('上涨占比 / 行业集中度 (%)')
ax3.set_xlabel('日期')
ax3.set_title('4个原始指标绝对值')
# Combine legends
lines1, labels1 = ax3.get_legend_handles_labels()
lines2, labels2 = ax3_twin.get_legend_handles_labels()
ax3.legend(lines1 + lines2, labels1 + labels2, loc='upper left', fontsize=8)
ax3.grid(True, alpha=0.3)

ax3.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
ax3.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
plt.xticks(rotation=45)

plt.tight_layout()
chart_path = os.path.join(BASE_DIR, '情绪指标_图表.png')
plt.savefig(chart_path, dpi=150, bbox_inches='tight')
print(f"\n图表已保存至: {chart_path}")
plt.close()
