# -*- coding: utf-8 -*-
"""
情绪指标的共用计算核心 —— A股与美股两条管线都从这里取分位数算法,
避免两处各自维护同一套数学而慢慢漂移。
"""

import numpy as np

WINDOW = 252          # 滚动窗口(交易日)
MIN_VALID_RATIO = 0.5  # 窗口内有效值不足此比例则返回 NaN


def rolling_percentile_rank(series, window=WINDOW):
    """
    对每个时点, 计算当前值在过去 window 个交易日内的百分位排名 (0~100)。

    采用中位排名法: (小于当前值的个数 + 0.5 × 等于当前值的个数) / 有效个数 × 100。
    窗口内有效值少于一半, 或当前值本身缺失, 返回 NaN。
    """
    def percentile_rank(arr):
        valid = arr[~np.isnan(arr)]
        if len(valid) < window * MIN_VALID_RATIO:
            return np.nan
        current = arr[-1]
        if np.isnan(current):
            return np.nan
        below = np.sum(valid < current)
        equal = np.sum(valid == current)
        return (below + 0.5 * equal) / len(valid) * 100

    return series.rolling(window=window, min_periods=int(window * MIN_VALID_RATIO)).apply(
        percentile_rank, raw=True
    )
