# -*- coding: utf-8 -*-
"""共享常量：主题、配色、图表类型、UI 颜色"""

from typing import Literal

ThemeName = Literal["darkgrid", "whitegrid", "dark", "white", "ticks"]

THEMES: dict[str, str] = {
    "darkgrid": "darkgrid",
    "whitegrid": "whitegrid",
    "dark": "dark",
    "white": "white",
    "ticks": "ticks",
}

PALETTES: dict[str, str] = {
    "deep": "deep",
    "muted": "muted",
    "bright": "bright",
    "pastel": "pastel",
    "colorblind": "colorblind",
    "Set2": "Set2",
    "tab10": "tab10",
    "viridis": "viridis",
    "plasma": "plasma",
}

CHART_TYPES: dict[str, str] = {
    "堆积柱形图": "stacked_bar",
    "分组柱形图": "grouped_bar",
    "横向堆积条形图": "stacked_barh",
    "横向分组条形图": "grouped_barh",
    "折线图": "line",
    "面积图": "area",
    "饼图": "pie",
    "环形图": "donut",
    "散点图": "scatter",
    "热力图": "heatmap",
    "箱线图": "boxplot",
}

# GUI 配色方案
COLORS = {
    # 品牌色
    "primary": "#4F46E5",
    "primary_hover": "#4338CA",
    "primary_light": "#EEF2FF",
    "primary_subtle": "#F5F3FF",
    # 背景
    "bg": "#F1F5F9",
    "surface": "#FFFFFF",
    "surface_secondary": "#F8FAFC",
    # 边框
    "border": "#E2E8F0",
    "border_light": "#F1F5F9",
    # 文字
    "text": "#0F172A",
    "text_secondary": "#64748B",
    "text_muted": "#94A3B8",
    # 语义色
    "success": "#10B981",
    "success_light": "#ECFDF5",
    "warning": "#F59E0B",
    "warning_light": "#FFFBEB",
    "error": "#EF4444",
    "error_light": "#FEF2F2",
    # 侧边栏
    "sidebar_bg": "#0F172A",
    "sidebar_surface": "#1E293B",
    "sidebar_text": "#F1F5F9",
    "sidebar_text_secondary": "#94A3B8",
    "sidebar_accent": "#6366F1",
    "sidebar_border": "#334155",
    # 数据表格
    "row_alt": "#F8FAFC",
    "row_hover": "#EEF2FF",
    # 图表区
    "chart_bg": "#FAFAFA",
}

