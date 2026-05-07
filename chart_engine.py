# -*- coding: utf-8 -*-
"""图表生成引擎 — 统一 11 种图表类型"""

from __future__ import annotations

from typing import Any

import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
from matplotlib.figure import Figure

from utils import setup_chinese_font


# ---------- 图表创建入口 ----------
def create_chart(
    df: pd.DataFrame,
    x_col: str,
    y_col: str,
    legend_col: str = "",
    chart_type: str = "stacked_bar",
    theme: str = "whitegrid",
    palette: str = "deep",
    custom_title: str = "",
    show_xlabel: bool = True,
    show_ylabel: bool = True,
    show_legend: bool = True,
    font_title: int = 15,
    font_xlabel: int = 11,
    font_ylabel: int = 11,
    font_legend: int = 10,
    figsize: tuple[float, float] = (14, 5.5),
    dpi: int = 72,
) -> tuple[pd.DataFrame | None, Figure]:
    """统一图表创建入口，返回 (pivot_table | None, figure)"""

    if chart_type == "scatter":
        return _create_scatter(df, x_col, y_col, legend_col, theme, palette,
                               custom_title, show_xlabel, show_ylabel, show_legend,
                               font_title, font_xlabel, font_ylabel, font_legend,
                               figsize, dpi)
    if chart_type in ("pie", "donut"):
        return _create_pie_donut(df, x_col, y_col, legend_col, chart_type, theme, palette,
                                 custom_title, font_title, figsize, dpi)
    if chart_type == "heatmap":
        return _create_heatmap(df, x_col, y_col, legend_col, theme, palette,
                               custom_title, show_xlabel, show_ylabel,
                               font_title, font_xlabel, font_ylabel, figsize, dpi)
    if chart_type == "boxplot":
        return _create_boxplot(df, x_col, y_col, legend_col, theme, palette,
                               custom_title, show_xlabel, show_ylabel, show_legend,
                               font_title, font_xlabel, font_ylabel, font_legend,
                               figsize, dpi)
    # bar / line / area (透视表类图表)
    return _create_pivot_chart(df, x_col, y_col, legend_col, chart_type, theme, palette,
                               custom_title, show_xlabel, show_ylabel, show_legend,
                               font_title, font_xlabel, font_ylabel, font_legend,
                               figsize, dpi)


# ---------- 散点图 ----------
def _create_scatter(
    df: pd.DataFrame, x_col: str, y_col: str, legend_col: str,
    theme: str, palette: str, custom_title: str,
    show_xlabel: bool, show_ylabel: bool, show_legend: bool,
    font_title: int, font_xlabel: int, font_ylabel: int, font_legend: int,
    figsize: tuple[float, float], dpi: int,
) -> tuple[None, Figure]:
    working = df.copy()
    working[y_col] = pd.to_numeric(working[y_col], errors="coerce")
    working[x_col] = pd.to_numeric(working[x_col], errors="coerce")
    working = working.dropna(subset=[x_col, y_col])
    if working.empty:
        raise ValueError("散点图需要 X 和 Y 均为数值列，筛除空值后无可用数据。")

    sns.set_theme(style=theme, palette=palette)
    setup_chinese_font()
    fig, ax = _make_figure(figsize, dpi)

    if legend_col and legend_col in working.columns:
        groups = working[legend_col].unique()
        colors = sns.color_palette(palette, n_colors=len(groups))
        for i, grp in enumerate(groups):
            subset = working[working[legend_col] == grp]
            ax.scatter(subset[x_col], subset[y_col], label=str(grp),
                       color=colors[i % len(colors)], alpha=0.7, s=40)
    else:
        ax.scatter(working[x_col], working[y_col], alpha=0.7, s=40,
                   color=sns.color_palette(palette, 1)[0])

    title = custom_title or (
        f"{x_col} vs {y_col} (按{legend_col}分组)" if legend_col else f"{x_col} vs {y_col}")
    _apply_labels_and_legend(ax, title, x_col if show_xlabel else "",
                             y_col if show_ylabel else "", show_legend and bool(legend_col),
                             font_title, font_xlabel, font_ylabel, font_legend,
                             groups if legend_col else None)
    fig.tight_layout()
    if legend_col:
        fig.subplots_adjust(bottom=0.22)
    return None, fig


# ---------- 饼图 / 环形图 ----------
def _create_pie_donut(
    df: pd.DataFrame, x_col: str, y_col: str, legend_col: str,
    chart_type: str, theme: str, palette: str,
    custom_title: str, font_title: int,
    figsize: tuple[float, float], dpi: int,
) -> tuple[pd.DataFrame, Figure]:
    working = df.copy()
    working[y_col] = pd.to_numeric(working[y_col], errors="coerce")
    working = working.dropna(subset=[y_col])
    if working.empty:
        raise ValueError("Y 轴列没有有效数值。")

    group_col = legend_col if legend_col else x_col
    aggregated = working.groupby(group_col)[y_col].sum().sort_values(ascending=False)
    if aggregated.empty:
        raise ValueError("聚合后无数据。")

    sns.set_theme(style=theme, palette=palette)
    setup_chinese_font()
    fig, ax = _make_figure((9, 7), dpi)
    colors = sns.color_palette(palette, n_colors=len(aggregated))
    wedge_props = dict(width=0.4) if chart_type == "donut" else None
    wedges, texts, autotexts = ax.pie(
        aggregated.values, labels=aggregated.index,
        autopct="%1.1f%%", colors=colors,
        startangle=140, pctdistance=0.78,
        wedgeprops=wedge_props,
    )
    for t in autotexts:
        t.set_fontsize(9)

    title = custom_title or (
        f"{legend_col} 各分类的 {y_col} 占比" if legend_col else f"{x_col} 各分类的 {y_col} 占比")
    ax.set_title(title, fontsize=font_title, pad=20)
    fig.tight_layout()

    pivot = pd.DataFrame({group_col: aggregated.index, y_col: aggregated.values})
    return pivot, fig


# ---------- 热力图 ----------
def _create_heatmap(
    df: pd.DataFrame, x_col: str, y_col: str, legend_col: str,
    theme: str, palette: str, custom_title: str,
    show_xlabel: bool, show_ylabel: bool,
    font_title: int, font_xlabel: int, font_ylabel: int,
    figsize: tuple[float, float], dpi: int,
) -> tuple[pd.DataFrame, Figure]:
    if not legend_col:
        raise ValueError("热力图需要选择图例列以生成二维透视表。")

    working = df.copy()
    working[y_col] = pd.to_numeric(working[y_col], errors="coerce")
    working = working.dropna(subset=[x_col, y_col, legend_col])
    if working.empty:
        raise ValueError("筛除空值后没有可用数据。")

    pivot = pd.pivot_table(
        working, index=x_col, columns=legend_col,
        values=y_col, aggfunc="sum", fill_value=0,
    )
    if pivot.empty:
        raise ValueError("透视表结果为空，无法生成热力图。")

    sns.set_theme(style=theme)
    setup_chinese_font()
    fig, ax = _make_figure((14, 7), dpi)
    cmap = palette if palette in ("viridis", "plasma") else "YlOrRd"
    sns.heatmap(pivot, annot=True, fmt=".0f", cmap=cmap,
                ax=ax, linewidths=0.5, linecolor="white",
                cbar_kws={"shrink": 0.8},
                annot_kws={"fontsize": font_xlabel})

    title = custom_title or f"{x_col}维度下{legend_col}分类的{y_col}热力图"
    ax.set_title(title, fontsize=font_title, pad=20)
    ax.set_xlabel(legend_col if show_xlabel else "", fontsize=font_xlabel)
    ax.set_ylabel(x_col if show_ylabel else "", fontsize=font_ylabel)
    ax.tick_params(labelsize=font_xlabel)
    fig.tight_layout()
    return pivot, fig


# ---------- 箱线图 ----------
def _create_boxplot(
    df: pd.DataFrame, x_col: str, y_col: str, legend_col: str,
    theme: str, palette: str, custom_title: str,
    show_xlabel: bool, show_ylabel: bool, show_legend: bool,
    font_title: int, font_xlabel: int, font_ylabel: int, font_legend: int,
    figsize: tuple[float, float], dpi: int,
) -> tuple[pd.DataFrame, Figure]:
    working = df.copy()
    working[y_col] = pd.to_numeric(working[y_col], errors="coerce")
    working = working.dropna(subset=[y_col])
    if working.empty:
        raise ValueError("Y 轴列没有有效数值。")

    sns.set_theme(style=theme, palette=palette)
    setup_chinese_font()
    fig, ax = _make_figure((14, 6), dpi)

    if legend_col and legend_col in working.columns:
        sns.boxplot(data=working, x=x_col, y=y_col, hue=legend_col,
                    palette=palette, ax=ax)
    else:
        sns.boxplot(data=working, x=x_col, y=y_col, palette=palette, ax=ax)

    title = custom_title or (
        f"{x_col}维度下{legend_col}分类的{y_col}箱线图" if legend_col else f"{x_col}维度的{y_col}箱线图")
    ax.set_title(title, fontsize=font_title, pad=20)
    ax.set_xlabel(x_col if show_xlabel else "", fontsize=font_xlabel)
    ax.set_ylabel(y_col if show_ylabel else "", fontsize=font_ylabel)
    ax.tick_params(axis="x", rotation=45, labelsize=font_xlabel)

    if show_legend and legend_col:
        n_items = working[legend_col].nunique()
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18),
                  ncol=min(n_items, 12), frameon=False, fontsize=font_legend)
        fig.subplots_adjust(bottom=0.22)

    fig.tight_layout()

    if legend_col:
        pivot = pd.pivot_table(
            working, index=x_col, columns=legend_col,
            values=y_col, aggfunc="median", fill_value=0,
        )
    else:
        pivot = pd.DataFrame(working.groupby(x_col)[y_col].median())
    return pivot, fig


# ---------- 透视表类图表 (bar / line / area) ----------
def _create_pivot_chart(
    df: pd.DataFrame, x_col: str, y_col: str, legend_col: str,
    chart_type: str, theme: str, palette: str,
    custom_title: str, show_xlabel: bool, show_ylabel: bool, show_legend: bool,
    font_title: int, font_xlabel: int, font_ylabel: int, font_legend: int,
    figsize: tuple[float, float], dpi: int,
) -> tuple[pd.DataFrame, Figure]:
    working = df.copy()
    working[y_col] = pd.to_numeric(working[y_col], errors="coerce")
    drop_cols = [x_col, y_col, legend_col] if legend_col else [x_col, y_col]
    working = working.dropna(subset=drop_cols)
    if working.empty:
        raise ValueError("筛除空值后没有可用数据。")

    if legend_col:
        pivot = pd.pivot_table(
            working, index=x_col, columns=legend_col,
            values=y_col, aggfunc="sum", fill_value=0,
        )
    else:
        pivot = pd.DataFrame(working.groupby(x_col)[y_col].sum())
        pivot.columns = [y_col]

    if pivot.empty:
        raise ValueError("透视表结果为空。")

    sns.set_theme(style=theme, palette=palette)
    setup_chinese_font()
    is_horizontal = chart_type in ("stacked_barh", "grouped_barh")
    fig_h = 7 if is_horizontal else 5.5
    fig, ax = _make_figure((14, fig_h), dpi)

    n_colors = max(len(pivot.columns), 1)
    colors = sns.color_palette(palette, n_colors=n_colors)

    plot_methods = {
        "stacked_bar": ("bar", True, False),
        "grouped_bar": ("bar", False, False),
        "stacked_barh": ("barh", True, False),
        "grouped_barh": ("barh", False, False),
        "line": ("line", False, True),
        "area": ("area", True, True),
    }
    kind, stacked, is_line_or_area = plot_methods.get(chart_type, ("bar", True, False))

    if chart_type in ("line", "area"):
        pivot.plot(kind=kind, ax=ax, color=colors, marker="o" if kind == "line" else None,
                   linewidth=2 if kind == "line" else None, alpha=0.6 if kind == "area" else None,
                   stacked=stacked)
    else:
        pivot.plot(kind=kind, stacked=stacked, ax=ax, color=colors, width=0.85)

    ax.ticklabel_format(style="plain", axis="x" if is_horizontal else "y")

    title_map = {
        "stacked_bar": f"{x_col}维度下{legend_col}分类的{y_col}堆积柱形图",
        "grouped_bar": f"{x_col}维度下{legend_col}分类的{y_col}分组柱形图",
        "stacked_barh": f"{x_col}维度下{legend_col}分类的{y_col}横向堆积条形图",
        "grouped_barh": f"{x_col}维度下{legend_col}分类的{y_col}横向分组条形图",
        "line": f"{x_col}维度下{legend_col}分类的{y_col}折线图",
        "area": f"{x_col}维度下{legend_col}分类的{y_col}面积图",
    }
    title_map_no_legend = {
        "stacked_bar": f"{x_col}维度的{y_col}柱形图",
        "grouped_bar": f"{x_col}维度的{y_col}柱形图",
        "stacked_barh": f"{x_col}维度的{y_col}条形图",
        "grouped_barh": f"{x_col}维度的{y_col}条形图",
        "line": f"{x_col}维度的{y_col}折线图",
        "area": f"{x_col}维度的{y_col}面积图",
    }
    auto_title = (title_map if legend_col else title_map_no_legend).get(chart_type, "")
    ax.set_title(custom_title or auto_title, fontsize=font_title, pad=20)

    if is_horizontal:
        ax.set_xlabel(y_col if show_xlabel else "", fontsize=font_xlabel)
        ax.set_ylabel(x_col if show_ylabel else "", fontsize=font_ylabel)
    else:
        ax.set_xlabel(x_col if show_xlabel else "", fontsize=font_xlabel)
        ax.set_ylabel(y_col if show_ylabel else "", fontsize=font_ylabel)

    if is_horizontal:
        ax.tick_params(axis="y", rotation=45, labelsize=font_ylabel)
        ax.tick_params(axis="x", labelsize=font_xlabel)
    else:
        ax.tick_params(axis="x", rotation=45, labelsize=font_xlabel)
        ax.tick_params(axis="y", labelsize=font_ylabel)

    if show_legend and legend_col and len(pivot.columns) > 0:
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18),
                  ncol=min(len(pivot.columns), 12), frameon=False, fontsize=font_legend)

    fig.tight_layout()
    if legend_col:
        fig.subplots_adjust(bottom=0.22)
    return pivot, fig


# ---------- 辅助 ----------
def _make_figure(figsize: tuple[float, float], dpi: int = 72) -> tuple[Figure, Any]:
    return plt.subplots(figsize=figsize, dpi=dpi)


def _apply_labels_and_legend(
    ax: Any, title: str, xlabel: str, ylabel: str,
    show_legend: bool, font_title: int, font_xlabel: int,
    font_ylabel: int, font_legend: int,
    groups: Any | None = None,
) -> None:
    ax.set_title(title, fontsize=font_title, pad=20)
    ax.set_xlabel(xlabel, fontsize=font_xlabel)
    ax.set_ylabel(ylabel, fontsize=font_ylabel)
    ax.ticklabel_format(style="plain", axis="y")
    ax.tick_params(labelsize=font_xlabel)
    if show_legend and groups is not None and len(groups) > 0:
        n_cols = max(1, min(len(groups), 12))
        ax.legend(frameon=False, loc="upper center",
                  bbox_to_anchor=(0.5, -0.18),
                  ncol=n_cols, fontsize=font_legend)


