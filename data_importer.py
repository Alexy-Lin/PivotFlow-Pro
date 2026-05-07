# -*- coding: utf-8 -*-
"""数据导入模块 — 多 Excel 合并、列对齐、增量追加"""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def read_excel_files(paths: list[str]) -> list[pd.DataFrame]:
    """读取多个 Excel 文件，返回 DataFrame 列表"""
    dfs: list[pd.DataFrame] = []
    for p in paths:
        source = Path(p)
        if not source.exists():
            continue
        try:
            dfs.append(pd.read_excel(source, engine="openpyxl"))
        except Exception:
            continue
    return dfs


def merge_dataframes(dfs: list[pd.DataFrame]) -> pd.DataFrame:
    """按列名合并多个 DataFrame（纵向追加，列不同时自动补齐，缺失值留空）"""
    if not dfs:
        raise ValueError("没有可合并的数据表。")
    return pd.concat(dfs, axis=0, join="outer", ignore_index=True, sort=False)


def import_files(paths: list[str]) -> tuple[pd.DataFrame, list[str]]:
    """读取并合并多个 Excel 文件，返回 (合并后的 DataFrame, 源文件名列表)"""
    dfs = read_excel_files(paths)
    if not dfs:
        raise ValueError("未能读取任何有效数据。")
    merged = merge_dataframes(dfs)
    filenames = [Path(p).name for p in paths if Path(p).exists()]
    return merged, filenames


def append_to_existing(existing: pd.DataFrame, paths: list[str]) -> tuple[pd.DataFrame, int]:
    """向已有 DataFrame 追加新文件数据，返回 (合并后的 DataFrame, 新增行数)"""
    dfs = read_excel_files(paths)
    if not dfs:
        raise ValueError("未能读取任何有效数据。")
    new_data = merge_dataframes(dfs)
    combined = merge_dataframes([existing, new_data])
    added_rows = len(combined) - len(existing)
    return combined, added_rows
