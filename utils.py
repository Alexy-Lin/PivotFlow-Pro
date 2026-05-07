# -*- coding: utf-8 -*-
"""共享工具：依赖检查、字体配置、日期聚合"""

from __future__ import annotations

import importlib
import subprocess
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib import font_manager

REQUIRED_PACKAGES = {
    "pandas": "pandas",
    "matplotlib": "matplotlib",
    "seaborn": "seaborn",
    "openpyxl": "openpyxl",
}

_CHINESE_FONT_NAME: str | None = None


# ---------- 依赖检查 ----------
def ensure_dependencies() -> None:
    missing = []
    for module_name, package_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(package_name)
    if missing:
        print("检测到缺少依赖，正在自动安装:", ", ".join(missing))
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])


# ---------- 字体配置 ----------
def _find_chinese_font_file() -> str | None:
    font_dir = Path("C:/Windows/Fonts")
    if not font_dir.exists():
        return None
    candidates = [
        "msyh.ttc", "msyhbd.ttc", "simhei.ttf", "simsun.ttc",
        "simkai.ttf", "simfang.ttf", "msjh.ttc",
    ]
    for candidate in candidates:
        font_path = font_dir / candidate
        if font_path.exists():
            return str(font_path)
    return None


def _scan_matplotlib_fonts() -> str | None:
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in ("Microsoft YaHei", "SimHei", "Microsoft JhengHei", "SimSun", "KaiTi", "FangSong"):
        if name in installed:
            return name
    return None


def _apply_font_rcparams(font_name: str) -> None:
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = [font_name, "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False


def setup_chinese_font() -> str:
    """中文字体配置（优先加载字体文件）"""
    global _CHINESE_FONT_NAME
    if _CHINESE_FONT_NAME:
        _apply_font_rcparams(_CHINESE_FONT_NAME)
        return _CHINESE_FONT_NAME

    font_path = _find_chinese_font_file()
    if font_path:
        try:
            font_manager.fontManager.addfont(font_path)
            font_manager.fontManager._load_fontmanager(try_read_cache=False)
        except Exception:
            pass
        font_name = _scan_matplotlib_fonts()
        if font_name:
            _CHINESE_FONT_NAME = font_name
            _apply_font_rcparams(font_name)
            return font_name

    font_name = _scan_matplotlib_fonts()
    if font_name:
        _CHINESE_FONT_NAME = font_name
        _apply_font_rcparams(font_name)
        return font_name

    _CHINESE_FONT_NAME = "sans-serif"
    _apply_font_rcparams("sans-serif")
    return "sans-serif"


# ---------- 日期列检测与富化 ----------
_FREQ_LABELS: list[tuple[str, str]] = [("周", "W"), ("月", "M"), ("季度", "Q")]


def detect_date_column(series: pd.Series) -> bool:
    """检测列是否包含日期数据"""
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    # 数值列不探测（Excel 序列号会被误转），仅对字符串/对象列尝试转换
    if pd.api.types.is_numeric_dtype(series):
        return False
    try:
        converted = pd.to_datetime(series, errors="coerce")
        return converted.notna().sum() > len(series) * 0.5
    except Exception:
        return False


def enrich_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    """检测日期列并自动添加周/月/季度/年聚合列，返回富化后的 DataFrame"""
    enriched = df.copy()
    for col in df.columns:
        if not detect_date_column(df[col]):
            continue
        dt = pd.to_datetime(df[col], errors="coerce")
        for suffix, freq in _FREQ_LABELS:
            new_col = f"{col}({suffix})"
            enriched[new_col] = dt.dt.to_period(freq).astype(str)
    return enriched
