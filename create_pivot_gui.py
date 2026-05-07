#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel 数据透视表生成器 — GUI
侧边栏 + 内容区布局 | 拖拽加载 | 11 种图表
"""

from __future__ import annotations

import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from utils import ensure_dependencies
ensure_dependencies()

import tkinter as tk
import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from config import THEMES, PALETTES, CHART_TYPES, COLORS
from chart_engine import create_chart
from utils import setup_chinese_font, enrich_date_columns
from data_importer import import_files, append_to_existing


# ---------- 主应用 ----------
class PivotChartApp:
    """Excel 数据透视表生成器"""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel 数据透视表生成器")
        self.root.geometry("1340x840")
        self.root.minsize(1100, 680)
        self.root.configure(bg=COLORS["bg"])

        self.dataframe: pd.DataFrame | None = None
        self.columns: list[str] = []
        self.current_figure: Figure | None = None
        self.current_pivot: pd.DataFrame | None = None
        self.temp_file: Path | None = None
        self.imported_files: list[str] = []

        setup_chinese_font()
        self._setup_styles()
        self._build_header()
        self._build_body()
        self._build_statusbar()
        self._bind_shortcuts()
        self._setup_drop_target()

    # ========== 样式 ==========
    def _setup_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(".", font=("Microsoft YaHei", 10), background=COLORS["bg"])

        # 标签页
        style.configure("App.TNotebook", background=COLORS["bg"], borderwidth=0)
        style.configure("App.TNotebook.Tab",
                        font=("Microsoft YaHei", 10), padding=[20, 8],
                        background=COLORS["surface"], foreground=COLORS["text_secondary"],
                        borderwidth=0)
        style.map("App.TNotebook.Tab",
                  background=[("selected", COLORS["surface"])],
                  foreground=[("selected", COLORS["primary"])],
                  expand=[("selected", [1, 1, 0, 0])])

        # Treeview
        style.configure("App.Treeview",
                        font=("Microsoft YaHei", 10), background=COLORS["surface"],
                        fieldbackground=COLORS["surface"], rowheight=32, borderwidth=0)
        style.configure("App.Treeview.Heading",
                        font=("Microsoft YaHei", 9, "bold"),
                        background=COLORS["surface_secondary"],
                        foreground=COLORS["text"], relief="flat", padding=(10, 7))
        style.map("App.Treeview.Heading", background=[("active", COLORS["primary_light"])])
        style.map("App.Treeview",
                  background=[("selected", COLORS["primary"])],
                  foreground=[("selected", "#FFFFFF")])

        # Combobox
        style.configure("App.TCombobox",
                        fieldbackground=COLORS["surface"], arrowsize=14, borderwidth=1,
                        relief="solid", bordercolor=COLORS["border"])

        # Progressbar
        style.configure("Horizontal.TProgressbar", thickness=3,
                        troughcolor=COLORS["border_light"], background=COLORS["primary"])

    # ========== 顶部导航栏 ==========
    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=COLORS["surface"], height=48)
        header.pack(fill=tk.X, side=tk.TOP, padx=0, pady=0)
        header.pack_propagate(False)

        tk.Frame(header, bg=COLORS["border"], height=1).pack(fill=tk.X, side=tk.BOTTOM)

        # 左侧品牌
        brand = tk.Frame(header, bg=COLORS["surface"])
        brand.pack(side=tk.LEFT, padx=(18, 0))

        logo = tk.Canvas(brand, width=32, height=32, bg=COLORS["surface"], highlightthickness=0)
        logo.create_rectangle(4, 4, 28, 28, fill=COLORS["primary"], outline="", stipple="")
        logo.create_rectangle(10, 10, 28, 28,
                              fill="", outline=COLORS["primary"], width=2)
        logo.create_rectangle(6, 14, 17, 22,
                              fill="", outline="#FFFFFF", width=1)
        logo.create_rectangle(10, 8, 17, 18,
                              fill="", outline="#FFFFFF", width=1)
        logo.pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(brand, text="PivotFlow", font=("Microsoft YaHei", 14, "bold"),
                 bg=COLORS["surface"], fg=COLORS["text"]).pack(side=tk.LEFT)
        tk.Label(brand, text="Pro",
                 font=("Microsoft YaHei", 9, "bold"),
                 bg=COLORS["primary"], fg="#FFFFFF",
                 padx=6, pady=1).pack(side=tk.LEFT, padx=(6, 0), pady=(2, 0))

        # 右侧文件信息
        self.header_file_var = tk.StringVar(value="未打开文件")
        file_info = tk.Frame(header, bg=COLORS["surface"])
        file_info.pack(side=tk.RIGHT, padx=18)

        self.header_dot = tk.Canvas(file_info, width=6, height=6,
                                    bg=COLORS["surface"], highlightthickness=0)
        self.header_dot.pack(side=tk.LEFT, padx=(0, 6))
        self._set_header_dot("#D1D5DB")

        tk.Label(file_info, textvariable=self.header_file_var,
                 font=("Microsoft YaHei", 9), bg=COLORS["surface"],
                 fg=COLORS["text_secondary"]).pack(side=tk.LEFT)

    def _set_header_dot(self, color: str) -> None:
        self.header_dot.delete("all")
        self.header_dot.create_oval(0, 0, 6, 6, fill=color, outline="")

    # ========== 主体 ==========
    def _build_body(self) -> None:
        body = tk.Frame(self.root, bg=COLORS["bg"])
        body.pack(fill=tk.BOTH, expand=True)

        self._build_sidebar(body)
        self._build_main(body)

    # ========== 侧边栏 ==========
    def _build_sidebar(self, parent: tk.Frame) -> None:
        sidebar = tk.Frame(parent, bg=COLORS["sidebar_bg"], width=285)
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        sidebar.pack_propagate(False)

        canvas = tk.Canvas(sidebar, bg=COLORS["sidebar_bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(sidebar, orient=tk.VERTICAL, command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=COLORS["sidebar_bg"])

        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw", width=272)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 鼠标滚轮
        def _scroll(e):
            w = e.widget
            while w is not None:
                if w is sidebar:
                    canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
                    return
                w = w.master
        for t in (canvas, scroll_frame, sidebar):
            t.bind("<MouseWheel>", _scroll)
            t.bind("<Button-4>", _scroll)
            t.bind("<Button-5>", _scroll)

        pad = 12
        inner_pad = 12

        # ---- 侧边栏标题 ----
        title_area = tk.Frame(scroll_frame, bg=COLORS["sidebar_bg"])
        title_area.pack(fill=tk.X, padx=pad, pady=(16, 0))
        tk.Label(title_area, text="图表配置",
                 font=("Microsoft YaHei", 15, "bold"),
                 bg=COLORS["sidebar_bg"], fg=COLORS["sidebar_text"]).pack(anchor=tk.W)
        tk.Label(title_area, text="选择数据列与图表样式",
                 font=("Microsoft YaHei", 9),
                 bg=COLORS["sidebar_bg"], fg=COLORS["sidebar_text_secondary"]).pack(
            anchor=tk.W, pady=(2, 0))

        # ---- 数据源卡片 ----
        self._sidebar_section_header(scroll_frame, "数据源", pad=pad)
        card = tk.Frame(scroll_frame, bg=COLORS["sidebar_surface"], highlightthickness=1,
                        highlightbackground=COLORS["sidebar_border"])
        card.pack(fill=tk.X, padx=pad, pady=(4, 0))

        # 数据库状态
        self.db_info_var = tk.StringVar(value="未导入数据")
        tk.Label(card, textvariable=self.db_info_var, font=("Microsoft YaHei", 9),
                 bg=COLORS["sidebar_surface"], fg=COLORS["sidebar_text_secondary"]).pack(
            anchor=tk.W, padx=inner_pad, pady=(8, 2))

        # 按钮行：导入 + 保存 + 关闭
        btn_row = tk.Frame(card, bg=COLORS["sidebar_surface"])
        btn_row.pack(fill=tk.X, padx=inner_pad, pady=(4, inner_pad))

        self.close_db_btn = tk.Button(btn_row, text="关闭", font=("Microsoft YaHei", 9),
                                      bg=COLORS["sidebar_surface"], fg=COLORS["error"],
                                      relief="flat", cursor="hand2", bd=0,
                                      activebackground=COLORS["error_light"],
                                      activeforeground=COLORS["error"],
                                      padx=10, pady=5, command=self._close_database,
                                      state=tk.DISABLED)
        self.close_db_btn.pack(side=tk.RIGHT)

        self.save_source_btn = tk.Button(btn_row, text="保存", font=("Microsoft YaHei", 9),
                                         bg=COLORS["sidebar_surface"], fg=COLORS["text_muted"],
                                         relief="flat", cursor="hand2", bd=0,
                                         activebackground=COLORS["sidebar_surface"],
                                         activeforeground=COLORS["sidebar_text"],
                                         padx=10, pady=5, command=self._save_source,
                                         state=tk.DISABLED)
        self.save_source_btn.pack(side=tk.RIGHT, padx=(0, 4))

        self._btn_primary(btn_row, "导入数据", self._import_files,
                          side=tk.LEFT, expand=True)

        # ---- 列映射卡片 ----
        self._sidebar_section_header(scroll_frame, "列映射", pad=pad)
        card = tk.Frame(scroll_frame, bg=COLORS["sidebar_surface"], highlightthickness=1,
                        highlightbackground=COLORS["sidebar_border"])
        card.pack(fill=tk.X, padx=pad, pady=(4, 0))

        for label, attr in [("X 轴（分类）", "x"), ("Y 轴（数值）", "y"), ("图例（分组）", "legend")]:
            row = tk.Frame(card, bg=COLORS["sidebar_surface"])
            row.pack(fill=tk.X, padx=inner_pad, pady=3)
            tk.Label(row, text=label, font=("Microsoft YaHei", 9),
                     bg=COLORS["sidebar_surface"], fg=COLORS["sidebar_text_secondary"],
                     width=10, anchor="w").pack(side=tk.LEFT)
            combo = ttk.Combobox(row, state="readonly", font=("Microsoft YaHei", 9),
                                 style="App.TCombobox")
            combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
            setattr(self, f"{attr}_col_combo", combo)

        # 底部留白
        tk.Frame(card, bg=COLORS["sidebar_surface"], height=4).pack()

        # ---- 图表设置卡片 ----
        self._sidebar_section_header(scroll_frame, "图表设置", pad=pad)
        card = tk.Frame(scroll_frame, bg=COLORS["sidebar_surface"], highlightthickness=1,
                        highlightbackground=COLORS["sidebar_border"])
        card.pack(fill=tk.X, padx=pad, pady=(4, 0))

        settings = [
            ("类型", "chart_type", list(CHART_TYPES.keys()), 0),
            ("主题", "theme", list(THEMES.keys()), 1),
            ("配色", "palette", list(PALETTES.keys()), 0),
        ]
        for label, attr, values, default_idx in settings:
            row = tk.Frame(card, bg=COLORS["sidebar_surface"])
            row.pack(fill=tk.X, padx=inner_pad, pady=3)
            tk.Label(row, text=label, font=("Microsoft YaHei", 9),
                     bg=COLORS["sidebar_surface"], fg=COLORS["sidebar_text_secondary"],
                     width=6, anchor="w").pack(side=tk.LEFT)
            combo = ttk.Combobox(row, state="readonly", font=("Microsoft YaHei", 9),
                                 values=values, style="App.TCombobox")
            combo.current(default_idx)
            combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
            setattr(self, f"{attr}_combo", combo)

        # 自定义标题
        row = tk.Frame(card, bg=COLORS["sidebar_surface"])
        row.pack(fill=tk.X, padx=inner_pad, pady=3)
        tk.Label(row, text="标题", font=("Microsoft YaHei", 9),
                 bg=COLORS["sidebar_surface"], fg=COLORS["sidebar_text_secondary"],
                 width=6, anchor="w").pack(side=tk.LEFT)
        self.title_var = tk.StringVar()
        self._sidebar_entry(row, self.title_var, "留空则自动生成")

        # 显示选项
        checks_row = tk.Frame(card, bg=COLORS["sidebar_surface"])
        checks_row.pack(fill=tk.X, padx=inner_pad, pady=(4, 4))
        self.show_xlabel_var = tk.BooleanVar(value=True)
        self.show_ylabel_var = tk.BooleanVar(value=True)
        self.show_legend_var = tk.BooleanVar(value=True)
        for var, text in [(self.show_xlabel_var, "X轴标题"),
                          (self.show_ylabel_var, "Y轴标题"),
                          (self.show_legend_var, "图例")]:
            cb = tk.Checkbutton(checks_row, text=text, variable=var,
                                font=("Microsoft YaHei", 9),
                                bg=COLORS["sidebar_surface"], fg=COLORS["sidebar_text_secondary"],
                                selectcolor=COLORS["sidebar_surface"],
                                activebackground=COLORS["sidebar_surface"],
                                activeforeground=COLORS["sidebar_text"])
            cb.pack(side=tk.LEFT, padx=(0, 8))

        tk.Frame(card, bg=COLORS["sidebar_surface"], height=2).pack()

        # ---- 字号卡片 ----
        self._sidebar_section_header(scroll_frame, "字号", pad=pad)
        card = tk.Frame(scroll_frame, bg=COLORS["sidebar_surface"], highlightthickness=1,
                        highlightbackground=COLORS["sidebar_border"])
        card.pack(fill=tk.X, padx=pad, pady=(4, 0))

        font_defaults = {"title": 15, "xlabel": 11, "ylabel": 11, "legend": 10}
        font_labels = {"title": "标题", "xlabel": "X轴标签", "ylabel": "Y轴标签", "legend": "图例"}
        for key in ("title", "xlabel", "ylabel", "legend"):
            row = tk.Frame(card, bg=COLORS["sidebar_surface"])
            row.pack(fill=tk.X, padx=inner_pad, pady=1)
            tk.Label(row, text=font_labels[key], font=("Microsoft YaHei", 9),
                     bg=COLORS["sidebar_surface"], fg=COLORS["sidebar_text_secondary"],
                     width=6, anchor="w").pack(side=tk.LEFT)
            var = tk.IntVar(value=font_defaults[key])
            setattr(self, f"font_{key}_var", var)

            # 加号
            self._btn_stepper(row, "+", lambda k=key: self._adjust_font(k, 1), side=tk.RIGHT)
            tk.Label(row, textvariable=var, font=("Microsoft YaHei", 9, "bold"),
                     bg=COLORS["sidebar_surface"], fg=COLORS["sidebar_text"],
                     width=3, anchor="center").pack(side=tk.RIGHT, padx=1)
            # 减号
            self._btn_stepper(row, "−", lambda k=key: self._adjust_font(k, -1), side=tk.RIGHT)

        tk.Frame(card, bg=COLORS["sidebar_surface"], height=2).pack()

        # ---- 操作按钮 ----
        btn_area = tk.Frame(scroll_frame, bg=COLORS["sidebar_bg"])
        btn_area.pack(fill=tk.X, padx=pad, pady=(14, 4))

        self.generate_btn = tk.Button(btn_area, text="生成透视图",
                                      font=("Microsoft YaHei", 11, "bold"),
                                      bg=COLORS["primary"], fg="#FFFFFF",
                                      relief="flat", cursor="hand2", bd=0,
                                      activebackground=COLORS["primary_hover"],
                                      activeforeground="#FFFFFF",
                                      padx=16, pady=10, command=self._generate_chart)
        self.generate_btn.pack(fill=tk.X)

        # 导出按钮
        export_row = tk.Frame(scroll_frame, bg=COLORS["sidebar_bg"])
        export_row.pack(fill=tk.X, padx=pad, pady=(6, 0))

        self._btn_secondary(export_row, "导出 Excel", self._export_excel,
                           side=tk.LEFT, padx=(0, 6))
        self._btn_secondary(export_row, "导出图片", self._export_image,
                           side=tk.LEFT)

        # 进度条
        self.progress = ttk.Progressbar(scroll_frame, style="Horizontal.TProgressbar",
                                        mode="indeterminate")
        self.progress.pack(fill=tk.X, padx=pad, pady=(12, 16))
        self.progress.pack_forget()

    def _sidebar_section_header(self, parent: tk.Frame, text: str, *, pad: int) -> None:
        tk.Label(parent, text=text, font=("Microsoft YaHei", 9, "bold"),
                 bg=COLORS["sidebar_bg"], fg=COLORS["sidebar_text_secondary"]).pack(
            anchor=tk.W, padx=pad, pady=(12, 1))

    def _sidebar_entry(self, parent: tk.Frame, textvar: tk.StringVar, placeholder: str) -> None:
        entry = tk.Entry(parent, textvariable=textvar, font=("Microsoft YaHei", 9),
                         bg=COLORS["sidebar_surface"], fg=COLORS["sidebar_text"],
                         insertbackground=COLORS["sidebar_text"], relief="flat", bd=0,
                         highlightthickness=1, highlightbackground=COLORS["sidebar_border"],
                         highlightcolor=COLORS["sidebar_accent"])
        entry.pack(fill=tk.X, padx=12, pady=(6, 2), ipady=5)

    def _btn_primary(self, parent: tk.Frame, text: str, command, **pack_kw) -> None:
        side = pack_kw.pop("side", tk.LEFT)
        expand = pack_kw.pop("expand", False)
        padx = pack_kw.pop("padx", 0)
        tk.Button(parent, text=text, font=("Microsoft YaHei", 9),
                  bg=COLORS["sidebar_accent"], fg="#FFFFFF", relief="flat",
                  cursor="hand2", bd=0,
                  activebackground="#4F46E5", activeforeground="#FFFFFF",
                  padx=12, pady=4, command=command).pack(
            side=side, fill=tk.X if expand else tk.NONE, padx=padx)

    def _btn_secondary(self, parent: tk.Frame, text: str, command, **pack_kw) -> None:
        side = pack_kw.pop("side", tk.LEFT)
        padx = pack_kw.pop("padx", 0)
        tk.Button(parent, text=text, font=("Microsoft YaHei", 9),
                  bg=COLORS["sidebar_surface"], fg=COLORS["sidebar_text_secondary"],
                  relief="flat", cursor="hand2", bd=0, highlightthickness=1,
                  highlightbackground=COLORS["sidebar_border"],
                  highlightcolor=COLORS["sidebar_accent"],
                  activebackground=COLORS["sidebar_surface"],
                  activeforeground=COLORS["sidebar_text"],
                  padx=10, pady=5, command=command).pack(
            side=side, fill=tk.X, expand=True, padx=padx)

    def _btn_stepper(self, parent: tk.Frame, text: str, command, **pack_kw) -> None:
        side = pack_kw.pop("side", tk.RIGHT)
        tk.Button(parent, text=text, font=("Microsoft YaHei", 9, "bold"),
                  bg=COLORS["sidebar_border"], fg=COLORS["sidebar_text"],
                  relief="flat", bd=0, cursor="hand2",
                  activebackground=COLORS["sidebar_accent"],
                  activeforeground="#FFFFFF",
                  width=2, padx=0, pady=0, command=command).pack(side=side)

    # ========== 主内容区 ==========
    def _build_main(self, parent: tk.Frame) -> None:
        main = tk.Frame(parent, bg=COLORS["bg"])
        main.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(main, style="App.TNotebook", padding=[12, 8, 12, 8])
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 标签下划线指示器
        tab_frame = tk.Frame(main, bg=COLORS["border"], height=2)
        tab_frame.pack(fill=tk.X, padx=12)

        self.data_tab = tk.Frame(self.notebook, bg=COLORS["surface"])
        self.notebook.add(self.data_tab, text="  数据预览  ")
        self._build_data_panel(self.data_tab)

        self.pivot_tab = tk.Frame(self.notebook, bg=COLORS["surface"])
        self.notebook.add(self.pivot_tab, text="  透视表  ")
        self._build_pivot_panel(self.pivot_tab)

        self.chart_tab = tk.Frame(self.notebook, bg=COLORS["surface"])
        self.notebook.add(self.chart_tab, text="  图表  ")
        self._build_chart_panel(self.chart_tab)

    def _build_data_panel(self, parent: tk.Frame) -> None:
        """数据预览面板"""
        # 工具栏
        toolbar = tk.Frame(parent, bg=COLORS["surface"], height=40)
        toolbar.pack(fill=tk.X, padx=14, pady=(8, 0))
        toolbar.pack_propagate(False)

        self.data_info_var = tk.StringVar(value="尚未加载数据")
        tk.Label(toolbar, textvariable=self.data_info_var,
                 font=("Microsoft YaHei", 11, "bold"),
                 bg=COLORS["surface"], fg=COLORS["text"]).pack(side=tk.LEFT, pady=6)
        self.data_row_var = tk.StringVar()
        tk.Label(toolbar, textvariable=self.data_row_var, font=("Microsoft YaHei", 9),
                 bg=COLORS["surface"], fg=COLORS["text_secondary"]).pack(side=tk.RIGHT, pady=6)

        tk.Frame(parent, bg=COLORS["border_light"], height=1).pack(fill=tk.X, padx=14)

        # 空状态 / 表格容器
        self.data_container = tk.Frame(parent, bg=COLORS["surface"])
        self.data_container.pack(fill=tk.BOTH, expand=True, padx=14, pady=(4, 12))

        self.data_tree = ttk.Treeview(self.data_container, style="App.Treeview",
                                      show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(self.data_container, orient=tk.VERTICAL, command=self.data_tree.yview)
        hsb = ttk.Scrollbar(self.data_container, orient=tk.HORIZONTAL, command=self.data_tree.xview)
        self.data_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.data_tree.tag_configure("odd", background=COLORS["row_alt"])
        self.data_tree.tag_configure("even", background=COLORS["surface"])

        self._show_empty(self.data_container, "file", "尚未加载数据",
                         "点击侧边栏「浏览…」选择 Excel 文件，\n或直接将文件拖入此窗口")

    def _build_pivot_panel(self, parent: tk.Frame) -> None:
        """透视表面板"""
        toolbar = tk.Frame(parent, bg=COLORS["surface"], height=40)
        toolbar.pack(fill=tk.X, padx=14, pady=(8, 0))
        toolbar.pack_propagate(False)

        self.pivot_info_var = tk.StringVar(value="尚未生成透视表")
        tk.Label(toolbar, textvariable=self.pivot_info_var,
                 font=("Microsoft YaHei", 11, "bold"),
                 bg=COLORS["surface"], fg=COLORS["text"]).pack(side=tk.LEFT, pady=6)
        self.pivot_row_var = tk.StringVar()
        tk.Label(toolbar, textvariable=self.pivot_row_var, font=("Microsoft YaHei", 9),
                 bg=COLORS["surface"], fg=COLORS["text_secondary"]).pack(side=tk.RIGHT, pady=6)

        tk.Frame(parent, bg=COLORS["border_light"], height=1).pack(fill=tk.X, padx=14)

        self.pivot_container = tk.Frame(parent, bg=COLORS["surface"])
        self.pivot_container.pack(fill=tk.BOTH, expand=True, padx=14, pady=(4, 12))

        self.pivot_tree = ttk.Treeview(self.pivot_container, style="App.Treeview",
                                       show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(self.pivot_container, orient=tk.VERTICAL, command=self.pivot_tree.yview)
        hsb = ttk.Scrollbar(self.pivot_container, orient=tk.HORIZONTAL, command=self.pivot_tree.xview)
        self.pivot_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.pivot_tree.tag_configure("odd", background=COLORS["row_alt"])
        self.pivot_tree.tag_configure("even", background=COLORS["surface"])

        self._show_empty(self.pivot_container, "chart", "尚未生成透视表",
                         "配置列映射后点击「生成透视图」")

    def _build_chart_panel(self, parent: tk.Frame) -> None:
        """图表面板"""
        self.chart_container = tk.Frame(parent, bg=COLORS["chart_bg"])
        self.chart_container.pack(fill=tk.BOTH, expand=True, padx=14, pady=12)

        self._show_empty(self.chart_container, "chart", "尚未生成图表",
                         "配置列映射后点击「生成透视图」")

    # ========== 空状态 ==========
    def _show_empty(self, container: tk.Frame, icon: str, title: str, subtitle: str) -> None:
        for w in container.winfo_children():
            w.pack_forget()
            w.grid_forget()
            w.place_forget()

        placeholder = tk.Frame(container, bg=COLORS["surface"])
        placeholder.pack(fill=tk.BOTH, expand=True)

        inner = tk.Frame(placeholder, bg=COLORS["surface"])
        inner.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        # 图标区域（Canvas 绘制）
        ico = tk.Canvas(inner, width=72, height=72, bg=COLORS["surface"], highlightthickness=0)
        ico.pack(anchor=tk.CENTER)
        self._draw_empty_icon(ico, icon)

        tk.Label(inner, text=title, font=("Microsoft YaHei", 12, "bold"),
                 bg=COLORS["surface"], fg=COLORS["text"]).pack(anchor=tk.CENTER, pady=(12, 4))
        tk.Label(inner, text=subtitle, font=("Microsoft YaHei", 9),
                 bg=COLORS["surface"], fg=COLORS["text_muted"],
                 justify=tk.CENTER).pack(anchor=tk.CENTER)

    def _draw_empty_icon(self, canvas: tk.Canvas, icon: str) -> None:
        w, h = 72, 72
        c = COLORS["border_light"]
        a = COLORS["primary_light"]

        if icon == "file":
            # 文件图标
            canvas.create_rectangle(16, 6, 56, 66, fill=c, outline=a, width=2, stipple="")
            canvas.create_polygon(40, 6, 56, 6, 56, 22, 40, 22, fill=a, outline="")
            canvas.create_line(24, 34, 48, 34, fill=a, width=2)
            canvas.create_line(24, 42, 48, 42, fill=a, width=2)
            canvas.create_line(24, 50, 36, 50, fill=a, width=2)
        elif icon == "chart":
            # 图表图标
            canvas.create_rectangle(14, 18, 26, 58, fill=COLORS["primary"], outline="", stipple="")
            canvas.create_rectangle(30, 30, 42, 58, fill=COLORS["primary"], outline="", stipple="")
            canvas.create_rectangle(46, 10, 58, 58, fill=COLORS["primary"], outline="", stipple="")
            canvas.create_line(14, 58, 60, 58, fill=COLORS["border"], width=2)
            canvas.create_line(14, 58, 14, 60, fill=COLORS["border"], width=2)

    # ========== 底部状态栏 ==========
    def _build_statusbar(self) -> None:
        bar = tk.Frame(self.root, bg=COLORS["surface"], height=28)
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)

        tk.Frame(bar, bg=COLORS["border"], height=1).pack(fill=tk.X)

        inner = tk.Frame(bar, bg=COLORS["surface"])
        inner.pack(fill=tk.X, padx=14)

        self.status_var = tk.StringVar(value="就绪")
        tk.Label(inner, textvariable=self.status_var, font=("Microsoft YaHei", 9),
                 bg=COLORS["surface"], fg=COLORS["text_secondary"]).pack(side=tk.LEFT, pady=3)

        tk.Label(inner, text="Ctrl+O 打开  |  Ctrl+G 生成  |  Ctrl+E 导出",
                 font=("Microsoft YaHei", 9),
                 bg=COLORS["surface"], fg=COLORS["text_muted"]).pack(side=tk.RIGHT, pady=3)

    # ========== 快捷键 ==========
    def _bind_shortcuts(self) -> None:
        self.root.bind("<Control-o>", lambda e: self._import_files())
        self.root.bind("<Control-O>", lambda e: self._import_files())
        self.root.bind("<Control-g>", lambda e: self._generate_chart())
        self.root.bind("<Control-G>", lambda e: self._generate_chart())
        self.root.bind("<Control-e>", lambda e: self._export_excel())
        self.root.bind("<Control-E>", lambda e: self._export_excel())
        self.root.bind("<Control-s>", lambda e: self._export_image())
        self.root.bind("<Control-S>", lambda e: self._export_image())

    # ========== 拖拽加载 ==========
    def _setup_drop_target(self) -> None:
        try:
            from tkinter import Tkdnd  # type: ignore
            self.root.drop_target_register("DND_Files")
            self.root.dnd_bind("<<Drop>>", self._on_drop)
        except ImportError:
            pass

        # 回退：绑定窗口 Enter 作为视觉提示
        self.root.bind("<Map>", lambda e: None)

    def _on_drop(self, event) -> None:
        try:
            data = event.data
            if data.startswith("{") and data.endswith("}"):
                data = data[1:-1]
            path = data.strip().replace("/", "\\")
            if path.lower().endswith((".xlsx", ".xls")):
                self._import_single_file(path)
        except Exception:
            pass

    # ========== 表格辅助 ==========
    def _show_tree_placeholder(self, tree: ttk.Treeview, message: str) -> None:
        tree.delete(*tree.get_children())
        tree["columns"] = ("_msg",)
        tree.heading("_msg", text="")
        tree.column("_msg", width=400, anchor="center")
        tree.insert("", tk.END, values=(message,))

    def _populate_tree(self, tree: ttk.Treeview, df: pd.DataFrame, max_rows: int = 5000) -> None:
        tree.delete(*tree.get_children())
        if df.empty:
            self._show_tree_placeholder(tree, "无数据")
            return

        columns = [str(c) for c in df.columns]
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=max(80, min(200, len(col) * 16 + 20)),
                       minwidth=60, anchor="center")

        display_df = df.head(max_rows) if len(df) > max_rows else df
        for i, (_, row) in enumerate(display_df.iterrows()):
            values = [str(v) if pd.notna(v) else "" for v in row]
            tag = "odd" if i % 2 == 1 else "even"
            tree.insert("", tk.END, values=values, tags=(tag,))

        if len(df) > max_rows:
            tree.insert("", tk.END, values=(f"... 仅显示前 {max_rows} 行，共 {len(df)} 行",))

    # ========== 文件操作 ==========
    def _import_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件（可多选）",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
            initialdir=Path.cwd(),
        )
        if not paths:
            return

        try:
            self._set_status("正在导入文件...", "loading")
            self._show_progress(True)

            if self.dataframe is not None:
                # 增量追加：先合并再对整体做日期富化
                combined, added_rows = append_to_existing(self.dataframe, list(paths))
                self.dataframe = enrich_date_columns(combined)
                self.imported_files.extend(Path(p).name for p in paths)
            else:
                # 首次导入
                raw_merged, filenames = import_files(list(paths))
                self.dataframe = enrich_date_columns(raw_merged)
                self.imported_files = filenames

            self.columns = [str(c) for c in self.dataframe.columns.tolist()]
            self._refresh_columns()

            # 保存到 temp
            self.temp_file = self._save_to_temp(self.dataframe)

            # 显示数据表格
            self._show_tree(self.data_container, self.data_tree, self.dataframe)
            self.data_info_var.set(f"已导入 {len(self.imported_files)} 个文件")
            self.data_row_var.set(f"{len(self.dataframe)} 行 × {len(self.columns)} 列")

            self.header_file_var.set(f"数据库  ({len(self.dataframe):,} 行)")
            self._set_header_dot(COLORS["success"])

            self.save_source_btn.configure(state=tk.NORMAL)
            self.close_db_btn.configure(state=tk.NORMAL)
            self.db_info_var.set(
                f"{len(self.imported_files)} 个文件 · "
                f"{len(self.dataframe):,} 行 × {len(self.columns)} 列")

            self._set_status(
                f"已导入 {len(self.imported_files)} 个文件"
                f"（{len(self.dataframe):,} 行 × {len(self.columns)} 列）", "success")
            self.notebook.select(0)
            self._update_pivot_empty()
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
            self._set_status("导入失败", "error")
            self._set_header_dot(COLORS["error"])
        finally:
            self._show_progress(False)

    def _import_single_file(self, path: str) -> None:
        """拖拽/命令行参数单文件导入 — 委托给 _import_files 相同的底层逻辑"""
        try:
            self._set_status("正在导入文件...", "loading")
            self._show_progress(True)

            if self.dataframe is not None:
                combined, added_rows = append_to_existing(self.dataframe, [path])
                self.dataframe = enrich_date_columns(combined)
                self.imported_files.append(Path(path).name)
            else:
                raw_merged, filenames = import_files([path])
                self.dataframe = enrich_date_columns(raw_merged)
                self.imported_files = filenames

            self.columns = [str(c) for c in self.dataframe.columns.tolist()]
            self._refresh_columns()
            self.temp_file = self._save_to_temp(self.dataframe)

            self._show_tree(self.data_container, self.data_tree, self.dataframe)
            self.data_info_var.set(f"已导入 {len(self.imported_files)} 个文件")
            self.data_row_var.set(f"{len(self.dataframe)} 行 × {len(self.columns)} 列")

            self.header_file_var.set(f"数据库  ({len(self.dataframe):,} 行)")
            self._set_header_dot(COLORS["success"])
            self.save_source_btn.configure(state=tk.NORMAL)
            self.close_db_btn.configure(state=tk.NORMAL)
            self.db_info_var.set(
                f"{len(self.imported_files)} 个文件 · "
                f"{len(self.dataframe):,} 行 × {len(self.columns)} 列")
            self._set_status(f"已导入 {Path(path).name}", "success")
            self.notebook.select(0)
            self._update_pivot_empty()
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
            self._set_status("导入失败", "error")
        finally:
            self._show_progress(False)

    def _close_database(self) -> None:
        if not messagebox.askyesno("确认关闭", "关闭数据库将清空所有已导入的数据，确定继续？"):
            return
        self.dataframe = None
        self.columns = []
        self.imported_files = []
        self.current_figure = None
        self.current_pivot = None
        self.temp_file = None

        if hasattr(self, "x_col_combo"):
            self.x_col_combo["values"] = []
            self.x_col_combo.set("")
            self.y_col_combo["values"] = []
            self.y_col_combo.set("")
            self.legend_col_combo["values"] = []
            self.legend_col_combo.set("")

        self._show_empty(self.data_container, "file", "尚未导入数据",
                         "点击「导入数据」选择 Excel 文件，可多选")
        self.data_info_var.set("尚未导入数据")
        self.data_row_var.set("")
        self._update_pivot_empty()

        # 清空图表区
        for w in self.chart_container.winfo_children():
            w.destroy()
        self._show_empty(self.chart_container, "chart", "尚未生成图表",
                         "导入数据后点击「生成透视图」")

        self.header_file_var.set("未导入数据")
        self._set_header_dot("#D1D5DB")
        self.db_info_var.set("未导入数据")

        self.save_source_btn.configure(state=tk.DISABLED)
        self.close_db_btn.configure(state=tk.DISABLED)

        # 清理 temp 文件
        temp_dir = Path.cwd() / "temp"
        if temp_dir.exists():
            for f in temp_dir.glob("*.xlsx"):
                try:
                    f.unlink()
                except Exception:
                    pass

        self._set_status("数据库已关闭", "info")

    def _show_tree(self, container: tk.Frame, tree: ttk.Treeview,
                  df: pd.DataFrame) -> None:
        for w in container.winfo_children():
            w.pack_forget()
            w.grid_forget()
            w.place_forget()

        vsb = ttk.Scrollbar(container, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self._populate_tree(tree, df)

    def _save_to_temp(self, df: pd.DataFrame) -> Path:
        temp_dir = Path.cwd() / "temp"
        temp_dir.mkdir(exist_ok=True)
        output = temp_dir / "merged_data.xlsx"
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="源数据", index=False)
        return output

    def _save_source(self) -> None:
        if self.dataframe is None:
            messagebox.showwarning("提示", "请先导入数据。")
            return
        default_name = "merged_data.xlsx"
        path = filedialog.asksaveasfilename(
            title="保存源数据",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            self.dataframe.to_excel(path, index=False, engine="openpyxl")
            self._set_status(f"源数据已保存: {Path(path).name}", "success")
            messagebox.showinfo("保存成功", f"已保存到:\n{path}")
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))

    def _update_pivot_empty(self) -> None:
        self._show_empty(self.pivot_container, "chart", "尚未生成透视表",
                         "配置列映射后点击「生成透视图」")

    def _refresh_columns(self) -> None:
        if not self.columns:
            return
        self.x_col_combo["values"] = self.columns
        self.y_col_combo["values"] = self.columns
        self.legend_col_combo["values"] = ["— 无 —"] + self.columns
        if self.x_col_combo.get() not in self.columns:
            self.x_col_combo.current(0)
        if self.y_col_combo.get() not in self.columns:
            self.y_col_combo.current(min(1, len(self.columns) - 1))
        if self.legend_col_combo.get() not in (["— 无 —"] + self.columns):
            self.legend_col_combo.current(min(3, len(self.columns)))

    # ========== 图表生成 ==========
    def _generate_chart(self) -> None:
        if self.dataframe is None:
            messagebox.showwarning("提示", "请先加载 Excel 文件。")
            return

        x_col = self.x_col_combo.get()
        y_col = self.y_col_combo.get()
        legend_col = self.legend_col_combo.get()
        if not x_col or not y_col:
            messagebox.showwarning("提示", "请选择 X 轴列和 Y 轴列。")
            return
        if legend_col == "— 无 —":
            legend_col = ""

        chart_label = self.chart_type_combo.get()
        chart_type = CHART_TYPES.get(chart_label, "stacked_bar")
        theme = self.theme_combo.get()
        palette = self.palette_combo.get()
        custom_title = self.title_var.get().strip()
        show_xlabel = self.show_xlabel_var.get()
        show_ylabel = self.show_ylabel_var.get()
        show_legend = self.show_legend_var.get()
        font_title = self.font_title_var.get()
        font_xlabel = self.font_xlabel_var.get()
        font_ylabel = self.font_ylabel_var.get()
        font_legend = self.font_legend_var.get()

        self.generate_btn.configure(state=tk.DISABLED, text="生成中…", bg=COLORS["primary_hover"])
        self._set_status("正在生成图表...", "loading")
        self._show_progress(True)

        def _run():
            try:
                pivot, figure = create_chart(
                    self.dataframe,  # type: ignore[arg-type]
                    x_col, y_col, legend_col,
                    chart_type=chart_type, theme=theme, palette=palette,
                    custom_title=custom_title, show_xlabel=show_xlabel,
                    show_ylabel=show_ylabel, show_legend=show_legend,
                    font_title=font_title, font_xlabel=font_xlabel,
                    font_ylabel=font_ylabel, font_legend=font_legend,
                )
                self.current_figure = figure
                self.current_pivot = pivot
                self.root.after(0, lambda: self._on_chart_ready(chart_label, theme, palette))
            except Exception as exc:
                self.root.after(0, lambda e=exc: self._on_chart_error(e))

        threading.Thread(target=_run, daemon=True).start()

    def _on_chart_ready(self, chart_label: str, theme: str, palette: str) -> None:
        self._embed_chart()
        self._update_pivot_view()
        self._set_status(
            f"图表已生成 — {chart_label} | {theme} | {palette}", "success")
        self.notebook.select(2)
        self.generate_btn.configure(state=tk.NORMAL, text="生成透视图",
                                    bg=COLORS["primary"])
        self._show_progress(False)

    def _update_pivot_view(self) -> None:
        if self.current_pivot is not None:
            self._show_tree(self.pivot_container, self.pivot_tree,
                           self.current_pivot.reset_index())
            self.pivot_info_var.set("数据透视表")
            self.pivot_row_var.set(
                f"{self.current_pivot.shape[0]} 行 × {self.current_pivot.shape[1]} 列")
        else:
            self._show_empty(self.pivot_container, "chart", "无透视表",
                             "当前图表类型不生成透视表")

    def _on_chart_error(self, exc: Exception) -> None:
        messagebox.showerror("生成失败", str(exc))
        self._set_status("生成失败", "error")
        self.generate_btn.configure(state=tk.NORMAL, text="生成透视图",
                                    bg=COLORS["primary"])
        self._show_progress(False)

    # ========== 图表嵌入 ==========
    def _embed_chart(self) -> None:
        for widget in self.chart_container.winfo_children():
            widget.destroy()

        if self.current_figure is None:
            return

        canvas = FigureCanvasTkAgg(self.current_figure, master=self.chart_container)
        canvas.draw()
        widget = canvas.get_tk_widget()
        widget.pack(fill=tk.BOTH, expand=True)

        for seq in ("<Button-1>", "<Button-2>", "<Button-3>",
                     "<B1-Motion>", "<B2-Motion>", "<B3-Motion>",
                     "<MouseWheel>", "<Button-4>", "<Button-5>",
                     "<Motion>", "<Enter>", "<Leave>"):
            widget.bind(seq, lambda e: "break")
        self._canvas = canvas

    # ========== 导出 ==========
    def _get_output_stem(self) -> str:
        return "merged_data"

    @staticmethod
    def _get_timestamp() -> str:
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    def _export_excel(self) -> None:
        if self.current_pivot is None or self.dataframe is None:
            messagebox.showwarning("提示", "请先生成透视图。")
            return
        default_name = f"{self._get_output_stem()}_透视结果_{self._get_timestamp()}.xlsx"
        path = filedialog.asksaveasfilename(
            title="导出透视表",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                self.current_pivot.to_excel(writer, sheet_name="透视表")
                self.dataframe.to_excel(writer, sheet_name="源数据", index=False)
            self._set_status(f"Excel 已导出: {Path(path).name}", "success")
            messagebox.showinfo("导出成功", f"已保存到:\n{path}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def _export_image(self) -> None:
        if self.current_figure is None:
            messagebox.showwarning("提示", "请先生成透视图。")
            return
        default_name = f"{self._get_output_stem()}_图表_{self._get_timestamp()}.png"
        path = filedialog.asksaveasfilename(
            title="导出图表图片",
            defaultextension=".png",
            initialfile=default_name,
            filetypes=[("PNG 图片", "*.png"), ("JPEG 图片", "*.jpg"), ("PDF 文档", "*.pdf")],
        )
        if not path:
            return
        try:
            self.current_figure.savefig(path, dpi=200, bbox_inches="tight")
            self._set_status(f"图片已导出: {Path(path).name}", "success")
            messagebox.showinfo("导出成功", f"已保存到:\n{path}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    # ========== 工具 ==========
    def _adjust_font(self, key: str, delta: int) -> None:
        var = getattr(self, f"font_{key}_var")
        new_val = var.get() + delta
        if 6 <= new_val <= 36:
            var.set(new_val)

    def _set_status(self, text: str, level: str = "info") -> None:
        self.status_var.set(text)
        dot_colors = {"info": "#D1D5DB", "loading": "#F59E0B",
                       "success": "#10B981", "error": "#EF4444"}
        self._set_header_dot(dot_colors.get(level, "#D1D5DB"))

    def _show_progress(self, show: bool) -> None:
        if show:
            self.progress.pack(fill=tk.X, padx=14, pady=(12, 16))
            self.progress.start(10)
        else:
            self.progress.stop()
            self.progress.pack_forget()


# ---------- 入口 ----------
def main() -> None:
    root = tk.Tk()

    # 设置窗口图标（如可用）
    try:
        root.iconbitmap(default="")
    except Exception:
        pass

    app = PivotChartApp(root)

    if len(sys.argv) > 1:
        arg = sys.argv[1]
        src = Path(arg)
        if src.exists() and src.suffix.lower() in (".xlsx", ".xls"):
            app._import_single_file(str(src.resolve()))

    root.mainloop()


if __name__ == "__main__":
    main()
