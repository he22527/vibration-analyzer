#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""振動連續監測資料分析工具 v2.0"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import os, glob, datetime, threading, queue
from concurrent.futures import ThreadPoolExecutor, as_completed

# ════════════════════════════════════════════════════════
#  常數
# ════════════════════════════════════════════════════════

FREQ_BANDS: list[tuple[float, str]] = [
    (1,"1 Hz"),(1.25,"1.25 Hz"),(1.6,"1.6 Hz"),(2,"2 Hz"),(2.5,"2.5 Hz"),
    (3.15,"3.15 Hz"),(4,"4 Hz"),(5,"5 Hz"),(6.3,"6.3 Hz"),(8,"8 Hz"),
    (10,"10 Hz"),(12.5,"12.5 Hz"),(16,"16 Hz"),(20,"20 Hz"),(25,"25 Hz"),
    (31.5,"31.5 Hz"),(40,"40 Hz"),(50,"50 Hz"),(63,"63 Hz"),(80,"80 Hz"),
    (100,"100 Hz"),(125,"125 Hz"),(160,"160 Hz"),(200,"200 Hz"),
    (250,"250 Hz"),(315,"315 Hz"),
]
FREQ_NAMES  = [n for _, n in FREQ_BANDS]
AXES        = ["X", "Y", "Z"]
STATS       = ["Leq", "L5", "L10", "L50", "L90", "L95", "Lmax"]
DEFAULT_ROOT = r"C:\Users\AC717\背景振動量測數據"

# 各統計量對應線型（VC 曲線分頁多統計量疊圖用）
STAT_LINESTYLES: dict[str, object] = {
    "Leq":  "-",
    "L5":   "--",
    "L10":  (0, (4, 1, 1, 1)),
    "L50":  ":",
    "L90":  (0, (3, 2)),
    "L95":  (0, (1, 1)),
    "Lmax": (0, (6, 2)),
}

AXIS_COLORS = {"X":"royalblue","Y":"forestgreen","Z":"crimson","XYZ":"darkorchid"}
AXIS_LABELS = {"X":"X 軸","Y":"Y 軸","Z":"Z 軸","XYZ":"XYZ 合成"}

INTERVAL_OPTIONS: list[tuple[str,float,str|None]] = [
    ("原始 (0.1s)",0.1,None),("1 秒",1,"1s"),("1 分鐘",60,"1min"),
    ("5 分鐘",300,"5min"),("10 分鐘",600,"10min"),
    ("30 分鐘",1800,"30min"),("60 分鐘",3600,"60min"),
]
INTERVAL_LABELS = [l for l,_,_ in INTERVAL_OPTIONS]

# (category, label, SI 說明, 線性換算係數 from base, dB 基準值)
# base: m/s²（加速度）或 m/s（速度，需先做 v=a/2πf）
UNIT_DEFS: dict[str, tuple] = {
    "m/s²":   ("acc","m/s²",  "SI 加速度基本單位",                       1.0,       None),
    "g":      ("acc","g",     "重力加速度  1 g = 9.80665 m/s²",          1/9.80665, None),
    "Gal":    ("acc","Gal",   "公制  1 Gal = 0.01 m/s² = 1 cm/s²",       100.0,     None),
    "dB_acc": ("acc","dB",    "加速度級  ISO 1683  基準 a₀ = 10⁻⁶ m/s²", 1.0,       1e-6),
    "m/s":    ("vel","m/s",   "SI 速度基本單位",                          1.0,       None),
    "mm/s":   ("vel","mm/s",  "公制  1 m/s = 1000 mm/s",                  1000.0,    None),
    "ips":    ("vel","ips",   "英制  1 m/s ≈ 39.3701 in/s",              39.3701,   None),
    "dB_vel": ("vel","dB",    "速度級  基準 v₀ = 2.54×10⁻⁸ m/s", 1.0, 2.54e-8),
}

# ── Ungar & Gordon（1991）VC 曲線 ──────────────────────────
# 平坦段限值（μm/s RMS，8–80 Hz）
# 數值由 dB ref 1×10⁻⁶ in/s 轉換：v_μm/s = 10^(dB/20) × 0.0254
# 工廠=90 dB, 辦公室=84, 住宅=78, 手術室=72,
# VC-A=66, VC-B=60, VC-C=54, VC-D=48, VC-E=42
VC_CURVES: dict[str, float] = {
    "工廠 (ISO)":   803.2,
    "辦公室 (ISO)": 402.6,
    "住宅 (ISO)":   201.8,
    "手術室 (ISO)": 101.1,
    "VC-A":          50.7,
    "VC-B":          25.4,
    "VC-C":          12.7,
    "VC-D":           6.38,
    "VC-E":           3.20,
}
VC_CURVES_DB: dict[str, int] = {
    "工廠 (ISO)": 90, "辦公室 (ISO)": 84, "住宅 (ISO)": 78,
    "手術室 (ISO)": 72, "VC-A": 66, "VC-B": 60,
    "VC-C": 54, "VC-D": 48, "VC-E": 42,
}
VC_DB_REF_UMS = 0.0254  # 1 μin/s in μm/s（= 10⁻⁶ in/s = 2.54×10⁻⁸ m/s）

# VC 準則曲線繪圖用頻率（4–80 Hz，依 Ungar & Gordon 標準）
VC_PLOT_FREQS = [4, 5, 6.3, 8, 10, 12.5, 16, 20, 25, 31.5, 40, 50, 63, 80]

# 量測數據繪圖用頻率（1–100 Hz，用於圖表顯示）
VC_MEAS_FREQS = [1, 1.25, 1.6, 2, 2.5, 3.15, 4, 5, 6.3, 8, 10, 12.5, 16,
                 20, 25, 31.5, 40, 50, 63, 80, 100]

# VC 準則曲線顏色：ISO 人體感知曲線用深灰階，VC-A～E 用深色系
# 量測數據另以 AXIS_COLORS 的鮮豔色顯示，形成明確對比
VC_COLORS = [
    "#2a2a2a",   # 工廠 (ISO)   — 近黑
    "#484848",   # 辦公室 (ISO) — 深灰
    "#636363",   # 住宅 (ISO)   — 中深灰
    "#7d7d7d",   # 手術室 (ISO) — 中灰
    "#1a3f7a",   # VC-A         — 深藍
    "#1a6070",   # VC-B         — 深青
    "#1a7030",   # VC-C         — 深綠
    "#7a3d10",   # VC-D         — 深棕
    "#7a0f0f",   # VC-E         — 深紅
]

# 等級排序：由最嚴格 → 最寬鬆
VC_ORDER = ["VC-E","VC-D","VC-C","VC-B","VC-A",
            "手術室 (ISO)","住宅 (ISO)","辦公室 (ISO)","工廠 (ISO)"]

# 各等級適用設備說明
VC_DESCRIPTIONS: dict[str, str] = {
    "工廠 (ISO)":   "800 μm/s　一般工廠作業（人員步行可明顯感受）",
    "辦公室 (ISO)": "400 μm/s　辦公室、電腦終端機",
    "住宅 (ISO)":   "200 μm/s　住宅、劇院",
    "手術室 (ISO)": "100 μm/s　手術室、醫院",
    "VC-A":          " 50 μm/s　光學顯微鏡（100× 以下）、微量天平",
    "VC-B":          " 25 μm/s　電子顯微鏡（TEM/SEM）、雷射干涉儀",
    "VC-C":          "12.5 μm/s　微影設備（≥3 μm 線寬）、核磁共振儀",
    "VC-D":          "   6 μm/s　半導體步進曝光機（晶圓步進機）",
    "VC-E":          "   3 μm/s　長光程雷射、重力波量測、奈米技術",
}

# 等級評定結果文字顏色
VC_GRADE_COLOR: dict[str, str] = {
    "VC-E":         "#1a6e1a",
    "VC-D":         "#2e8b57",
    "VC-C":         "#3a7abf",
    "VC-B":         "#6060bf",
    "VC-A":         "#8b5a00",
    "手術室 (ISO)": "#cc8800",
    "住宅 (ISO)":   "#cc6600",
    "辦公室 (ISO)": "#cc3300",
    "工廠 (ISO)":   "#bb1111",
    "超過工廠 (ISO)": "#990000",
}


def _vc_curve_val(flat_limit: float, f: float) -> float:
    """單一頻率的 VC 曲線值（μm/s）。
    4–8 Hz：每 1/3 八度音帶 +2 dB（10^(2/20) ≈ ×1.25893 / step）。
    ≥ 8 Hz：平坦段。"""
    if f >= 8.0:
        return flat_limit
    steps = round(np.log2(8.0 / f) * 3)  # 1/3 octave steps from 8 Hz
    return flat_limit * 10 ** (steps * 2 / 20)


def _vc_curve_vals(flat_limit: float, freqs: list = VC_PLOT_FREQS) -> list:
    return [_vc_curve_val(flat_limit, f) for f in freqs]


def _vc_grade(band_vels: list[tuple[float, float]]) -> str:
    """逐頻帶對照 VC 曲線值評定等級（Ungar & Gordon, 1991）。"""
    for name in VC_ORDER:
        flat = VC_CURVES[name]
        if all(vel <= _vc_curve_val(flat, fv) for fv, vel in band_vels):
            return name
    return "超過工廠 (ISO)"


# ════════════════════════════════════════════════════════
#  工具函式
# ════════════════════════════════════════════════════════

def calc_stats(vals: np.ndarray) -> dict:
    v = vals[np.isfinite(vals) & (vals > 0)]
    if len(v) == 0:
        return {s: np.nan for s in STATS}
    return {
        "Leq":  float(np.sqrt(np.mean(v**2))),
        "L5":   float(np.percentile(v, 95)),
        "L10":  float(np.percentile(v, 90)),
        "L50":  float(np.percentile(v, 50)),
        "L90":  float(np.percentile(v, 10)),
        "L95":  float(np.percentile(v, 5)),
        "Lmax": float(np.max(v)),
    }

def read_rnd_file(path: str, needed_cols: list[str]) -> pd.DataFrame | None:
    try:
        df = pd.read_csv(path, skiprows=1,
                         na_values=["--","UN","OL",""], dtype=str, low_memory=False)
        df.columns = [c.strip() for c in df.columns]
        if "Start Time" not in df.columns:
            return None
        df["Start Time"] = pd.to_datetime(
            df["Start Time"].str.strip(),
            format="%Y/%m/%d %H:%M:%S.%f", errors="coerce")
        df.dropna(subset=["Start Time"], inplace=True)
        keep = ["Start Time"] + [c for c in needed_cols if c in df.columns]
        df = df[keep].copy()
        for c in keep[1:]:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
        return df if len(df) > 0 else None
    except Exception:
        return None


# ════════════════════════════════════════════════════════
#  主應用程式
# ════════════════════════════════════════════════════════

class VibrationApp:
    RAW_PAGE_SIZE = 500

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("振動連續監測資料分析工具 v2.0")
        self.root.geometry("1500x900")
        self.root.minsize(1150, 720)

        self._df: pd.DataFrame | None = None
        self._results: dict | None    = None
        self._params:  dict | None    = None
        self._queue: queue.Queue      = queue.Queue()
        self._raw_page                = 0
        self._raw_total_pages         = 0
        self._raw_display_df: pd.DataFrame | None = None

        self._build_ui()
        self._poll_queue()

    # ─────────────── UI 框架 ─────────────────────────────

    def _build_ui(self):
        pw = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        pw.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        left = ttk.Frame(pw, width=340)
        left.pack_propagate(False)
        pw.add(left, weight=0)
        right = ttk.Frame(pw)
        pw.add(right, weight=1)
        self._build_controls(left)
        self._build_results(right)

    # ─────────────── 左側控制面板 ────────────────────────

    def _build_controls(self, parent):
        pad = dict(padx=6, pady=3)
        # 量測數據
        g = ttk.LabelFrame(parent, text="量測數據", padding=5)
        g.pack(fill=tk.X, **pad)
        self.var_dir = tk.StringVar(value=DEFAULT_ROOT)
        ttk.Entry(g, textvariable=self.var_dir).pack(fill=tk.X)
        ttk.Button(g, text="瀏覽...", command=self._browse_dir).pack(
            anchor=tk.E, pady=(2,0))

        # 日期 / 時間範圍
        g = ttk.LabelFrame(parent, text="日期 / 時間範圍", padding=5)
        g.pack(fill=tk.X, **pad)
        for lbl, attr, val in [
            ("開始日期:","var_start_date","2026/04/16"),
            ("開始時間:","var_start_time","00:00:00"),
            ("結束日期:","var_end_date",  "2026/04/23"),
            ("結束時間:","var_end_time",  "23:59:59"),
        ]:
            row = ttk.Frame(g); row.pack(fill=tk.X, pady=1)
            ttk.Label(row, text=lbl, width=9).pack(side=tk.LEFT)
            var = tk.StringVar(value=val); setattr(self, attr, var)
            ttk.Entry(row, textvariable=var, width=14).pack(side=tk.RIGHT)
        ttk.Label(g, text="格式  YYYY/MM/DD  HH:MM:SS",
                  font=("",7), foreground="gray").pack(anchor=tk.W)

        # 頻率範圍
        g = ttk.LabelFrame(parent, text="頻率範圍", padding=5)
        g.pack(fill=tk.X, **pad)
        for lbl, attr, val in [("最低頻率:","var_freq_low","1 Hz"),
                                ("最高頻率:","var_freq_high","100 Hz")]:
            row = ttk.Frame(g); row.pack(fill=tk.X, pady=1)
            ttk.Label(row, text=lbl, width=9).pack(side=tk.LEFT)
            var = tk.StringVar(value=val); setattr(self, attr, var)
            ttk.Combobox(row, textvariable=var, values=FREQ_NAMES,
                         width=10, state="readonly").pack(side=tk.RIGHT)
        # 目前選擇的物理量單位（動態更新）
        self.var_unit_hint = tk.StringVar(value="物理量: m/s²（SI 加速度基本單位）")
        ttk.Label(g, textvariable=self.var_unit_hint,
                  font=("",7), foreground="#2c5282",
                  wraplength=295).pack(anchor=tk.W, pady=(3,0))

        # 振動量單位（Tree-view 單位選擇器）
        g = ttk.LabelFrame(parent, text="振動量單位", padding=5)
        g.pack(fill=tk.X, **pad)
        self._build_unit_tree(g)

        # 分析間距
        g = ttk.LabelFrame(parent, text="分析間距", padding=5)
        g.pack(fill=tk.X, **pad)
        self.var_interval = tk.StringVar(value="1 秒")
        cols_f = ttk.Frame(g); cols_f.pack(fill=tk.X)
        for lbl, _, _ in INTERVAL_OPTIONS:
            ttk.Radiobutton(cols_f, text=lbl, variable=self.var_interval,
                            value=lbl).pack(anchor=tk.W, pady=1)
        ttk.Label(g, text="原始資料 0.1s/筆；間距決定 RMS 聚合粒度",
                  font=("",7), foreground="gray").pack(anchor=tk.W)

        # 分析按鈕
        ttk.Separator(parent, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=6, pady=5)
        self.btn_analyze = ttk.Button(parent, text="▶  開始分析",
                                      command=self._start_analysis)
        self.btn_analyze.pack(fill=tk.X, padx=8)
        self.var_status = tk.StringVar(value="就緒")
        ttk.Label(parent, textvariable=self.var_status, wraplength=310,
                  foreground="gray", font=("",8)).pack(padx=6, pady=(4,1), anchor=tk.W)
        self.progress = ttk.Progressbar(parent, mode="indeterminate", length=310)
        self.progress.pack(padx=6)

    # ─────────────── 振動量單位下拉選單 ─────────────────

    def _build_unit_tree(self, parent):
        UNIT_OPTIONS = [
            ("m/s²",   "加速度 m/s² — SI 基本單位"),
            ("g",      "加速度 g — 9.80665 m/s²"),
            ("Gal",    "加速度 Gal — 0.01 m/s²"),
            ("dB_acc", "加速度 dB — ref 10⁻⁶ m/s²"),
            ("m/s",    "速度 m/s — SI 基本單位"),
            ("mm/s",   "速度 mm/s — ×10³ m/s"),
            ("ips",    "速度 ips — 英制 in/s"),
            ("dB_vel", "速度 dB — ref 2.54×10⁻⁸ m/s"),
        ]
        self._unit_options_map = {text: key for key, text in UNIT_OPTIONS}
        display_values = [text for _, text in UNIT_OPTIONS]

        self.var_unit_combo = tk.StringVar(value=display_values[0])
        self._unit_combo = ttk.Combobox(
            parent, textvariable=self.var_unit_combo,
            values=display_values, state="readonly", width=35)
        self._unit_combo.pack(fill=tk.X)
        self._unit_combo.bind("<<ComboboxSelected>>", self._on_unit_select)

        # dB 基準值編輯區（預設隱藏）
        self._db_frame = ttk.Frame(parent)
        ttk.Label(self._db_frame, text="dB 基準值:", font=("",8)).pack(side=tk.LEFT)
        self.var_db_ref = tk.StringVar(value="1e-6")
        ttk.Entry(self._db_frame, textvariable=self.var_db_ref,
                  width=10).pack(side=tk.LEFT, padx=3)
        self.lbl_db_unit = ttk.Label(self._db_frame, text="m/s²", font=("",8,"italic"))
        self.lbl_db_unit.pack(side=tk.LEFT)

    def _get_unit_key(self) -> str:
        text = self.var_unit_combo.get()
        return self._unit_options_map.get(text, "m/s²")

    def _on_unit_select(self, _event=None):
        iid = self._get_unit_key()
        if iid not in UNIT_DEFS:
            return
        cat, lbl, desc, _, db_ref = UNIT_DEFS[iid]
        self.var_unit_hint.set(f"物理量: {lbl}（{desc}）")
        if db_ref is not None:
            self.var_db_ref.set(str(db_ref))
            base_unit = "m/s²" if cat == "acc" else "m/s"
            self.lbl_db_unit.config(text=base_unit)
            self._db_frame.pack(fill=tk.X, pady=(3,0))
        else:
            self._db_frame.pack_forget()

    # ─────────────── 右側結果面板 ────────────────────────

    def _build_results(self, parent):
        self.nb = ttk.Notebook(parent)
        self.nb.pack(fill=tk.BOTH, expand=True)

        for text, builder in [
            ("  原始數據  ",  self._build_raw_tab),
            ("  時間序列  ",  self._build_time_tab),
            ("  頻率分析  ",  self._build_freq_tab),
            ("  統計表    ",  self._build_stats_tab),
            ("  VC 曲線   ",  self._build_vc_tab),
            ("  環境振動  ",  self._build_env_tab),
        ]:
            t = ttk.Frame(self.nb)
            self.nb.add(t, text=text)
            builder(t)

    # ─────────────── 原始數據分頁 ────────────────────────

    def _build_raw_tab(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(bar, text="顯示軸別:").pack(side=tk.LEFT)
        self.var_raw_axis = tk.StringVar(value="全部")
        for opt in ["全部","X 軸","Y 軸","Z 軸"]:
            ttk.Radiobutton(bar, text=opt, variable=self.var_raw_axis,
                            value=opt, command=self._refresh_raw).pack(side=tk.LEFT, padx=2)
        self.var_raw_bb = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="含寬頻欄位", variable=self.var_raw_bb,
                        command=self._refresh_raw).pack(side=tk.LEFT, padx=(10,2))
        ttk.Button(bar, text="匯出原始數據 CSV",
                   command=self._export_raw).pack(side=tk.RIGHT, padx=4)

        pbar = ttk.Frame(parent)
        pbar.pack(fill=tk.X, padx=6, pady=(0,2))
        ttk.Button(pbar, text="◀◀", width=3,
                   command=lambda: self._raw_goto(0)).pack(side=tk.LEFT)
        ttk.Button(pbar, text="◀",  width=3,
                   command=lambda: self._raw_goto(self._raw_page-1)).pack(side=tk.LEFT, padx=(2,0))
        self.lbl_raw_page = ttk.Label(pbar, text="尚無資料", width=46, anchor=tk.CENTER)
        self.lbl_raw_page.pack(side=tk.LEFT, padx=6)
        ttk.Button(pbar, text="▶",  width=3,
                   command=lambda: self._raw_goto(self._raw_page+1)).pack(side=tk.LEFT)
        ttk.Button(pbar, text="▶▶", width=3,
                   command=lambda: self._raw_goto(self._raw_total_pages-1)).pack(side=tk.LEFT, padx=(2,0))

        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0,6))
        self.raw_tree = ttk.Treeview(frame, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL,   command=self.raw_tree.yview)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.raw_tree.xview)
        self.raw_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.raw_tree.pack(fill=tk.BOTH, expand=True)
        self.raw_tree.tag_configure("odd",  background="#f8f8f8")
        self.raw_tree.tag_configure("even", background="white")

    # ─────────────── 時間序列分頁 ────────────────────────

    def _build_time_tab(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(bar, text="顯示軸別:").pack(side=tk.LEFT)
        self.show_x   = tk.BooleanVar(value=True)
        self.show_y   = tk.BooleanVar(value=True)
        self.show_z   = tk.BooleanVar(value=True)
        self.show_xyz = tk.BooleanVar(value=True)
        for text, var in [("X 軸",self.show_x),("Y 軸",self.show_y),
                          ("Z 軸",self.show_z),("XYZ 合成",self.show_xyz)]:
            ttk.Checkbutton(bar, text=text, variable=var,
                            command=self._refresh_time).pack(side=tk.LEFT, padx=2)
        ttk.Label(bar, text="  顯示抽樣:").pack(side=tk.LEFT, padx=(8,0))
        self.var_resample = tk.StringVar(value="與分析間距相同")
        ttk.Combobox(bar, textvariable=self.var_resample,
                     values=["與分析間距相同","1s","10s","1min","10min","30min"],
                     width=14, state="readonly").pack(side=tk.LEFT)
        ttk.Button(bar, text="更新", command=self._refresh_time).pack(side=tk.LEFT, padx=3)

        self.fig_time = Figure(figsize=(11,5), dpi=96, tight_layout=True)
        self.ax_time  = self.fig_time.add_subplot(111)
        self.canvas_time = FigureCanvasTkAgg(self.fig_time, parent)
        NavigationToolbar2Tk(self.canvas_time, parent)
        self.canvas_time.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    # ─────────────── 頻率分析分頁 ────────────────────────

    def _build_freq_tab(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(bar, text="統計量:").pack(side=tk.LEFT)
        self.var_freq_stat = tk.StringVar(value="Leq")
        for stat in STATS:
            ttk.Radiobutton(bar, text=stat, variable=self.var_freq_stat,
                            value=stat, command=self._refresh_freq).pack(side=tk.LEFT, padx=2)
        ttk.Separator(bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=6, pady=2)
        ttk.Label(bar, text="軸別:").pack(side=tk.LEFT)
        self.show_fx   = tk.BooleanVar(value=True)
        self.show_fy   = tk.BooleanVar(value=True)
        self.show_fz   = tk.BooleanVar(value=True)
        self.show_fxyz = tk.BooleanVar(value=True)
        for text, var in [("X",self.show_fx),("Y",self.show_fy),
                          ("Z",self.show_fz),("XYZ",self.show_fxyz)]:
            ttk.Checkbutton(bar, text=text, variable=var,
                            command=self._refresh_freq).pack(side=tk.LEFT, padx=2)

        self.fig_freq = Figure(figsize=(11,5), dpi=96, tight_layout=True)
        self.ax_freq  = self.fig_freq.add_subplot(111)
        self.canvas_freq = FigureCanvasTkAgg(self.fig_freq, parent)
        NavigationToolbar2Tk(self.canvas_freq, parent)
        self.canvas_freq.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    # ─────────────── 統計表分頁（含時間區段）────────────

    def _build_stats_tab(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(bar, text="顯示軸別:").pack(side=tk.LEFT)
        self.var_stats_axis = tk.StringVar(value="全部")
        for opt in ["全部","X 軸","Y 軸","Z 軸","XYZ 合成"]:
            ttk.Radiobutton(bar, text=opt, variable=self.var_stats_axis,
                            value=opt, command=self._refresh_stats).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="匯出 CSV",
                   command=self._export_stats).pack(side=tk.RIGHT, padx=4)

        # 統計量定義說明列
        _stat_def = (
            "統計量定義（量測期間中有 n% 的數據大於或等於 Ln 值）："
            "Leq = 能量均方根；"
            "L5 > L10 > L50 > L90 > L95 > Lmax"
        )
        ttk.Label(parent, text=_stat_def,
                  font=("Microsoft JhengHei", 7), foreground="#5a6a7a",
                  wraplength=1100, justify=tk.LEFT
                  ).pack(anchor=tk.W, padx=8, pady=(0,1))

        # 時間區段說明列
        self.lbl_stats_period = ttk.Label(parent, text="", foreground="#2c5282",
                                           font=("",8))
        self.lbl_stats_period.pack(anchor=tk.W, padx=8, pady=(0,2))

        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0,6))
        cols = ["軸別","時間區段","頻段","Leq","L5","L10","L50","L90","L95","Lmax"]
        # 欄位標題顯示文字（含統計量定義）
        _col_headings = {
            "軸別":   "軸別",
            "時間區段": "時間區段",
            "頻段":   "頻段",
            "Leq":  "Leq\n(均方根)",
            "L5":   "L5\n(超過 5%)",
            "L10":  "L10\n(超過 10%)",
            "L50":  "L50\n(超過 50%)",
            "L90":  "L90\n(超過 90%)",
            "L95":  "L95\n(超過 95%)",
            "Lmax": "Lmax\n(最大值)",
        }
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        widths = {"軸別":70,"時間區段":200,"頻段":90}
        for col in cols:
            w = widths.get(col, 92)
            self.tree.heading(col, text=_col_headings.get(col, col), anchor=tk.CENTER)
            self.tree.column(col, width=w, anchor=tk.CENTER, minwidth=55)
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL,   command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree.tag_configure("section",   background="#c8d8f0",
                                font=("Microsoft JhengHei",9,"bold"))
        self.tree.tag_configure("broadband", background="#edf2ff")
        self.tree.tag_configure("odd",  background="#f8f8f8")
        self.tree.tag_configure("even", background="white")

    # ─────────────── VC 曲線比較分頁 ─────────────────────

    def _build_vc_tab(self, parent):
        toolbar = ttk.Frame(parent)
        toolbar.pack(fill=tk.X, padx=6, pady=(4, 0))

        # ── 第一列：量測軸別 ──────────────────────────────
        row1 = ttk.Frame(toolbar)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="量測軸別:", width=9, anchor=tk.W).pack(side=tk.LEFT)
        self.vc_show_x   = tk.BooleanVar(value=False)
        self.vc_show_y   = tk.BooleanVar(value=False)
        self.vc_show_z   = tk.BooleanVar(value=False)
        self.vc_show_xyz = tk.BooleanVar(value=True)
        for text, var in [("X", self.vc_show_x), ("Y", self.vc_show_y),
                          ("Z", self.vc_show_z), ("XYZ 合成", self.vc_show_xyz)]:
            ttk.Checkbutton(row1, text=text, variable=var,
                            command=self._refresh_vc).pack(side=tk.LEFT, padx=4)

        ttk.Separator(toolbar, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=1)

        # ── 第二列：統計量（多選）─────────────────────────
        # 統計量定義說明（與統計表一致）
        _stat_hint = (
            "Ln：量測期間有 n% 的數據 ≥ 此值，"
            "故 L5 > L10 > L50 > L90 > L95"
        )
        row2 = ttk.Frame(toolbar)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="統計量:", width=9, anchor=tk.W).pack(side=tk.LEFT)
        self._vc_stat_vars: dict[str, tk.BooleanVar] = {}
        for s in STATS:
            var = tk.BooleanVar(value=(s == "Leq"))
            self._vc_stat_vars[s] = var
            ttk.Checkbutton(row2, text=s, variable=var,
                            command=self._refresh_vc).pack(side=tk.LEFT, padx=4)
        ttk.Label(row2, text=_stat_hint,
                  font=("", 7), foreground="#5a6a7a").pack(side=tk.LEFT, padx=8)

        ttk.Separator(toolbar, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=1)

        # ── 第三列：Y 軸單位 ─────────────────────────────
        row_yu = ttk.Frame(toolbar)
        row_yu.pack(fill=tk.X, pady=2)
        ttk.Label(row_yu, text="Y 軸單位:", width=9, anchor=tk.W).pack(side=tk.LEFT)
        self.var_vc_yunit = tk.StringVar(value="μm/s")
        _vc_yunit_opts = [
            ("μm/s",   "μm/s"),
            ("μin/s",  "μin/s"),
            ("dB_vel", "速度位準 (dB, ref:2.54×10⁻⁸ m/s)"),
            ("dB_acc", "加速度位準 (dB, ref:10⁻⁶ m/s²)"),
        ]
        for val, txt in _vc_yunit_opts:
            ttk.Radiobutton(row_yu, text=txt, variable=self.var_vc_yunit,
                            value=val, command=self._refresh_vc).pack(side=tk.LEFT, padx=4)

        ttk.Separator(toolbar, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=1)

        # ── 第四列：VC 準則曲線選擇 ──────────────────────
        row3 = ttk.Frame(toolbar)
        row3.pack(fill=tk.X, pady=2)
        ttk.Label(row3, text="VC 曲線:", width=9, anchor=tk.W).pack(side=tk.LEFT)
        self._vc_vars: dict[str, tk.BooleanVar] = {}
        for name in VC_CURVES:
            var = tk.BooleanVar(value=(name in ("VC-A", "VC-B", "VC-C")))
            self._vc_vars[name] = var
            ttk.Checkbutton(row3, text=name, variable=var,
                            command=self._refresh_vc).pack(side=tk.LEFT, padx=2)

        # 說明
        note = ttk.Frame(parent)
        note.pack(fill=tk.X, padx=6)
        self.lbl_vc_note = ttk.Label(
            note,
            text="依據 Ungar & Gordon（1991）振動準則（Vibration Criterion），"
                 "以速度 RMS 呈現各 1/3 八度音帶。"
                 "dB 基準值：1×10⁻⁶ in/s（2.54×10⁻⁸ m/s）。"
                 "加速度模式下自動換算 v = a / (2π·f)。",
            font=("",8), foreground="gray")
        self.lbl_vc_note.pack(anchor=tk.W)

        # VC 等級評定結果
        frm_result = ttk.LabelFrame(
            parent,
            text="VC 等級評定結果　（評定頻段：1–80 Hz，Ungar & Gordon 1991）")
        frm_result.pack(fill=tk.X, padx=6, pady=(2,3))
        self.lbl_vc_result = ttk.Label(
            frm_result, text="請先執行分析",
            font=("Microsoft JhengHei", 9), justify=tk.LEFT, foreground="#555555")
        self.lbl_vc_result.pack(padx=10, pady=5, anchor=tk.W)

        self.fig_vc = Figure(figsize=(11,5.0), dpi=96, tight_layout=True)
        self.ax_vc  = self.fig_vc.add_subplot(111)
        self.canvas_vc = FigureCanvasTkAgg(self.fig_vc, parent)
        NavigationToolbar2Tk(self.canvas_vc, parent)
        self.canvas_vc.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    # ─────────────── 環境振動分頁 ──────────────────────────

    def _build_env_tab(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(bar, text="顯示軸別:").pack(side=tk.LEFT)
        self.var_env_axis = tk.StringVar(value="XYZ 合成")
        for opt in ["X 軸", "Y 軸", "Z 軸", "XYZ 合成"]:
            ttk.Radiobutton(bar, text=opt, variable=self.var_env_axis,
                            value=opt, command=self._refresh_env).pack(side=tk.LEFT, padx=2)
        ttk.Button(bar, text="匯出 CSV",
                   command=self._export_env).pack(side=tk.RIGHT, padx=4)

        self.lbl_env_period = ttk.Label(parent, text="", foreground="#2c5282",
                                         font=("", 8))
        self.lbl_env_period.pack(anchor=tk.W, padx=8, pady=(0, 2))

        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        env_cols = ["時間", "Lveq", "Lvmax", "Lv5", "Lv10", "Lv50", "Lv90", "Lv95"]
        self.env_tree = ttk.Treeview(frame, columns=env_cols, show="headings",
                                      selectmode="browse")
        _env_widths = {"時間": 110}
        for col in env_cols:
            w = _env_widths.get(col, 90)
            self.env_tree.heading(col, text=col, anchor=tk.CENTER)
            self.env_tree.column(col, width=w, anchor=tk.CENTER, minwidth=60)
        vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.env_tree.yview)
        hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.env_tree.xview)
        self.env_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.env_tree.pack(fill=tk.BOTH, expand=True)
        self.env_tree.tag_configure("header",
            background="#c8d8f0", font=("Microsoft JhengHei", 9, "bold"))
        self.env_tree.tag_configure("summary_header",
            background="#a8c8e8", font=("Microsoft JhengHei", 14, "bold"))
        self.env_tree.tag_configure("summary",
            background="#dce8f5", font=("Microsoft JhengHei", 14))
        self.env_tree.tag_configure("odd",  background="#f8f8f8")
        self.env_tree.tag_configure("even", background="white")

    # ════════════════════════════════════════════════════════
    #  事件處理 / 參數解析
    # ════════════════════════════════════════════════════════

    def _browse_dir(self):
        d = filedialog.askdirectory(initialdir=self.var_dir.get())
        if d:
            self.var_dir.set(d)

    def _parse_params(self) -> dict:
        root_dir = self.var_dir.get().strip()
        if not os.path.isdir(root_dir):
            raise ValueError(f"目錄不存在:\n{root_dir}")
        try:
            start_dt = datetime.datetime.strptime(
                f"{self.var_start_date.get().strip()} {self.var_start_time.get().strip()}",
                "%Y/%m/%d %H:%M:%S")
            end_dt = datetime.datetime.strptime(
                f"{self.var_end_date.get().strip()} {self.var_end_time.get().strip()}",
                "%Y/%m/%d %H:%M:%S")
        except ValueError as e:
            raise ValueError(f"日期/時間格式錯誤:\n{e}")
        if start_dt >= end_dt:
            raise ValueError("開始時間必須早於結束時間")

        lo_idx = FREQ_NAMES.index(self.var_freq_low.get())
        hi_idx = FREQ_NAMES.index(self.var_freq_high.get())
        if lo_idx > hi_idx:
            raise ValueError("最低頻率不可高於最高頻率")
        sel_bands = FREQ_BANDS[lo_idx: hi_idx+1]

        # 單位
        unit_key = self._get_unit_key()
        cat, lbl, desc, scale, db_ref_default = UNIT_DEFS[unit_key]

        db_ref = db_ref_default
        if db_ref_default is not None:
            try:
                db_ref = float(self.var_db_ref.get())
            except ValueError:
                db_ref = db_ref_default

        # 分析間距
        interval_label = self.var_interval.get()
        _, interval_sec, interval_rule = next(
            e for e in INTERVAL_OPTIONS if e[0] == interval_label)

        return {
            "root_dir":       root_dir,
            "start_dt":       start_dt,
            "end_dt":         end_dt,
            "sel_bands":      sel_bands,
            "unit_key":       unit_key,
            "unit_cat":       cat,
            "unit_label":     lbl,
            "unit_desc":      desc,
            "unit_scale":     scale,
            "db_ref":         db_ref,
            "interval_label": interval_label,
            "interval_sec":   interval_sec,
            "interval_rule":  interval_rule,
        }

    def _start_analysis(self):
        try:
            params = self._parse_params()
        except ValueError as e:
            messagebox.showerror("參數錯誤", str(e))
            return
        self.btn_analyze.config(state=tk.DISABLED)
        self.progress.start()
        self.var_status.set("正在分析中...")
        threading.Thread(target=self._run_analysis, args=(params,), daemon=True).start()

    # ════════════════════════════════════════════════════════
    #  分析執行緒
    # ════════════════════════════════════════════════════════

    def _run_analysis(self, params: dict):
        try:
            self._post("status", "搜尋 .rnd 資料檔...")
            rnd_files = self._find_rnd_files(params["root_dir"])
            if not rnd_files:
                self._post("error", "找不到任何 .rnd 資料檔案"); return

            sel_bands   = params["sel_bands"]
            needed_cols = [f"{ax}_{n}" for _,n in sel_bands for ax in AXES]

            self._post("status", f"找到 {len(rnd_files)} 個資料檔，平行讀取...")
            all_dfs, completed, total = [], 0, len(rnd_files)

            with ThreadPoolExecutor(max_workers=min(6,total)) as ex:
                futures = {ex.submit(read_rnd_file, f, needed_cols): f for f in rnd_files}
                for fut in as_completed(futures):
                    df = fut.result(); completed += 1
                    if df is not None:
                        mask = ((df["Start Time"] >= params["start_dt"]) &
                                (df["Start Time"] <= params["end_dt"]))
                        df = df.loc[mask]
                        if len(df) > 0: all_dfs.append(df)
                    if completed % 20 == 0 or completed == total:
                        self._post("status", f"讀取 {completed}/{total} 檔（{len(all_dfs)} 有效）")

            if not all_dfs:
                self._post("error", "指定時間範圍內無資料"); return

            self._post("status", "合併排序...")
            df = pd.concat(all_dfs, ignore_index=True)
            df.sort_values("Start Time", inplace=True)
            df.reset_index(drop=True, inplace=True)
            del all_dfs
            n_rows = len(df)

            # 速度轉換（只有速度單位類才做）
            if params["unit_cat"] == "vel":
                self._post("status", "套用速度轉換 v = a / (2π·f)...")
                for f_val, f_name in sel_bands:
                    for ax in AXES:
                        col = f"{ax}_{f_name}"
                        if col in df.columns:
                            df[col] = df[col] / (2.0 * np.pi * f_val)

            # 分析間距 RMS 聚合
            if params["interval_rule"] is not None:
                self._post("status", f"依 {params['interval_label']} 做 RMS 聚合...")
                df = self._aggregate_data(df, params)

            n_agg = len(df)
            self._post("status", "計算統計量...")
            results = self._compute_stats(df, params)
            self._post("done", (df, results, params, n_rows, n_agg))

        except Exception as e:
            import traceback
            self._post("error", f"{e}\n\n{traceback.format_exc()}")

    def _find_rnd_files(self, root_dir: str) -> list[str]:
        files = []
        try:
            for e in sorted(os.scandir(root_dir), key=lambda x: x.name):
                if e.is_dir() and not e.name.startswith("."):
                    files.extend(sorted(glob.glob(os.path.join(e.path,"*.rnd"))))
        except Exception: pass
        files.extend(sorted(glob.glob(os.path.join(root_dir,"*.rnd"))))
        return files

    def _aggregate_data(self, df: pd.DataFrame, params: dict) -> pd.DataFrame:
        sel_bands = params["sel_bands"]
        data_cols = [f"{ax}_{n}" for _,n in sel_bands for ax in AXES if f"{ax}_{n}" in df.columns]
        def rms(x):
            a = np.asarray(x, dtype=np.float64)
            return float(np.sqrt(np.mean(a**2))) if len(a) > 0 else np.nan
        agg = df.set_index("Start Time")[data_cols].resample(
            params["interval_rule"]).apply(rms)
        agg.dropna(how="all", inplace=True)
        agg.reset_index(inplace=True)
        return agg

    def _compute_stats(self, df: pd.DataFrame, params: dict) -> dict:
        sel_bands = params["sel_bands"]
        results = {}
        for axis in AXES + ["XYZ"]:
            band_arrs: dict[str, np.ndarray] = {}
            for _, f_name in sel_bands:
                if axis == "XYZ":
                    cols = [f"{ax}_{f_name}" for ax in AXES if f"{ax}_{f_name}" in df.columns]
                    if cols:
                        mat = np.nan_to_num(df[cols].values.astype(np.float64))
                        band_arrs[f_name] = np.sqrt(np.sum(mat**2, axis=1))
                else:
                    col = f"{axis}_{f_name}"
                    if col in df.columns:
                        band_arrs[f_name] = np.nan_to_num(df[col].values.astype(np.float64))

            if band_arrs:
                sq = np.zeros(len(df), dtype=np.float64)
                for a in band_arrs.values(): sq += a**2
                bb = np.sqrt(sq)
                results[axis] = {
                    "bands":      {n: calc_stats(a) for n,a in band_arrs.items()},
                    "broadband":  calc_stats(bb),
                    "timeseries": pd.Series(bb.astype(np.float32),
                                            index=pd.DatetimeIndex(df["Start Time"].values)),
                }
            else:
                results[axis] = {"bands": {}, "broadband": {s: np.nan for s in STATS},
                                  "timeseries": None}
        return results

    # ════════════════════════════════════════════════════════
    #  佇列輪詢
    # ════════════════════════════════════════════════════════

    def _post(self, t, p): self._queue.put((t, p))

    def _poll_queue(self):
        try:
            while True:
                msg_type, payload = self._queue.get_nowait()
                if msg_type == "status":
                    self.var_status.set(payload)
                elif msg_type == "error":
                    self.progress.stop()
                    self.btn_analyze.config(state=tk.NORMAL)
                    self.var_status.set("分析失敗")
                    messagebox.showerror("錯誤", payload)
                elif msg_type == "done":
                    df, results, params, n_rows, n_agg = payload
                    self._df = df; self._results = results; self._params = params
                    self.progress.stop()
                    self.btn_analyze.config(state=tk.NORMAL)
                    self.var_status.set(
                        f"完成！原始 {n_rows:,} 筆 → 間距 {params['interval_label']}："
                        f"{n_agg:,} 個時段 ／ {len(params['sel_bands'])} 個頻段")
                    self._refresh_raw()
                    self._refresh_time()
                    self._refresh_freq()
                    self._refresh_stats()
                    self._refresh_vc()
                    self._refresh_env()
                    self.nb.select(0)
        except queue.Empty: pass
        self.root.after(100, self._poll_queue)

    # ════════════════════════════════════════════════════════
    #  單位換算輔助
    # ════════════════════════════════════════════════════════

    def _unit_label(self) -> str:
        if not self._params: return "m/s²"
        return self._params["unit_label"]

    def _to_display(self, val: float) -> float:
        """將基底單位值換算為顯示單位（dB 或線性）。"""
        if not self._params or not np.isfinite(val): return val
        _, _, _, scale, db_ref_default = UNIT_DEFS[self._params["unit_key"]]
        db_ref = self._params["db_ref"]
        if db_ref_default is not None:
            return 20.0 * np.log10(max(val, 1e-300) / db_ref) if db_ref else val
        return val * scale

    def _fmt(self, v) -> str:
        if v is None or (isinstance(v, float) and np.isnan(v)): return "—"
        if abs(v) < 1e-3 and v != 0: return f"{v:.3e}"
        return f"{v:.6f}"

    def _fmt_d(self, v) -> str:
        """換算後格式化；dB 單位取一位小數。"""
        dv = self._to_display(v)
        if dv is None or (isinstance(dv, float) and np.isnan(dv)): return "—"
        if self._params:
            _, _, _, _, db_ref_default = UNIT_DEFS[self._params["unit_key"]]
            if db_ref_default is not None:
                return f"{dv:.1f}"
        return self._fmt(dv)

    # ════════════════════════════════════════════════════════
    #  刷新：原始數據
    # ════════════════════════════════════════════════════════

    def _prepare_raw_df(self) -> pd.DataFrame | None:
        if self._df is None or self._params is None: return None
        df = self._df.copy()
        for ax in AXES:
            cols = [f"{ax}_{n}" for _,n in self._params["sel_bands"]
                    if f"{ax}_{n}" in df.columns]
            if cols:
                mat = np.nan_to_num(df[cols].values.astype(np.float64))
                df[f"_BB_{ax}"] = np.sqrt(np.sum(mat**2, axis=1))
        return df

    def _refresh_raw(self):
        if self._df is None or self._params is None: return
        self._raw_display_df  = self._prepare_raw_df()
        total                 = len(self._raw_display_df) if self._raw_display_df is not None else 0
        self._raw_page        = 0
        self._raw_total_pages = max(1, -(-total // self.RAW_PAGE_SIZE))
        self._raw_setup_columns()
        self._raw_show_page(0)

    def _raw_setup_columns(self):
        if not self._params: return
        sel_bands   = self._params["sel_bands"]
        axis_filter = self.var_raw_axis.get()
        show_bb     = self.var_raw_bb.get()
        unit        = self._unit_label()
        show_axes   = AXES if axis_filter == "全部" else [axis_filter[0]]

        cols = ["編號", "時間"]
        if show_bb:
            for ax in show_axes: cols.append(f"{ax} 寬頻({unit})")
        for _, f_name in sel_bands:
            for ax in show_axes: cols.append(f"{ax}_{f_name}")

        self.raw_tree["columns"] = cols
        self.raw_tree.column("編號", width=60, anchor=tk.CENTER, minwidth=50)
        self.raw_tree.heading("編號", text="編號")
        self.raw_tree.column("時間", width=160, anchor=tk.CENTER, minwidth=140)
        self.raw_tree.heading("時間",
            text=f"時間（間距: {self._params['interval_label']}）")
        for col in cols[2:]:
            self.raw_tree.column(col, width=90, anchor=tk.CENTER, minwidth=65)
            self.raw_tree.heading(col, text=col)
        self._raw_show_axes = show_axes
        self._raw_show_bb   = show_bb

    def _raw_show_page(self, page: int):
        if self._raw_display_df is None: return
        page = max(0, min(page, self._raw_total_pages-1))
        self._raw_page = page
        total = len(self._raw_display_df)
        start = page * self.RAW_PAGE_SIZE
        end   = min(start + self.RAW_PAGE_SIZE, total)
        chunk = self._raw_display_df.iloc[start:end]
        self.lbl_raw_page.config(
            text=f"第 {page+1}/{self._raw_total_pages} 頁"
                 f"（每頁 {self.RAW_PAGE_SIZE} 筆，共 {total:,} 筆）")
        self.raw_tree.delete(*self.raw_tree.get_children())
        sel_bands = self._params["sel_bands"]

        for i, (_, row) in enumerate(chunk.iterrows()):
            ts = row.get("Start Time")
            global_no = start + i + 1
            vals = [str(global_no), ts.strftime("%Y/%m/%d %H:%M:%S") if pd.notna(ts) else ""]
            if self._raw_show_bb:
                for ax in self._raw_show_axes:
                    vals.append(self._fmt_d(row.get(f"_BB_{ax}", np.nan)))
            for _, f_name in sel_bands:
                for ax in self._raw_show_axes:
                    vals.append(self._fmt_d(row.get(f"{ax}_{f_name}", np.nan)))
            self.raw_tree.insert("","end", values=vals,
                                 tags=("odd" if i%2==0 else "even",))

    def _raw_goto(self, page: int):
        if self._raw_display_df is not None: self._raw_show_page(page)

    def _export_raw(self):
        if self._raw_display_df is None:
            messagebox.showinfo("提示","請先執行分析"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV","*.csv")],
            initialfile="raw_data.csv")
        if not path: return
        df  = self._raw_display_df.copy()
        unit = self._unit_label()
        sel_bands = self._params["sel_bands"]
        # 轉換值並重命名
        for col in df.columns:
            if col == "Start Time" or col.startswith("_BB_"): continue
            df[col] = df[col].apply(lambda v: self._to_display(v) if np.isfinite(v) else v)
        rename = {f"_BB_{ax}": f"{ax}_寬頻({unit})" for ax in AXES}
        df.rename(columns=rename, inplace=True)
        # 輸出欄位
        out_cols = ["Start Time"]
        for ax in AXES:
            c = f"{ax}_寬頻({unit})"
            if c in df.columns: out_cols.append(c)
        for _, f_name in sel_bands:
            for ax in AXES:
                c = f"{ax}_{f_name}"
                if c in df.columns: out_cols.append(c)
        df[[c for c in out_cols if c in df.columns]].to_csv(
            path, index=False, encoding="utf-8-sig")
        messagebox.showinfo("匯出完成", f"已儲存 {len(df):,} 筆至:\n{path}")

    # ════════════════════════════════════════════════════════
    #  刷新：時間序列
    # ════════════════════════════════════════════════════════

    def _refresh_time(self):
        if self._results is None: return
        self.ax_time.clear()
        unit = self._unit_label()
        resample_map = {
            "與分析間距相同": self._params.get("interval_rule") if self._params else None,
            "1s":"1s","10s":"10s","1min":"1min","10min":"10min","30min":"30min",
        }
        rule = resample_map.get(self.var_resample.get())

        cfg = [("X",self.show_x,"X 軸"),("Y",self.show_y,"Y 軸"),
               ("Z",self.show_z,"Z 軸"),("XYZ",self.show_xyz,"XYZ 合成")]
        has = False
        for axis, var, label in cfg:
            if not var.get(): continue
            ts: pd.Series|None = self._results.get(axis,{}).get("timeseries")
            if ts is None: continue
            if rule:
                ts = ts.resample(rule).apply(
                    lambda x: float(np.sqrt(np.mean(np.asarray(x)**2))) if len(x)>0 else np.nan
                ).dropna()
            # 換算顯示單位
            disp = ts.apply(self._to_display)
            self.ax_time.plot(disp.index, disp.values,
                              color=AXIS_COLORS[axis], label=label,
                              linewidth=0.9, alpha=0.85)
            has = True

        if has:
            self.ax_time.set_xlabel("時間")
            self.ax_time.set_ylabel(f"寬頻振動量 ({unit})")
            self.ax_time.set_title("寬頻振動時間序列")
            self.ax_time.legend(loc="upper right", fontsize=9)
            self.ax_time.grid(True, alpha=0.3)
            self.fig_time.autofmt_xdate(rotation=25)
            self.fig_time.tight_layout()
        self.canvas_time.draw()

    # ════════════════════════════════════════════════════════
    #  刷新：頻率分析
    # ════════════════════════════════════════════════════════

    def _refresh_freq(self):
        if self._results is None: return
        self.ax_freq.clear()
        stat = self.var_freq_stat.get()
        sel_bands = self._params["sel_bands"]
        unit = self._unit_label()

        cfg = [("X",self.show_fx,"X 軸"),("Y",self.show_fy,"Y 軸"),
               ("Z",self.show_fz,"Z 軸"),("XYZ",self.show_fxyz,"XYZ 合成")]
        n_vis = sum(v.get() for _,v,_ in cfg)
        if n_vis == 0: self.canvas_freq.draw(); return

        x = np.arange(len(sel_bands))
        bw = min(0.18, 0.72/max(n_vis,1))
        offs = np.linspace(-(n_vis-1)/2, (n_vis-1)/2, n_vis) * bw
        vi = 0
        for axis, var, label in cfg:
            if not var.get(): continue
            bd = self._results.get(axis,{}).get("bands",{})
            vals = [self._to_display(bd.get(n,{}).get(stat, np.nan))
                    for _,n in sel_bands]
            self.ax_freq.bar(x+offs[vi], vals, bw, label=label,
                             color=AXIS_COLORS[axis], alpha=0.78)
            vi += 1

        self.ax_freq.set_xticks(x)
        self.ax_freq.set_xticklabels([n for _,n in sel_bands],
                                      rotation=45, ha="right", fontsize=8)
        self.ax_freq.set_xlabel("1/3 八度音頻段")
        self.ax_freq.set_ylabel(f"{stat} ({unit})")
        self.ax_freq.set_title(f"各頻段 {stat} 比較")
        self.ax_freq.legend(fontsize=9)
        self.ax_freq.grid(True, alpha=0.3, axis="y")
        self.fig_freq.tight_layout()
        self.canvas_freq.draw()

    # ════════════════════════════════════════════════════════
    #  刷新：統計表（含時間區段）
    # ════════════════════════════════════════════════════════

    def _refresh_stats(self):
        if self._results is None: return
        for item in self.tree.get_children(): self.tree.delete(item)

        p     = self._params
        unit  = self._unit_label()
        bands = p["sel_bands"]
        start_str = p["start_dt"].strftime("%Y/%m/%d %H:%M:%S")
        end_str   = p["end_dt"].strftime("%Y/%m/%d %H:%M:%S")
        period_str = f"{start_str} ～ {end_str}"

        self.lbl_stats_period.config(
            text=f"分析期間：{period_str}  ／  間距：{p['interval_label']}"
                 f"  ／  單位：{unit}")

        axis_map = {
            "全部":     [("X","X 軸"),("Y","Y 軸"),("Z","Z 軸"),("XYZ","XYZ 合成")],
            "X 軸":     [("X","X 軸")],
            "Y 軸":     [("Y","Y 軸")],
            "Z 軸":     [("Z","Z 軸")],
            "XYZ 合成": [("XYZ","XYZ 合成")],
        }
        for axis_key, axis_label in axis_map.get(self.var_stats_axis.get(),
                                                  axis_map["全部"]):
            data = self._results.get(axis_key, {})
            self.tree.insert("","end",
                values=(axis_label, period_str, f"單位: {unit}",
                        "Leq","L5","L10","L50","L90","L95","Lmax"),
                tags=("section",))
            bb = data.get("broadband", {})
            if bb:
                self.tree.insert("","end",
                    values=("","", "★ 寬頻合計",
                            self._fmt_d(bb.get("Leq")), self._fmt_d(bb.get("L5")),
                            self._fmt_d(bb.get("L10")), self._fmt_d(bb.get("L50")),
                            self._fmt_d(bb.get("L90")), self._fmt_d(bb.get("L95")),
                            self._fmt_d(bb.get("Lmax"))),
                    tags=("broadband",))
            bd = data.get("bands", {})
            for i, (_, f_name) in enumerate(bands):
                s = bd.get(f_name, {})
                self.tree.insert("","end",
                    values=("","", f_name,
                            self._fmt_d(s.get("Leq")), self._fmt_d(s.get("L5")),
                            self._fmt_d(s.get("L10")), self._fmt_d(s.get("L50")),
                            self._fmt_d(s.get("L90")), self._fmt_d(s.get("L95")),
                            self._fmt_d(s.get("Lmax"))),
                    tags=("odd" if i%2==0 else "even",))

    # ════════════════════════════════════════════════════════
    #  刷新：VC 曲線
    # ════════════════════════════════════════════════════════

    def _refresh_vc(self):
        if self._results is None or self._params is None:
            self.ax_vc.clear()
            self.ax_vc.text(0.5, 0.5, "請先執行分析", ha="center", va="center",
                            transform=self.ax_vc.transAxes, fontsize=14)
            self.canvas_vc.draw()
            self.lbl_vc_result.config(text="請先執行分析", foreground="#555555")
            return

        self.ax_vc.clear()
        p         = self._params
        sel_bands = p["sel_bands"]
        is_vel    = (p["unit_cat"] == "vel")
        _, _, _, scale, db_ref_key = UNIT_DEFS[p["unit_key"]]

        fn_map: dict[float, str] = {fv: fn for fv, fn in sel_bands}

        def _match_bands(target_freqs):
            """從 sel_bands 中比對最接近 target_freqs 的頻段（容差 15%）。"""
            result = []
            for fv in target_freqs:
                closest = min(fn_map.keys(), key=lambda x: abs(x - fv), default=None)
                if closest is not None and abs(closest - fv) / fv < 0.15:
                    result.append((fv, fn_map[closest]))
            return result

        # 量測數據繪圖用頻段：1–100 Hz（顯示完整頻譜）
        plot_bands: list[tuple[float, str]] = _match_bands(VC_MEAS_FREQS)
        # VC 等級評定用頻段：4–80 Hz（依 Ungar & Gordon 標準）
        eval_bands: list[tuple[float, str]] = _match_bands(VC_PLOT_FREQS)

        if not eval_bands:
            self.ax_vc.text(0.5, 0.5, "VC 等級評定需要 4–80 Hz 頻段數據\n請調整頻率範圍",
                            ha="center", va="center", transform=self.ax_vc.transAxes,
                            fontsize=12)
            self.canvas_vc.draw()
            self.lbl_vc_result.config(text="無可用頻段數據（需包含 4–80 Hz）",
                                       foreground="#cc3300")
            return

        vc_yunit = self.var_vc_yunit.get()

        def to_vel_ums(stat_base: float, fv: float) -> float:
            """原始統計量 → μm/s（速度）。"""
            if not np.isfinite(stat_base): return np.nan
            vel_ms = stat_base if is_vel else stat_base / (2.0 * np.pi * fv)
            return vel_ms * 1e6

        def _disp(val_ums: float, freq: float = 0.0) -> float:
            """μm/s → 顯示單位。freq 僅加速度位準 dB 需要。"""
            if not np.isfinite(val_ums) or val_ums <= 0:
                return np.nan
            if vc_yunit == "μin/s":
                return val_ums / 0.0254
            elif vc_yunit == "dB_vel":
                return 20.0 * np.log10(val_ums / VC_DB_REF_UMS)
            elif vc_yunit == "dB_acc":
                return 20.0 * np.log10(val_ums * 2.0 * np.pi * freq)
            return val_ums

        cfg = [("X", self.vc_show_x), ("Y", self.vc_show_y),
               ("Z", self.vc_show_z), ("XYZ", self.vc_show_xyz)]

        sel_stats = [s for s in STATS if self._vc_stat_vars[s].get()]
        if not sel_stats:
            sel_stats = ["Leq"]   # 至少保留一條線

        stats_label = "、".join(sel_stats)
        result_parts = [f"評定頻段：4–80 Hz（Ungar & Gordon, 1991）　統計量：{stats_label}"]
        any_axis = False

        for axis, var in cfg:
            if not var.get(): continue
            any_axis = True
            bd = self._results.get(axis, {}).get("bands", {})
            stat_grade_lines: list[str] = []

            for stat_key in sel_stats:
                # ── 繪圖：1–100 Hz 量測數據 ──
                plot_freqs = [fv for fv, _ in plot_bands]
                plot_vels  = [_disp(to_vel_ums(bd.get(fn, {}).get(stat_key, np.nan), fv), fv)
                              for fv, fn in plot_bands]
                self.ax_vc.plot(
                    plot_freqs, plot_vels,
                    color=AXIS_COLORS[axis],
                    linestyle=STAT_LINESTYLES.get(stat_key, "-"),
                    linewidth=2.0, marker="o", markersize=3, zorder=4,
                    label=f"{AXIS_LABELS[axis]}  {stat_key}")

                # ── 評定：4–80 Hz 逐頻帶對照 VC 曲線 ──
                eval_vels = [to_vel_ums(bd.get(fn, {}).get(stat_key, np.nan), fv)
                             for fv, fn in eval_bands]
                band_vels = [(fv, v) for (fv, _), v in zip(eval_bands, eval_vels)
                             if np.isfinite(v) and v > 0]
                if band_vels:
                    peak  = max(v for _, v in band_vels)
                    grade = _vc_grade(band_vels)
                    limit = VC_CURVES.get(grade)
                    limit_str = f"{limit:g}" if limit else "—"
                    stat_grade_lines.append(
                        f"    {stat_key}：峰值 {peak:.2f} μm/s　→　達到 {grade}"
                        f"（平坦段限值 {limit_str} μm/s）")
                else:
                    stat_grade_lines.append(f"    {stat_key}：無有效數據")

            if stat_grade_lines:
                result_parts.append(f"  【{AXIS_LABELS[axis]}】")
                result_parts.extend(stat_grade_lines)
            else:
                result_parts.append(f"  【{AXIS_LABELS[axis]}】無有效數據")

        if not any_axis:
            result_parts.append("  （未選取任何量測軸別）")

        # ── 繪製 VC 準則曲線（4–80 Hz，斜率段 4–8 Hz + 水平段 8–80 Hz）──
        for idx, (name, flat_limit) in enumerate(VC_CURVES.items()):
            if not self._vc_vars[name].get(): continue
            color = VC_COLORS[idx % len(VC_COLORS)]
            cv    = [_disp(v, f) for v, f in zip(
                        _vc_curve_vals(flat_limit, VC_PLOT_FREQS), VC_PLOT_FREQS)]
            self.ax_vc.plot(
                VC_PLOT_FREQS, cv,
                color=color, linestyle="-", linewidth=1.6, alpha=0.9, zorder=2)
            flat_db = VC_CURVES_DB.get(name, "")
            if vc_yunit == "dB_vel":
                label_txt = f"{name} ({flat_db} dB)"
            elif vc_yunit == "dB_acc":
                acc_db = 20.0 * np.log10(flat_limit * 2.0 * np.pi * 80.0)
                label_txt = f"{name} ({acc_db:.0f} dB)"
            elif vc_yunit == "μin/s":
                label_txt = f"{name} ({_disp(flat_limit, 80):.0f} μin/s)"
            else:
                label_txt = f"{name} ({flat_limit:.1f} μm/s)"
            disp_flat = _disp(flat_limit, 80)
            self.ax_vc.annotate(
                label_txt,
                xy=(80, disp_flat), xycoords="data",
                xytext=(4, 0), textcoords="offset points",
                va="center", fontsize=7, color=color, clip_on=False)

        # x 軸：1–100 Hz（量測數據顯示範圍）
        x_ticks      = VC_MEAS_FREQS
        x_tick_labels = ["1","1.25","1.6","2","2.5","3.15","4","5","6.3","8","10",
                          "12.5","16","20","25","31.5","40","50","63","80","100"]
        self.ax_vc.set_xscale("log")
        if vc_yunit in ("dB_vel", "dB_acc"):
            self.ax_vc.set_yscale("linear")
            if vc_yunit == "dB_acc":
                self.ax_vc.set_ylim(20, 140)
            else:
                self.ax_vc.set_ylim(20, 120)
        else:
            self.ax_vc.set_yscale("log")
            if vc_yunit == "μin/s":
                self.ax_vc.set_ylim(0.5 / 0.0254, 10000 / 0.0254)
            else:
                self.ax_vc.set_ylim(0.5, 10000)
        self.ax_vc.set_xlim(0.8, 115)
        self.ax_vc.set_xlabel("One-Third Octave Band Center Frequency, Hz", fontsize=9)
        _ylabel_map = {
            "μm/s": "rms Velocity, μm/s",
            "μin/s": "rms Velocity, μin/s",
            "dB_vel": "Velocity Level, dB re 10⁻⁶ in/s (2.54×10⁻⁸ m/s)",
            "dB_acc": "Acceleration Level, dB re 10⁻⁶ m/s²",
        }
        self.ax_vc.set_ylabel(_ylabel_map.get(vc_yunit, "rms Velocity"), fontsize=9)
        self.ax_vc.set_title(
            "Generic Vibration Criterion (VC) Curves　Ungar & Gordon (1991)"
            f"\n統計量：{stats_label}",
            fontsize=9)
        handles, labels = self.ax_vc.get_legend_handles_labels()
        if handles:
            self.ax_vc.legend(handles, labels, fontsize=7.5,
                              loc="upper left", framealpha=0.85)
        self.ax_vc.grid(True, which="major", alpha=0.35, linewidth=0.6)
        self.ax_vc.grid(True, which="minor", alpha=0.15, linewidth=0.4)
        self.ax_vc.set_xticks(x_ticks)
        self.ax_vc.set_xticklabels(x_tick_labels, fontsize=7.5)
        self.ax_vc.minorticks_on()
        self.fig_vc.tight_layout(rect=[0, 0, 0.83, 1])
        self.canvas_vc.draw()

        # ── 更新評定結果文字 ──
        self.lbl_vc_result.config(
            text="\n".join(result_parts), foreground="#1a1a2e")

    # ════════════════════════════════════════════════════════
    #  刷新：環境振動
    # ════════════════════════════════════════════════════════

    def _refresh_env(self):
        if self._df is None or self._params is None:
            return
        for item in self.env_tree.get_children():
            self.env_tree.delete(item)

        p = self._params
        axis_map = {"X 軸": "X", "Y 軸": "Y", "Z 軸": "Z", "XYZ 合成": "XYZ"}
        axis_key = axis_map.get(self.var_env_axis.get(), "XYZ")
        unit = self._unit_label()
        df = self._df.copy()

        sel_bands = p["sel_bands"]
        if axis_key == "XYZ":
            cols = []
            for _, fn in sel_bands:
                ax_cols = [f"{ax}_{fn}" for ax in AXES if f"{ax}_{fn}" in df.columns]
                if ax_cols:
                    mat = np.nan_to_num(df[ax_cols].values.astype(np.float64))
                    df[f"_XYZ_{fn}"] = np.sqrt(np.sum(mat ** 2, axis=1))
                    cols.append(f"_XYZ_{fn}")
        else:
            cols = [f"{axis_key}_{fn}" for _, fn in sel_bands
                    if f"{axis_key}_{fn}" in df.columns]

        if not cols:
            return

        mat_all = np.nan_to_num(df[cols].values.astype(np.float64))
        df["_BB"] = np.sqrt(np.sum(mat_all ** 2, axis=1))

        to_dsp = self._to_display
        fmt = self._fmt_d

        start_str = p["start_dt"].strftime("%Y/%m/%d %H:%M")
        end_str = p["end_dt"].strftime("%Y/%m/%d %H:%M")
        self.lbl_env_period.config(
            text=f"分析期間：{start_str} ～ {end_str}　／　"
                 f"間距：{p['interval_label']}　／　"
                 f"單位：{unit}　／　軸別：{self.var_env_axis.get()}")

        ts = pd.DatetimeIndex(df["Start Time"])
        df["_hour"] = ts.hour

        start_hour = p["start_dt"].hour
        end_hour = p["end_dt"].hour
        if end_hour == 0:
            end_hour = 24

        hour_groups = df.groupby(df["Start Time"].dt.hour)

        hourly_rows = []
        h = start_hour
        while h < end_hour:
            hour = h % 24
            grp = hour_groups.get_group(hour) if hour in hour_groups.groups else None
            if grp is not None and len(grp) > 0:
                bb = grp["_BB"].values
                st = calc_stats(bb)
                label = f"{hour:02d}~{(hour + 1) % 24:02d}"
                row_vals = (label, fmt(st["Leq"]), fmt(st["Lmax"]),
                            fmt(st["L5"]), fmt(st["L10"]), fmt(st["L50"]),
                            fmt(st["L90"]), fmt(st["L95"]))
                hourly_rows.append((hour, row_vals, st))
            h += 1

        for i, (hour, row_vals, _) in enumerate(hourly_rows):
            self.env_tree.insert("", "end", values=row_vals,
                                 tags=("odd" if i % 2 == 0 else "even",))

        # ── 監測成果摘要（早/日/晚/夜）──
        _periods = [
            ("L早(05~07)", 5, 7),
            ("L日(07~20)", 7, 20),
            ("L晚(20~22)", 20, 22),
            ("L夜(22~05)", 22, 5),
        ]

        self.env_tree.insert("", "end",
            values=("", "", "", "", "", "", "", ""),
            tags=("header",))
        self.env_tree.insert("", "end",
            values=("監測成果", "", "", "Lv5", "Lv10", "", "", "Lvmax"),
            tags=("summary_header",))

        for label, h_start, h_end in _periods:
            if h_start < h_end:
                mask = (df["_hour"] >= h_start) & (df["_hour"] < h_end)
            else:
                mask = (df["_hour"] >= h_start) | (df["_hour"] < h_end)
            subset = df.loc[mask, "_BB"].values
            if len(subset) == 0:
                self.env_tree.insert("", "end",
                    values=(label, "", "", "—", "—", "", "", "—"),
                    tags=("summary",))
                continue
            st = calc_stats(subset)
            self.env_tree.insert("", "end",
                values=(label, "", "", fmt(st["L5"]), fmt(st["L10"]),
                        "", "", fmt(st["Lmax"])),
                tags=("summary",))

    def _export_env(self):
        if self._df is None or self._params is None:
            messagebox.showinfo("提示", "請先執行分析")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 檔案", "*.csv")],
            title="匯出環境振動統計")
        if not path:
            return
        rows = []
        for item in self.env_tree.get_children():
            rows.append(self.env_tree.item(item, "values"))
        out = pd.DataFrame(rows, columns=[
            "時間", "Lveq", "Lvmax", "Lv5", "Lv10", "Lv50", "Lv90", "Lv95"])
        out.to_csv(path, index=False, encoding="utf-8-sig")
        messagebox.showinfo("完成", f"已匯出：\n{path}")

    # ════════════════════════════════════════════════════════
    #  匯出統計表
    # ════════════════════════════════════════════════════════

    def _export_stats(self):
        if self._results is None:
            messagebox.showinfo("提示","請先執行分析"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV","*.csv")],
            initialfile="vibration_statistics.csv")
        if not path: return
        p = self._params
        unit = self._unit_label()
        period = (f"{p['start_dt'].strftime('%Y/%m/%d %H:%M:%S')}"
                  f" ~ {p['end_dt'].strftime('%Y/%m/%d %H:%M:%S')}")
        rows = []
        for axis in AXES + ["XYZ"]:
            data = self._results.get(axis,{})
            bb = data.get("broadband",{})
            if bb:
                rows.append({"軸別":AXIS_LABELS[axis],"分析期間":period,
                             "頻段":"Broadband","單位":unit,
                             **{s: self._to_display(bb.get(s,np.nan)) for s in STATS}})
            for _,f_name in p["sel_bands"]:
                s = data.get("bands",{}).get(f_name,{})
                rows.append({"軸別":AXIS_LABELS[axis],"分析期間":period,
                             "頻段":f_name,"單位":unit,
                             **{st: self._to_display(s.get(st,np.nan)) for st in STATS}})
        pd.DataFrame(rows, columns=["軸別","分析期間","頻段","單位"]+STATS).to_csv(
            path, index=False, encoding="utf-8-sig")
        messagebox.showinfo("匯出完成", f"已儲存至:\n{path}")


# ════════════════════════════════════════════════════════
#  入口
# ════════════════════════════════════════════════════════

def main():
    root = tk.Tk()
    style = ttk.Style()
    for theme in ("vista","winnative","clam","default"):
        try: style.theme_use(theme); break
        except tk.TclError: continue
    import matplotlib.pyplot as plt
    plt.rcParams["font.family"]       = "Microsoft JhengHei"
    plt.rcParams["axes.unicode_minus"] = False
    VibrationApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
