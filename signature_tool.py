#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
签名小工具
导入 Excel 模板（含姓名列）-> 按姓名批量生成手写风签名图片（随机字体/随机笔迹）-> 导出内嵌签名图片的 Excel 或纯图片包。
依赖: pillow, openpyxl
"""

import os
import sys
import math
import random
import shutil
import threading
import queue
import hashlib
import webbrowser
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageTk
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU

APP_TITLE = "签名小工具"


def _app_dir():
    """程序所在目录（可直接写文件）：打包后是 exe 同级，源码运行是脚本目录"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _res_dir():
    """打包资源目录（只读）：PyInstaller onefile 解压到 _MEIPASS"""
    base = getattr(sys, "_MEIPASS", None)
    return Path(base) if base else Path(__file__).resolve().parent


APP_DIR = _app_dir()
RES_DIR = _res_dir()
BASE_DIR = APP_DIR
FONT_DIR = APP_DIR / "fonts"           # 可写：用户可自行放入字体
BUILTIN_FONT_DIR = RES_DIR / "fonts"   # 内置字体（随 exe 打包，只读）

# ---------------------------------------------------------------- 字体管理

# 系统里可能存在的手写/书法类字体关键字（新版 macOS 已精简，旧系统/Windows 可命中）
SYSTEM_FONT_KEYWORDS = [
    "xingkai", "kaiti", "stkaiti", "hanzipen", "baoli", "weibei", "yuppy",
    "lingwai", "libian", "wawati", "yuanti", "hannotate", "手札", "行楷", "楷体",
]
FONT_EXTS = {".ttf", ".otf", ".ttc"}

# 内置开源字体（OFL 协议，Google Fonts）
BUILTIN_FONTS = {
    "马善政毛笔楷书": "MaShanZheng-Regular.ttf",
    "志莽行书": "ZhiMangXing-Regular.ttf",
    "龙藏行草": "LongCang-Regular.ttf",
    "柳建毛草": "LiuJianMaoCao-Regular.ttf",
    "站酷小薇": "ZCOOLXiaoWei-Regular.ttf",
}

FONT_DOWNLOAD_URLS = [
    "https://ghfast.top/https://raw.githubusercontent.com/google/fonts/main/ofl/{slug}/{file}",
    "https://raw.githubusercontent.com/google/fonts/main/ofl/{slug}/{file}",
]
BUILTIN_FONT_SLUGS = {
    "MaShanZheng-Regular.ttf": "mashanzheng",
    "ZhiMangXing-Regular.ttf": "zhimangxing",
    "LongCang-Regular.ttf": "longcang",
    "LiuJianMaoCao-Regular.ttf": "liujianmaocao",
    "ZCOOLXiaoWei-Regular.ttf": "zcoolxiaowei",
}

_font_cache = {}


def scan_fonts():
    """扫描内置字体目录 + 用户字体目录 + 系统手写体，返回 {显示名: 路径}"""
    fonts = {}
    seen = set()
    name_by_file = {v: k for k, v in BUILTIN_FONTS.items()}
    for d, use_alias in ((BUILTIN_FONT_DIR, True), (FONT_DIR, False)):
        if not d.is_dir():
            continue
        for p in sorted(d.iterdir()):
            if p.suffix.lower() in FONT_EXTS and str(p.resolve()) not in seen:
                seen.add(str(p.resolve()))
                key = name_by_file.get(p.name, p.stem) if use_alias else p.stem
                fonts[key] = str(p)
    sys_dirs = [
        Path("/System/Library/Fonts"), Path("/System/Library/Fonts/Supplemental"),
        Path("/Library/Fonts"), Path.home() / "Library" / "Fonts",
        Path("C:/Windows/Fonts"),
    ]
    for d in sys_dirs:
        if not d.is_dir():
            continue
        for p in d.iterdir():
            if p.suffix.lower() in FONT_EXTS and any(k in p.name.lower() for k in SYSTEM_FONT_KEYWORDS):
                fonts.setdefault(p.stem, str(p))
    return fonts


def load_font(path, size):
    key = (path, size)
    if key not in _font_cache:
        try:
            _font_cache[key] = ImageFont.truetype(path, size)
        except Exception:
            _font_cache[key] = ImageFont.truetype(path, size, index=0)
    return _font_cache[key]

# ---------------------------------------------------------------- 签名渲染

INKS = {
    "纯黑": (18, 18, 20),
    "碳黑": (38, 38, 42),
    "蓝黑": (28, 44, 82),
    "深蓝": (22, 46, 112),
}


def render_signature(name, font_path, width=300, height=120, seed=0,
                     slant=0.6, compact=0.4, bold=True, underline=False,
                     ink_mode="随机"):
    """
    渲染单个姓名签名，返回 RGBA PIL.Image（透明底）。
    所有随机数来自 seed+name，保证同名重渲染结果一致（预览即所得）。
    slant: 0~1 倾斜强度; compact: 0~1 字距紧凑(连笔)程度
    """
    rng = random.Random(f"{seed}|{name}")
    name = name.strip()
    if not name:
        raise ValueError("姓名为空")

    ink = INKS["碳黑"] if ink_mode == "随机" else INKS[ink_mode]
    if ink_mode == "随机":
        ink = rng.choice(list(INKS.values()))

    base_size = int(height * rng.uniform(0.58, 0.68))
    # 大画布留足旋转/斜切余量
    W = int(width * 2.2) + base_size
    H = int(height * 2.2) + base_size
    img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    chars = list(name)
    # 连笔紧凑度：字距 0.90~1.08，compact 越大越紧凑
    spacing_lo = 1.08 - 0.14 * compact
    spacing_hi = 1.16 - 0.12 * compact

    x = base_size * 0.6
    baseline = H * 0.62
    stroke_w = rng.choice([1, 1, 2]) if bold else 0
    for i, ch in enumerate(chars):
        if ch == " ":
            x += base_size * 0.35
            continue
        fs = int(base_size * rng.uniform(0.95, 1.05))
        ff = load_font(font_path, fs)
        dy = rng.uniform(-2.5, 2.5) * (height / 120)
        # 每字微旋转：独立小图绘制后旋转再贴合
        pad = 12 + stroke_w * 2
        cw = int(d.textlength(ch, font=ff)) + pad * 2
        ch_img = Image.new("RGBA", (cw, fs * 2 + pad * 2), (0, 0, 0, 0))
        cd = ImageDraw.Draw(ch_img)
        cd.text((pad, pad), ch, font=ff, fill=ink + (255,), stroke_width=stroke_w,
                stroke_fill=ink + (255,))
        ang = rng.uniform(-1.8, 1.8)
        if rng.random() < 0.35:  # 偶发更大的字形摆动，模拟手写不一致
            ang = rng.uniform(-3.2, 3.2)
        ch_img = ch_img.rotate(ang, resample=Image.BICUBIC, expand=True)
        y = int(baseline - fs + dy)
        img.alpha_composite(ch_img, (int(x) - pad, y))
        adv = d.textlength(ch, font=ff) * rng.uniform(spacing_lo, spacing_hi)
        x += adv + 1

    bbox = img.getbbox()
    if bbox is None:
        raise ValueError("签名渲染失败")
    img = img.crop(bbox)

    # 整体斜切（向右上倾，模拟签名习惯）
    k = rng.uniform(-0.20, -0.03) * slant
    w0, h0 = img.size
    new_w = w0 + int(abs(k) * h0) + 2
    sheared = Image.new("RGBA", (new_w, h0), (0, 0, 0, 0))
    # AFFINE: x' = x + k*(h - y)  → 顶部右移
    sheared.paste(img.transform((new_w, h0), Image.AFFINE,
                                (1, -k, -k * h0, 0, 1, 0),
                                resample=Image.BICUBIC), (0, 0))
    img = sheared

    # 整体微旋转
    img = img.rotate(rng.uniform(-3.5, 1.8), resample=Image.BICUBIC,
                     expand=True, fillcolor=(0, 0, 0, 0))
    img = img.crop(img.getbbox())

    # 下划线（一笔收尾的弧线）
    if underline and rng.random() < 0.45:
        w1, h1 = img.size
        pad = int(h1 * 0.25) + 6
        canvas = Image.new("RGBA", (w1 + pad * 2, h1 + pad * 2), (0, 0, 0, 0))
        canvas.alpha_composite(img, (pad, pad))
        d2 = ImageDraw.Draw(canvas)
        x0 = pad + w1 * rng.uniform(-0.02, 0.04)
        x1 = pad + w1 * rng.uniform(0.96, 1.04)
        y = pad + h1 + pad * 0.55
        bow = h1 * rng.uniform(0.10, 0.22) * rng.choice([1, 1, -1])
        pts = []
        n_pts = 16
        for t in range(n_pts + 1):
            tt = t / n_pts
            px = x0 + (x1 - x0) * tt
            py = y + math.sin(tt * math.pi) * bow + rng.uniform(-1, 1)
            pts.append((px, py))
        lw = max(2, int(h1 * rng.uniform(0.028, 0.05)))
        d2.line(pts, fill=ink + (255,), width=lw, joint="curve")
        img = canvas.crop(canvas.getbbox())

    # 缩放到目标尺寸（保持比例，限制在画布内，留 6% 边距）
    tw, th = int(width * 0.92), int(height * 0.88)
    scale = min(tw / img.width, th / img.height)
    img = img.resize((max(1, int(img.width * scale)), max(1, int(img.height * scale))),
                     Image.LANCZOS)
    # 轻微墨迹柔化
    img = img.filter(ImageFilter.GaussianBlur(0.4))
    return img

# ---------------------------------------------------------------- Excel 处理

NAME_COL_KEYWORDS = ["姓名", "名字", "学生", "员工", "人员", "成员", "name"]


def read_excel_sheets(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheets = {}
    for ws in wb.worksheets:
        rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            rows.append(["" if v is None else str(v).strip() for v in row])
            if i >= 2000:
                break
        sheets[ws.title] = rows
    wb.close()
    return sheets


def guess_name_col(header):
    best = None
    for idx, cell in enumerate(header):
        c = (cell or "").strip().lower()
        if not c:
            continue
        for kw in NAME_COL_KEYWORDS:
            if kw in c:
                best = idx
                if c == "姓名" or c == "name":
                    return idx
                break
    return best


def export_excel(template_path, sheet_name, names, sig_dir, out_path,
                 col_name="签名", width=300, height=120):
    """在模板副本中插入签名图片列。返回 (写入行数)"""
    wb = openpyxl.load_workbook(template_path)
    ws = wb[sheet_name]
    # 找签名列：已存在同名表头则复用，否则追加到最后
    header = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    sig_col = None
    for idx, cell in enumerate(header, start=1):
        if (cell or "").strip() == col_name:
            sig_col = idx
            break
    if sig_col is None:
        sig_col = ws.max_column + 1
        ws.cell(row=1, column=sig_col, value=col_name)

    col_w = (width / 7.0) + 2
    ws.column_dimensions[get_column_letter(sig_col)].width = col_w

    n = 0
    for i, name in enumerate(names):
        row = i + 2  # 首行表头
        png = sig_dir / f"{i + 1:04d}_{safe_name(name)}.png"
        img = XLImage(str(png))
        # 缩放到单元格
        scale = min(width / img.width, height / img.height)
        disp_w, disp_h = int(img.width * scale), int(img.height * scale)
        marker = AnchorMarker(col=sig_col - 1, colOff=pixels_to_EMU(4),
                              row=row - 1, rowOff=pixels_to_EMU(4))
        img.anchor = OneCellAnchor(_from=marker,
                                   ext=XDRPositiveSize2D(pixels_to_EMU(disp_w),
                                                         pixels_to_EMU(disp_h)))
        ws.add_image(img)
        rh = ws.row_dimensions[row].height
        need = height * 0.75 + 6
        if rh is None or rh < need:
            ws.row_dimensions[row].height = need
        n += 1
    wb.save(out_path)
    return n


def safe_name(s):
    for ch in '\\/:*?"<>|\n\r\t':
        s = s.replace(ch, "_")
    return s.strip() or "unnamed"

# ---------------------------------------------------------------- GUI

SAMPLE_COLS = 4
PREVIEW_LIMIT = 24


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.minsize(820, 600)

        self.fonts = scan_fonts()
        self.font_vars = {}
        self.xlsx_path = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.name_col_var = tk.StringVar()
        self.width_var = tk.IntVar(value=300)
        self.height_var = tk.IntVar(value=120)
        self.ink_var = tk.StringVar(value="随机")
        self.slant_var = tk.DoubleVar(value=0.6)
        self.compact_var = tk.DoubleVar(value=0.45)
        self.bold_var = tk.BooleanVar(value=True)
        self.underline_var = tk.BooleanVar(value=False)
        self.header_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="就绪。先选择包含姓名的 Excel 文件。")
        self.seed = random.randrange(1 << 30)
        self.names = []
        self.rows_data = []
        self.preview_refs = []
        self.result_q = queue.Queue()
        self.exporting = False

        self._build_ui()
        self._poll_queue()

    # ---------- UI ----------
    def _build_ui(self):
        pad = dict(padx=8, pady=4)
        top = ttk.Frame(self.root)
        top.pack(fill="x", **pad)

        ttk.Button(top, text="选择 Excel 模板", command=self.pick_file).pack(side="left")
        ttk.Label(top, textvariable=self.xlsx_path, foreground="#555").pack(
            side="left", padx=8, fill="x", expand=True)

        row2 = ttk.Frame(self.root)
        row2.pack(fill="x", **pad)
        ttk.Label(row2, text="工作表:").pack(side="left")
        self.sheet_cb = ttk.Combobox(row2, textvariable=self.sheet_var, state="readonly",
                                     width=14)
        self.sheet_cb.pack(side="left", padx=4)
        self.sheet_cb.bind("<<ComboboxSelected>>", lambda e: self.on_sheet_change())
        ttk.Label(row2, text="姓名列:").pack(side="left", padx=(12, 0))
        self.name_cb = ttk.Combobox(row2, textvariable=self.name_col_var,
                                    state="readonly", width=14)
        self.name_cb.pack(side="left", padx=4)

        params = ttk.LabelFrame(self.root, text="生成参数")
        params.pack(fill="x", **pad)

        prow = ttk.Frame(params)
        prow.pack(fill="x", padx=8, pady=4)
        ttk.Label(prow, text="画布:").pack(side="left")
        ttk.Spinbox(prow, from_=120, to=800, textvariable=self.width_var, width=6).pack(side="left", padx=2)
        ttk.Label(prow, text="×").pack(side="left")
        ttk.Spinbox(prow, from_=60, to=400, textvariable=self.height_var, width=6).pack(side="left", padx=2)
        ttk.Label(prow, text="px").pack(side="left")
        ttk.Label(prow, text="    墨色:").pack(side="left")
        for m in ["随机", "纯黑", "碳黑", "蓝黑", "深蓝"]:
            ttk.Radiobutton(prow, text=m, value=m, variable=self.ink_var).pack(side="left")

        frow = ttk.Frame(params)
        frow.pack(fill="x", padx=8, pady=2)
        ttk.Label(frow, text="字体(随机抽取):").pack(side="left")
        if not self.fonts:
            ttk.Label(frow, text="未发现字体，请点击右侧【下载字体】", foreground="red").pack(side="left")
        for disp, path in self.fonts.items():
            v = tk.BooleanVar(value=True)
            self.font_vars[disp] = v
            cb = ttk.Checkbutton(frow, text=disp, variable=v, command=self.refresh_preview)
            cb.pack(side="left", padx=4)
        ttk.Button(frow, text="下载字体", command=self.download_fonts).pack(side="left", padx=8)
        ttk.Button(frow, text="打开字体目录", command=lambda: self.open_dir(FONT_DIR)).pack(side="left")

        srow = ttk.Frame(params)
        srow.pack(fill="x", padx=8, pady=2)
        ttk.Label(srow, text="倾斜:").pack(side="left")
        ttk.Scale(srow, from_=0, to=1, variable=self.slant_var, length=130,
                  command=lambda e: None).pack(side="left", padx=4)
        ttk.Label(srow, text="连笔紧凑:").pack(side="left", padx=(10, 0))
        ttk.Scale(srow, from_=0, to=1, variable=self.compact_var, length=130).pack(side="left", padx=4)
        ttk.Checkbutton(srow, text="笔画加粗", variable=self.bold_var).pack(side="left", padx=(14, 0))
        ttk.Checkbutton(srow, text="随机下划线", variable=self.underline_var).pack(side="left", padx=4)
        ttk.Checkbutton(srow, text="首行为表头", variable=self.header_var).pack(side="left", padx=4)

        btns = ttk.Frame(self.root)
        btns.pack(fill="x", **pad)
        self.btn_preview = ttk.Button(btns, text="生成预览", command=self.refresh_preview)
        self.btn_preview.pack(side="left")
        ttk.Button(btns, text="重新随机 ⟳", command=self.reroll).pack(side="left", padx=6)
        self.btn_xlsx = ttk.Button(btns, text="导出 Excel（内嵌签名）", command=self.export_xlsx)
        self.btn_xlsx.pack(side="left", padx=6)
        self.btn_imgs = ttk.Button(btns, text="仅导出图片文件夹", command=self.export_images)
        self.btn_imgs.pack(side="left")
        self.progress = ttk.Progressbar(btns, length=200, mode="determinate")
        self.progress.pack(side="right", padx=6)

        mid = ttk.LabelFrame(self.root, text="预览（前 %d 个姓名，导出结果与预览一致）" % PREVIEW_LIMIT)
        mid.pack(fill="both", expand=True, **pad)
        self.canvas = tk.Canvas(mid, bg="#f2f2f2", highlightthickness=0)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda e: self.canvas.yview_scroll(0, "units"))
        self.inner = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(
            scrollregion=self.canvas.bbox("all")))

        status = ttk.Label(self.root, textvariable=self.status_var, anchor="w",
                           foreground="#444")
        status.pack(fill="x", padx=10, pady=(0, 6))

    # ---------- 文件 ----------
    def pick_file(self):
        p = filedialog.askopenfilename(title="选择 Excel 模板",
                                       filetypes=[("Excel", "*.xlsx"), ("所有文件", "*.*")])
        if not p:
            return
        try:
            self.rows_data = read_excel_sheets(p)
        except Exception as e:
            messagebox.showerror("读取失败", f"无法读取该文件：\n{e}\n\n仅支持 .xlsx 格式。")
            return
        self.xlsx_path.set(p)
        self.sheet_cb["values"] = list(self.rows_data.keys())
        self.sheet_var.set(list(self.rows_data.keys())[0])
        self.on_sheet_change()

    def on_sheet_change(self):
        rows = self.rows_data.get(self.sheet_var.get(), [])
        if not rows:
            return
        header = rows[0] if self.header_var.get() else ["列%d" % (i + 1) for i in range(len(rows[0] or []))]
        self.name_cb["values"] = [f"{i + 1}. {h}" for i, h in enumerate(header)]
        guess = guess_name_col(header)
        if guess is not None:
            self.name_col_var.set(f"{guess + 1}. {header[guess]}")
        self.load_names()

    def load_names(self):
        rows = self.rows_data.get(self.sheet_var.get(), [])
        if not rows:
            return
        try:
            col_idx = int(self.name_col_var.get().split(".")[0]) - 1
        except (ValueError, IndexError):
            self.names = []
            return
        start = 1 if self.header_var.get() else 0
        names = []
        for r in rows[start:]:
            if col_idx < len(r) and r[col_idx].strip():
                names.append(r[col_idx].strip())
        self.names = names
        self.status_var.set(f"已加载 {len(names)} 个姓名（工作表: {self.sheet_var.get()}）")

    # ---------- 渲染 ----------
    def pick_font(self, rng):
        avail = [p for d, v in self.font_vars.items() if v.get() for p in [self.fonts.get(d)] if p]
        if not avail:
            raise ValueError("请至少勾选一款字体")
        return rng.choice(avail)

    def gen_one(self, name):
        rng = random.Random(f"{self.seed}|{name}")
        font = self.pick_font(rng)
        return render_signature(
            name, font,
            width=self.width_var.get(), height=self.height_var.get(),
            seed=self.seed,
            slant=self.slant_var.get(), compact=self.compact_var.get(),
            bold=self.bold_var.get(), underline=self.underline_var.get(),
            ink_mode=self.ink_var.get(),
        )

    def refresh_preview(self):
        if not self.names:
            self.load_names()
        for w in self.inner.winfo_children():
            w.destroy()
        self.preview_refs.clear()
        if not self.names:
            return
        cw, ch = self.width_var.get() // 2 + 16, self.height_var.get() // 2 + 30
        try:
            targets = self.names[:PREVIEW_LIMIT]
            for i, name in enumerate(targets):
                r, c = divmod(i, SAMPLE_COLS)
                cell = ttk.Frame(self.inner, padding=4)
                cell.grid(row=r, column=c, padx=4, pady=4, sticky="n")
                try:
                    im = self.gen_one(name)
                except Exception as e:
                    tk.Label(cell, text=f"{name}: {e}", foreground="red").pack()
                    continue
                disp = Image.new("RGBA", im.size, (255, 255, 255, 255))
                disp.alpha_composite(im)
                scale = (cw - 16) / disp.width
                disp = disp.resize((int(disp.width * scale), int(disp.height * scale)), Image.LANCZOS)
                photo = ImageTk.PhotoImage(disp)
                self.preview_refs.append(photo)
                card = tk.Label(cell, image=photo, bg="white",
                                relief="solid", borderwidth=1)
                card.pack()
                tk.Label(cell, text=name, foreground="#333").pack()
        except Exception as e:
            messagebox.showerror("预览失败", str(e))

    def reroll(self):
        self.seed = random.randrange(1 << 30)
        self.refresh_preview()

    # ---------- 导出 ----------
    def _do_export(self, mode):
        if self.exporting:
            return
        if not self.xlsx_path.get():
            messagebox.showwarning("提示", "请先选择 Excel 模板文件")
            return
        self.load_names()
        if not self.names:
            messagebox.showwarning("提示", "没有读到任何姓名，请检查工作表 / 姓名列 / 首行表头设置")
            return
        out_dir = filedialog.askdirectory(title="选择输出目录")
        if not out_dir:
            return
        out_dir = Path(out_dir)
        self.exporting = True
        self._set_buttons(False)
        threading.Thread(target=self._export_worker,
                         args=(mode, out_dir), daemon=True).start()

    def _export_worker(self, mode, out_dir):
        try:
            total = len(self.names)
            sig_dir = out_dir / f"签名图片_{safe_name(self.sheet_var.get())}"
            sig_dir.mkdir(parents=True, exist_ok=True)
            for i, name in enumerate(self.names):
                im = self.gen_one(name)
                im.save(sig_dir / f"{i + 1:04d}_{safe_name(name)}.png")
                self.result_q.put(("progress", (i + 1) * 70 // total))
            if mode == "xlsx":
                stamp = ""
                tpl = self.xlsx_path.get()
                out_xlsx = out_dir / (Path(tpl).stem + "_含签名.xlsx")
                export_excel(tpl, self.sheet_var.get(), self.names, sig_dir,
                             out_xlsx, width=self.width_var.get(),
                             height=self.height_var.get())
                self.result_q.put(("done", f"完成！\n签名图片: {sig_dir}\nExcel: {out_xlsx}\n\n注意：输出 Excel 由模板数据重建，原文件中的图片/图表不会保留。"))
            else:
                self.result_q.put(("done", f"完成！共 {total} 张签名图片\n目录: {sig_dir}"))
        except Exception as e:
            self.result_q.put(("error", str(e)))

    def export_xlsx(self):
        self._do_export("xlsx")

    def export_images(self):
        self._do_export("images")

    def _set_buttons(self, enabled):
        state = "normal" if enabled else "disabled"
        for b in (self.btn_preview, self.btn_xlsx, self.btn_imgs):
            b.configure(state=state)

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.result_q.get_nowait()
                if kind == "progress":
                    self.progress["value"] = payload
                elif kind == "done":
                    self.progress["value"] = 100
                    self.status_var.set(payload.splitlines()[0])
                    self.exporting = False
                    self._set_buttons(True)
                    messagebox.showinfo("导出成功", payload)
                elif kind == "error":
                    self.progress["value"] = 0
                    self.exporting = False
                    self._set_buttons(True)
                    messagebox.showerror("导出失败", payload)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ---------- 字体下载 ----------
    def download_fonts(self):
        FONT_DIR.mkdir(exist_ok=True)
        missing = [f for f in BUILTIN_FONTS.values()
                   if not (FONT_DIR / f).exists() and not (BUILTIN_FONT_DIR / f).exists()]
        if not missing:
            messagebox.showinfo("提示", f"内置字体已齐全（{FONT_DIR}）")
            return
        import urllib.request
        ok, fail = [], []
        for f in missing:
            slug = BUILTIN_FONT_SLUGS.get(f)
            got = False
            for tpl in FONT_DOWNLOAD_URLS:
                url = tpl.format(slug=slug, file=f)
                try:
                    with urllib.request.urlopen(url, timeout=60) as r, open(FONT_DIR / f, "wb") as w:
                        shutil.copyfileobj(r, w)
                    got = True
                    break
                except Exception:
                    continue
            (ok if got else fail).append(f)
        self.fonts = scan_fonts()
        messagebox.showinfo("字体下载", f"成功: {len(ok)} 个\n失败: {len(fail) or 0} 个 {fail if fail else ''}\n请重新打开工具以刷新字体列表。")

    def open_dir(self, d):
        d.mkdir(exist_ok=True)
        webbrowser.open(str(d))


def _selftest():
    """无 GUI 自检：验证字体扫描与签名渲染（CI 冒烟测试用）"""
    try:
        fonts = scan_fonts()
        if not fonts:
            raise RuntimeError("no fonts found")
        ok = 0
        for i, name in enumerate(["张三", "李四", "王五"]):
            rng = random.Random(f"ci|{name}")
            im = render_signature(name, list(fonts.values())[i % len(fonts)],
                                  300, 120, seed=12345)
            if im.getbbox() is None:
                raise RuntimeError(f"blank render: {name}")
            ok += 1
        msg = f"selftest ok: {len(fonts)} fonts, {ok} signatures rendered"
        if sys.stdout:
            print(msg)
        return 0
    except Exception as e:
        if sys.stderr:
            print(f"selftest failed: {e}", file=sys.stderr)
        return 1


def main():
    if "--selftest" in sys.argv:
        sys.exit(_selftest())

    # Windows 高分屏适配，避免界面模糊
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    # 按屏幕大小自适应窗口，兼容 1366x768 小屏
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    w, h = min(980, int(sw * 0.92)), min(760, int(sh * 0.86))
    root.geometry(f"{w}x{h}+{(sw - w) // 2}+{max(0, (sh - h) // 2 - 20)}")
    try:
        style = ttk.Style(root)
        if "aqua" in style.theme_names():
            style.theme_use("aqua")
        elif "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
