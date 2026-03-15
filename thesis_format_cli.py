"""Unified CLI / GUI entry point for universal thesis formatter.

Supports: .docx .doc .txt .md .tex
- With --input: CLI mode
- Without args: tkinter GUI
"""

import argparse
import copy
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading

from preprocess_txt_to_md import preprocess
from thesis_config import DEFAULT_CONFIG, resolve_config, dump_default_config
from thesis_format_2024 import apply_format
from word_postprocess import postprocess


def find_pandoc():
    """Locate pandoc: exe sibling dir -> _MEIPASS -> PATH."""
    candidates = []
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(sys.executable), "pandoc.exe"))
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    candidates.append(os.path.join(base, "pandoc.exe"))
    for c in candidates:
        if os.path.isfile(c):
            return c
    found = shutil.which("pandoc")
    if found:
        return found
    return None


def convert_doc_to_docx(doc_path, out_docx):
    """Convert .doc to .docx via Word COM."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(os.path.abspath(doc_path))
        doc.SaveAs(os.path.abspath(out_docx), 12)
        doc.Close()
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def run_format(input_path, output_path, skip_postprocess, log,
               config=None, config_path=None):
    """Core formatting pipeline. log(str) receives progress messages."""
    ext = os.path.splitext(input_path)[1].lower()
    supported = {".docx", ".doc", ".txt", ".md", ".tex"}
    if ext not in supported:
        log(f"不支持的格式: {ext} (支持: {' '.join(sorted(supported))})")
        return False

    # Resolve config
    if config is None:
        config, config_path = resolve_config(input_path=input_path)
    school = config.get("meta", {}).get("school_name", "")

    tmp_dir = tempfile.mkdtemp(prefix="thesisfmt_")
    tmp_docx = os.path.join(tmp_dir, "input.docx")

    try:
        # Stage 1: Convert to docx
        if ext == ".docx":
            shutil.copy2(input_path, tmp_docx)
            log("[1/3] 输入为 docx，直接复制。")
        elif ext == ".doc":
            log("[1/3] 通过 Word COM 转换 .doc...")
            convert_doc_to_docx(input_path, tmp_docx)
            log("[1/3] 转换完成。")
        elif ext in (".txt", ".md", ".tex"):
            pandoc = find_pandoc()
            if not pandoc:
                log("错误: 未找到 pandoc。请将 pandoc.exe 放在程序同目录或加入 PATH。")
                return False
            if ext == ".txt":
                log("[1/3] 预处理 txt -> md...")
                tmp_md = os.path.join(tmp_dir, "input.md")
                preprocess(input_path, tmp_md)
                source, fmt_from = tmp_md, "markdown-smart"
            elif ext == ".md":
                source, fmt_from = input_path, "markdown-smart"
            else:
                source, fmt_from = input_path, "latex"
            log(f"[1/3] pandoc 转换中 ({fmt_from} -> docx)...")
            ret = subprocess.run(
                [pandoc, source, f"--from={fmt_from}", "--to=docx", "--standalone", "-o", tmp_docx],
                capture_output=True, text=True,
            )
            if ret.returncode != 0:
                log(f"pandoc 失败:\n{ret.stderr}")
                return False
            log("[1/3] 转换完成。")

        # Stage 2: Format
        label = f"{school} " if school else ""
        log(f"[2/3] 应用 {label}格式规范...")
        fmt_warnings = apply_format(tmp_docx, output_path, config=config, config_path=config_path) or []
        log("[2/3] 格式化完成。")
        for w in fmt_warnings:
            log(w)

        # Stage 3: Post-process
        if not skip_postprocess:
            log("[3/3] Word COM 后处理（更新目录）...")
            try:
                postprocess(output_path, config=config)
                log("[3/3] 后处理完成。")
            except Exception as e:
                log(f"[3/3] 后处理失败（非致命）: {e}")
                log("[3/3] 已跳过。可在 Word 中手动更新目录。")
        else:
            log("[3/3] 已跳过目录更新。")

        log(f"\n输出文件: {output_path}")
        return True
    except Exception as e:
        log(f"\n错误: {e}")
        return False
    finally:
        if os.path.isdir(tmp_dir):
            shutil.rmtree(tmp_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------


class FormatterGUI:
    FILETYPES = [
        ("所有支持格式", "*.docx *.doc *.txt *.md *.tex"),
        ("Word 文档", "*.docx *.doc"),
        ("文本/Markdown", "*.txt *.md"),
        ("LaTeX", "*.tex"),
    ]
    CATEGORIES = ["页面", "正文", "标题", "页眉页码", "目录参考", "图表", "封面声明"]
    PT_SIZES = ["9pt", "10.5pt", "12pt", "14pt", "16pt", "18pt", "22pt", "24pt", "26pt", "36pt"]
    ALIGN_LABELS = ["左对齐", "居中", "右对齐", "两端对齐"]
    ALIGN_LABELS_KEEP = ["保持原样", "左对齐", "居中", "右对齐", "两端对齐"]
    _ALIGN = {"左对齐": "left", "居中": "center", "右对齐": "right", "两端对齐": "justify", "保持原样": "keep"}
    _ALIGN_R = {v: k for k, v in _ALIGN.items()}
    BOLD_LABELS = ["加粗", "不加粗", "保持原样"]
    _BOLD = {"加粗": True, "不加粗": False, "保持原样": "keep"}
    _BOLD_R = {True: "加粗", False: "不加粗", "keep": "保持原样"}
    _FM_MODE = {"自动识别": "auto", "跳过（不处理）": "skip", "强制格式化": "format"}
    _FM_MODE_R = {v: k for k, v in _FM_MODE.items()}
    _PGFMT = {"大写罗马 (I, II, III)": "upperRoman", "小写罗马 (i, ii, iii)": "lowerRoman", "阿拉伯数字 (1, 2, 3)": "decimal"}
    _PGFMT_R = {v: k for k, v in _PGFMT.items()}
    PGFMT_LABELS = list(_PGFMT.keys())
    _PGPOS = {"居中": "center", "居左": "left", "居右": "right", "奇右偶左": "alternate"}
    _PGPOS_R = {v: k for k, v in _PGPOS.items()}
    PGPOS_LABELS = list(_PGPOS.keys())
    PGPOS_LABELS_SIMPLE = ["居中", "居左", "居右"]
    _HF_SCOPE = {"仅正文": "body", "全部": "all"}
    _HF_SCOPE_R = {v: k for k, v in _HF_SCOPE.items()}
    _BORDER_STYLE = {"单线": "single", "双线": "double"}
    _BORDER_STYLE_R = {v: k for k, v in _BORDER_STYLE.items()}
    HEADING_PRESETS = {
        "第X章 / X.X / X.X.X (SCAU)": {
            "h1": r"^第\s*\d+\s*章\b", "h2": r"^\d+\.\d+\s",
            "h3": r"^\d+\.\d+\.\d+\s", "h4": r"^\d+\.\d+\.\d+\.\d+\s",
        },
        "X / X.X / X.X.X (纯数字)": {
            "h1": r"^\d+\s", "h2": r"^\d+\.\d+\s",
            "h3": r"^\d+\.\d+\.\d+\s", "h4": r"^\d+\.\d+\.\d+\.\d+\s",
        },
        "一、/ (一) / 1. (中文序号)": {
            "h1": r"^[一二三四五六七八九十百]+、", "h2": r"^（[一二三四五六七八九十百]+）",
            "h3": r"^\d+\.\s", "h4": r"^\(\d+\)",
        },
        "Chapter X / X.X / X.X.X (英文)": {
            "h1": r"(?i)^Chapter\s+\d+", "h2": r"^\d+\.\d+\s",
            "h3": r"^\d+\.\d+\.\d+\s", "h4": r"^\d+\.\d+\.\d+\.\d+\s",
        },
    }
    PRESET_NAMES = list(HEADING_PRESETS.keys())

    def __init__(self):
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, ttk

        self._tk = tk
        self._ttk = ttk
        self._filedialog = filedialog
        self._messagebox = messagebox
        self._scrolledtext = scrolledtext

        # Windows high-DPI: declare per-monitor DPI awareness
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            pass

        root = self._root = tk.Tk()
        root.title("论文格式化工具")
        root.resizable(True, True)
        root.minsize(680, 580)

        self._msg_q = queue.Queue()
        self._running = False

        self._init_vars(tk)

        # --- main layout: sidebar | content ---
        top = ttk.Frame(root, padding=8)
        top.pack(fill="both", expand=True)

        sidebar = ttk.Frame(top, width=80)
        sidebar.pack(side="left", fill="y", padx=(0, 8))
        sidebar.pack_propagate(False)

        self._cat_list = tk.Listbox(
            sidebar, font=("Microsoft YaHei UI", 10), activestyle="none",
            selectbackground="#0078D4", selectforeground="white",
            exportselection=False,
        )
        for c in self.CATEGORIES:
            self._cat_list.insert("end", c)
        self._cat_list.pack(fill="both", expand=True)
        self._cat_list.bind("<<ListboxSelect>>", self._on_cat_select)

        self._content = ttk.Frame(top)
        self._content.pack(side="left", fill="both", expand=True)

        # build category panels
        self._panels = {}
        self._panel_canvas = {}
        for name, builder in [
            ("页面", self._build_page), ("正文", self._build_body),
            ("标题", self._build_heading), ("页眉页码", self._build_header_pn),
            ("目录参考", self._build_toc_ref), ("图表", self._build_caption),
            ("封面声明", self._build_cover_decl),
        ]:
            wrapper = ttk.Frame(self._content)
            canvas = tk.Canvas(wrapper, highlightthickness=0, borderwidth=0)
            vsb = ttk.Scrollbar(wrapper, orient="vertical", command=canvas.yview)
            inner = ttk.Frame(canvas, padding=8)
            inner.bind("<Configure>",
                       lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
            canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=vsb.set)
            canvas.pack(side="left", fill="both", expand=True)
            vsb.pack(side="right", fill="y")
            builder(inner)
            self._panels[name] = wrapper
            self._panel_canvas[name] = canvas

        # --- bottom bar ---
        self._build_bottom(root)

        # show first panel & load defaults
        self._cur_panel = None
        self._cat_list.selection_set(0)
        self._show_panel("页面")
        self._load_vars_from_config(copy.deepcopy(DEFAULT_CONFIG))

        root.mainloop()

    # ---- tk variable init ----

    def _init_vars(self, tk):
        c = DEFAULT_CONFIG
        # page
        self._v_mt = tk.DoubleVar(value=c["page"]["margins"]["top"])
        self._v_mb = tk.DoubleVar(value=c["page"]["margins"]["bottom"])
        self._v_ml = tk.DoubleVar(value=c["page"]["margins"]["left"])
        self._v_mr = tk.DoubleVar(value=c["page"]["margins"]["right"])
        self._v_gutter = tk.DoubleVar(value=c["page"]["gutter"])
        self._v_hdist = tk.DoubleVar(value=c["page"]["header_distance"])
        self._v_fdist = tk.DoubleVar(value=c["page"]["footer_distance"])
        # fonts
        self._v_flat = tk.StringVar(value=c["fonts"]["latin"])
        self._v_fbody = tk.StringVar(value=c["fonts"]["body"])
        self._v_fh1 = tk.StringVar(value=c["fonts"]["h1"])
        self._v_fh2 = tk.StringVar(value=c["fonts"]["h2"])
        self._v_fh3 = tk.StringVar(value=c["fonts"]["h3"])
        self._v_fh4 = tk.StringVar(value=c["fonts"]["h4"])
        # sizes
        self._v_sbody = tk.StringVar(value=str(c["sizes"]["body"]) + "pt")
        self._v_sh1 = tk.StringVar(value=str(c["sizes"]["h1"]) + "pt")
        self._v_sh2 = tk.StringVar(value=str(c["sizes"]["h2"]) + "pt")
        self._v_sh3 = tk.StringVar(value=str(c["sizes"]["h3"]) + "pt")
        self._v_sh4 = tk.StringVar(value=str(c["sizes"]["h4"]) + "pt")
        self._v_scap = tk.StringVar(value=str(c["sizes"]["caption"]) + "pt")
        self._v_sfn = tk.StringVar(value=str(c["sizes"]["footnote"]) + "pt")
        # headings
        self._v_h1b = tk.StringVar(value=self._BOLD_R.get(c["headings"]["h1"]["bold"], "加粗"))
        self._v_h1a = tk.StringVar(value=self._ALIGN_R.get(c["headings"]["h1"]["align"], "左对齐"))
        self._v_h2b = tk.StringVar(value=self._BOLD_R.get(c["headings"]["h2"]["bold"], "加粗"))
        self._v_h2a = tk.StringVar(value=self._ALIGN_R.get(c["headings"]["h2"]["align"], "左对齐"))
        self._v_h3b = tk.StringVar(value=self._BOLD_R.get(c["headings"]["h3"]["bold"], "不加粗"))
        self._v_h3a = tk.StringVar(value=self._ALIGN_R.get(c["headings"]["h3"].get("align", "left"), "左对齐"))
        self._v_h4b = tk.StringVar(value=self._BOLD_R.get(c["headings"]["h4"]["bold"], "不加粗"))
        self._v_h4a = tk.StringVar(value=self._ALIGN_R.get(c["headings"]["h4"].get("align", "left"), "左对齐"))
        # heading spacing (lines)
        self._v_h1sb = tk.DoubleVar(value=c["headings"]["h1"].get("space_before", 0))
        self._v_h1sa = tk.DoubleVar(value=c["headings"]["h1"].get("space_after", 0))
        self._v_h2sb = tk.DoubleVar(value=c["headings"]["h2"].get("space_before", 0))
        self._v_h2sa = tk.DoubleVar(value=c["headings"]["h2"].get("space_after", 0))
        self._v_h3sb = tk.DoubleVar(value=c["headings"]["h3"].get("space_before", 0))
        self._v_h3sa = tk.DoubleVar(value=c["headings"]["h3"].get("space_after", 0))
        self._v_h4sb = tk.DoubleVar(value=c["headings"]["h4"].get("space_before", 0))
        self._v_h4sa = tk.DoubleVar(value=c["headings"]["h4"].get("space_after", 0))
        self._v_lsp = tk.DoubleVar(value=c["body"]["line_spacing"])
        self._v_ind = tk.DoubleVar(value=c["body"]["first_line_indent"])
        self._v_body_sb = tk.DoubleVar(value=c["body"].get("space_before", 0))
        self._v_body_sa = tk.DoubleVar(value=c["body"].get("space_after", 0))
        # heading numbering patterns
        sec = c["sections"]
        self._v_hpreset = tk.StringVar(value=self.PRESET_NAMES[0])
        self._v_pat_h1 = tk.StringVar(value=sec["chapter_pattern"])
        self._v_pat_h2 = tk.StringVar(value=sec["h2_pattern"])
        self._v_pat_h3 = tk.StringVar(value=sec["h3_pattern"])
        self._v_pat_h4 = tk.StringVar(value=sec["h4_pattern"])
        self._v_renum = tk.BooleanVar(value=sec.get("renumber_headings", True))
        # captions
        cap = c.get("captions", {})
        self._v_cap_fig = tk.StringVar(value=cap.get("figure_pattern", r"^图\s*\d"))
        self._v_cap_tbl = tk.StringVar(value=cap.get("table_pattern", r"^(续)?表\s*\d"))
        self._v_cap_sub = tk.StringVar(value=cap.get("subfigure_pattern", r"^\([a-z]\)"))
        self._v_cap_note = tk.StringVar(value=cap.get("note_pattern", r"^注[：:]"))
        self._v_cap_kwn = tk.BooleanVar(value=cap.get("keep_with_next", True))
        self._v_cap_chk = tk.BooleanVar(value=cap.get("check_numbering", True))
        # cover
        self._v_cov_en = tk.BooleanVar(value=c["cover"]["enabled"])
        self._v_school = tk.StringVar(value=c["meta"]["school_name"])
        self._v_logo = tk.StringVar(value=c["cover"]["logo"])
        self._v_covtitle = tk.StringVar(value=c["cover"]["title_text"])
        self._v_custom_cover = tk.StringVar()
        # declaration
        self._v_decl_en = tk.BooleanVar(value=True)
        # advanced
        self._v_tocd = tk.IntVar(value=c["toc"]["depth"])
        self._v_tocfont = tk.StringVar(value=c["toc"].get("font", c["fonts"]["body"]))
        self._v_tocsz = tk.StringVar(value=str(c["toc"].get("font_size", c["sizes"]["body"])) + "pt")
        self._v_tocls = tk.DoubleVar(value=c["toc"].get("line_spacing", c["body"]["line_spacing"]))
        self._v_toc_h1font = tk.StringVar(value=c["toc"].get("h1_font", c["fonts"]["h1"]))
        self._v_toc_h1sz = tk.StringVar(value=str(c["toc"].get("h1_font_size", c["sizes"]["h1"])) + "pt")
        self._v_toc_sb = tk.DoubleVar(value=c["toc"].get("space_before", 0))
        self._v_toc_sa = tk.DoubleVar(value=c["toc"].get("space_after", 0))
        self._v_refind = tk.DoubleVar(value=c["references"]["left_indent"])
        self._v_tbl_top = tk.DoubleVar(value=c["table"]["top_border_sz"] / 8)
        self._v_tbl_hdr = tk.DoubleVar(value=c["table"]["header_border_sz"] / 8)
        self._v_tbl_bot = tk.DoubleVar(value=c["table"]["bottom_border_sz"] / 8)
        self._v_pgfmt_f = tk.StringVar(value=self._PGFMT_R.get(c["page_numbers"]["front_format"], "大写罗马"))
        self._v_pgfmt_b = tk.StringVar(value=self._PGFMT_R.get(c["page_numbers"]["body_format"], "阿拉伯数字"))
        # header_footer
        hf = c["header_footer"]
        self._v_hf_en = tk.BooleanVar(value=hf["enabled"])
        self._v_hf_scope = tk.StringVar(value=self._HF_SCOPE_R.get(hf.get("scope", "body"), "仅正文"))
        self._v_hf_diff_oe = tk.BooleanVar(value=hf.get("different_odd_even", True))
        self._v_hf_first_no = tk.BooleanVar(value=hf.get("first_page_no_header", False))
        self._v_hf_odd_text = tk.StringVar(value=hf["odd_page_text"])
        self._v_hf_even_text = tk.StringVar(value=hf["even_page_text"])
        self._v_hf_odd_chap = tk.BooleanVar(value="{chapter_title}" in hf["odd_page_text"])
        self._v_hf_even_chap = tk.BooleanVar(value="{chapter_title}" in hf["even_page_text"])
        self._v_hf_font = tk.StringVar(value=hf["font"])
        self._v_hf_size = tk.StringVar(value=str(hf["font_size"]) + "pt")
        self._v_hf_bold = tk.BooleanVar(value=hf.get("bold", False))
        self._v_hf_odd_align = tk.StringVar(value=self._ALIGN_R.get(hf.get("odd_page_align", "center"), "居中"))
        self._v_hf_even_align = tk.StringVar(value=self._ALIGN_R.get(hf.get("even_page_align", "center"), "居中"))
        self._v_hf_border = tk.BooleanVar(value=hf.get("border_bottom", True))
        self._v_hf_bwidth = tk.DoubleVar(value=hf.get("border_bottom_width", 0.75))
        self._v_hf_bstyle = tk.StringVar(value=self._BORDER_STYLE_R.get(hf.get("border_bottom_style", "single"), "单线"))
        # page_numbers position
        pn = c["page_numbers"]
        self._v_pn_fpos = tk.StringVar(value=self._PGPOS_R.get(pn.get("front_position", "center"), "居中"))
        self._v_pn_bpos = tk.StringVar(value=self._PGPOS_R.get(pn.get("body_position", "center"), "居中"))
        self._v_pn_deco = tk.StringVar(value=pn.get("decorator", "{page}"))
        self._v_pn_font = tk.StringVar(value=pn.get("font", ""))
        self._v_pn_bold = tk.BooleanVar(value=pn.get("bold", False))
        self._v_pn_size = tk.StringVar(value=str(c["sizes"]["page_number"]) + "pt")
        self._v_pn_fstart = tk.IntVar(value=pn.get("front_start", 1))
        self._v_pn_bstart = tk.IntVar(value=pn.get("body_start", 1))
        # advanced extras
        self._v_body_align = tk.StringVar(value=self._ALIGN_R.get(c["body"]["align"], "两端对齐"))
        self._v_tbl_ls = tk.DoubleVar(value=c["table"]["line_spacing"])
        self._v_fn_ls = tk.DoubleVar(value=c["footnote"]["line_spacing"])
        # front_matter
        fm = c.get("front_matter", {})
        self._v_fm_mode = tk.StringVar(value=self._FM_MODE_R.get(fm.get("mode", "auto"), "自动识别"))
        # file I/O
        self._v_in = tk.StringVar()
        self._v_out = tk.StringVar()
        self._v_skip = tk.BooleanVar()
        self._v_cfglbl = tk.StringVar(value="默认 (SCAU)")

    # ---- row helpers ----

    def _row_spin(self, p, r, lbl, var, lo=0.0, hi=100.0, step=0.1, unit="cm"):
        self._ttk.Label(p, text=lbl).grid(row=r, column=0, sticky="w", pady=3)
        self._ttk.Spinbox(p, from_=lo, to=hi, increment=step,
                          textvariable=var, width=8).grid(row=r, column=1, sticky="w", padx=4, pady=3)
        if unit:
            self._ttk.Label(p, text=unit).grid(row=r, column=2, sticky="w", pady=3)
        return r + 1

    def _row_entry(self, p, r, lbl, var, w=28, hint=None):
        self._ttk.Label(p, text=lbl).grid(row=r, column=0, sticky="w", pady=3)
        if hint:
            self._ttk.Entry(p, textvariable=var, width=20).grid(
                row=r, column=1, sticky="w", padx=4, pady=3)
            self._ttk.Label(p, text=hint, foreground="gray").grid(
                row=r, column=2, sticky="w", pady=3)
        else:
            self._ttk.Entry(p, textvariable=var, width=w).grid(
                row=r, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        return r + 1

    def _row_combo(self, p, r, lbl, var, vals, w=10):
        self._ttk.Label(p, text=lbl).grid(row=r, column=0, sticky="w", pady=3)
        self._ttk.Combobox(p, textvariable=var, values=vals,
                           width=w, state="readonly").grid(row=r, column=1, sticky="w", padx=4, pady=3)
        return r + 1

    def _row_check(self, p, r, lbl, var):
        self._ttk.Checkbutton(p, text=lbl, variable=var).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=3)
        return r + 1

    def _sep(self, p, r):
        self._ttk.Separator(p, orient="horizontal").grid(
            row=r, column=0, columnspan=3, sticky="ew", pady=6)
        return r + 1

    # ---- panel builders ----

    def _build_page(self, p):
        r = 0
        r = self._row_spin(p, r, "上边距:", self._v_mt)
        r = self._row_spin(p, r, "下边距:", self._v_mb)
        r = self._row_spin(p, r, "左边距:", self._v_ml)
        r = self._row_spin(p, r, "右边距:", self._v_mr)
        r = self._row_spin(p, r, "装订线:", self._v_gutter)
        r = self._row_spin(p, r, "页眉距:", self._v_hdist)
        r = self._row_spin(p, r, "页脚距:", self._v_fdist)
        r = self._sep(p, r)
        r = self._row_combo(p, r, "前置页处理:", self._v_fm_mode, list(self._FM_MODE.keys()))
        self._ttk.Label(
            p, text="前置页包括封面、声明和摘要，不同学校设置差别大",
            foreground="gray", font=("Microsoft YaHei UI", 8)
        ).grid(row=r, column=0, columnspan=3, sticky="w", padx=18)
        r += 1

    def _build_header_pn(self, p):
        r = 0
        # -- 页眉 --
        self._ttk.Label(p, text="页眉", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(0, 4))
        r += 1
        r = self._row_check(p, r, "启用页眉", self._v_hf_en)
        r = self._row_combo(p, r, "作用范围:", self._v_hf_scope, list(self._HF_SCOPE.keys()))
        r = self._row_check(p, r, "奇偶页不同", self._v_hf_diff_oe)
        r = self._row_check(p, r, "首页不显示页眉", self._v_hf_first_no)
        # odd page (right side)
        r = self._row_entry(p, r, "奇数页(右):", self._v_hf_odd_text)
        r = self._row_check(p, r, "  ↳ 自动显示章标题", self._v_hf_odd_chap)
        r = self._row_combo(p, r, "奇数页对齐:", self._v_hf_odd_align, self.ALIGN_LABELS)
        # even page (left side)
        r = self._row_entry(p, r, "偶数页(左):", self._v_hf_even_text)
        r = self._row_check(p, r, "  ↳ 自动显示章标题", self._v_hf_even_chap)
        r = self._row_combo(p, r, "偶数页对齐:", self._v_hf_even_align, self.ALIGN_LABELS)
        r = self._row_entry(p, r, "页眉字体:", self._v_hf_font)
        r = self._row_combo(p, r, "页眉字号:", self._v_hf_size, self.PT_SIZES)
        r = self._row_check(p, r, "页眉文字加粗", self._v_hf_bold)
        r = self._row_check(p, r, "页眉下划线", self._v_hf_border)
        r = self._row_spin(p, r, "下划线粗细:", self._v_hf_bwidth, lo=0.25, hi=3.0, step=0.25, unit="磅")
        r = self._row_combo(p, r, "下划线样式:", self._v_hf_bstyle, list(self._BORDER_STYLE.keys()))
        r = self._sep(p, r)
        # -- 页码 --
        self._ttk.Label(p, text="页码", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(0, 4))
        r += 1
        r = self._row_combo(p, r, "前置页码位置:", self._v_pn_fpos, self.PGPOS_LABELS_SIMPLE)
        r = self._row_combo(p, r, "正文页码位置:", self._v_pn_bpos, self.PGPOS_LABELS)
        r = self._row_combo(p, r, "前置页码格式:", self._v_pgfmt_f, self.PGFMT_LABELS)
        r = self._row_combo(p, r, "正文页码格式:", self._v_pgfmt_b, self.PGFMT_LABELS)
        r = self._row_spin(p, r, "前置起始编号:", self._v_pn_fstart, lo=1, hi=999, step=1, unit="")
        r = self._row_spin(p, r, "正文起始编号:", self._v_pn_bstart, lo=1, hi=999, step=1, unit="")
        r = self._row_entry(p, r, "页码修饰:", self._v_pn_deco, hint="如: - {page} -")
        r = self._row_entry(p, r, "页码字体:", self._v_pn_font, hint="空=跟随西文")
        r = self._row_combo(p, r, "页码字号:", self._v_pn_size, self.PT_SIZES)
        r = self._row_check(p, r, "页码加粗", self._v_pn_bold)

    def _build_body(self, p):
        r = 0
        sz = self.PT_SIZES
        r = self._row_entry(p, r, "西文字体:", self._v_flat)
        r = self._row_entry(p, r, "正文中文字体:", self._v_fbody)
        r = self._row_combo(p, r, "正文字号:", self._v_sbody, sz)
        r = self._row_combo(p, r, "正文对齐:", self._v_body_align, self.ALIGN_LABELS_KEEP)
        r = self._row_spin(p, r, "首行缩进:", self._v_ind, lo=0, hi=100, step=1, unit="pt")
        r = self._row_spin(p, r, "行距:", self._v_lsp, lo=1.0, hi=3.0, step=0.25, unit="倍")
        r = self._row_spin(p, r, "段前:", self._v_body_sb, lo=0, hi=5, step=0.5, unit="行")
        r = self._row_spin(p, r, "段后:", self._v_body_sa, lo=0, hi=5, step=0.5, unit="行")
        r = self._sep(p, r)
        r = self._row_combo(p, r, "脚注字号:", self._v_sfn, sz)
        r = self._row_spin(p, r, "脚注行距:", self._v_fn_ls, lo=0.5, hi=3.0, step=0.25, unit="倍")

    def _build_heading(self, p):
        r = 0
        sz = self.PT_SIZES

        def _heading_block(r, label, v_font, v_size, v_bold, v_align, v_sb, v_sa):
            self._ttk.Label(p, text=label, font=("Microsoft YaHei UI", 9, "bold")).grid(
                row=r, column=0, columnspan=3, sticky="w", pady=(4, 2))
            r += 1
            r = self._row_entry(p, r, "  字体:", v_font)
            r = self._row_combo(p, r, "  字号:", v_size, sz)
            r = self._row_combo(p, r, "  加粗:", v_bold, self.BOLD_LABELS)
            r = self._row_combo(p, r, "  对齐:", v_align, self.ALIGN_LABELS_KEEP)
            self._ttk.Label(p, text="  段前:").grid(row=r, column=0, sticky="w", pady=3)
            sf = self._ttk.Frame(p)
            sf.grid(row=r, column=1, columnspan=2, sticky="w", padx=4, pady=3)
            self._ttk.Spinbox(sf, from_=-1, to=5, increment=0.5, textvariable=v_sb, width=5).pack(side="left")
            self._ttk.Label(sf, text="行  段后:").pack(side="left", padx=(4, 0))
            self._ttk.Spinbox(sf, from_=-1, to=5, increment=0.5, textvariable=v_sa, width=5).pack(side="left")
            self._ttk.Label(sf, text="行").pack(side="left")
            r += 1
            return r

        r = _heading_block(r, "一级标题 (H1)",
                           self._v_fh1, self._v_sh1, self._v_h1b, self._v_h1a,
                           self._v_h1sb, self._v_h1sa)
        r = _heading_block(r, "二级标题 (H2)",
                           self._v_fh2, self._v_sh2, self._v_h2b, self._v_h2a,
                           self._v_h2sb, self._v_h2sa)
        r = _heading_block(r, "三级标题 (H3)",
                           self._v_fh3, self._v_sh3, self._v_h3b, self._v_h3a,
                           self._v_h3sb, self._v_h3sa)
        r = _heading_block(r, "四级标题 (H4)",
                           self._v_fh4, self._v_sh4, self._v_h4b, self._v_h4a,
                           self._v_h4sb, self._v_h4sa)

        self._ttk.Label(p, text="(-1 = 保持原样)", foreground="gray").grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(0, 4))
        r += 1
        r = self._sep(p, r)
        r = self._row_check(p, r, "自动修正标题编号（检测缺失/跳号并重编号）", self._v_renum)
        self._ttk.Label(p, text="编号预设:").grid(row=r, column=0, sticky="w", pady=3)
        pcb = self._ttk.Combobox(p, textvariable=self._v_hpreset,
                                 values=self.PRESET_NAMES, width=28, state="readonly")
        pcb.grid(row=r, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        pcb.bind("<<ComboboxSelected>>", self._on_preset_select)
        r += 1
        r = self._row_entry(p, r, "一级标题:", self._v_pat_h1, hint="(如: 第1章)")
        r = self._row_entry(p, r, "二级标题:", self._v_pat_h2, hint="(如: 1.1)")
        r = self._row_entry(p, r, "三级标题:", self._v_pat_h3, hint="(如: 1.1.1)")
        r = self._row_entry(p, r, "四级标题:", self._v_pat_h4, hint="(如: 1.1.1.1)")

    def _on_preset_select(self, _event=None):
        preset = self.HEADING_PRESETS.get(self._v_hpreset.get())
        if preset:
            self._v_pat_h1.set(preset["h1"])
            self._v_pat_h2.set(preset["h2"])
            self._v_pat_h3.set(preset["h3"])
            self._v_pat_h4.set(preset["h4"])

    def _build_caption(self, p):
        r = 0
        sz = self.PT_SIZES
        r = self._row_combo(p, r, "图表题字号:", self._v_scap, sz)
        r = self._row_check(p, r, "图表题防分页 (keep with next)", self._v_cap_kwn)
        r = self._row_check(p, r, "检查图表编号连续性", self._v_cap_chk)
        r = self._sep(p, r)
        r = self._row_entry(p, r, "图题模式:", self._v_cap_fig, hint="(如: 图1)")
        r = self._row_entry(p, r, "表题模式:", self._v_cap_tbl, hint="(如: 表1)")
        r = self._row_entry(p, r, "分图模式:", self._v_cap_sub, hint="(如: (a))")
        r = self._row_entry(p, r, "表注模式:", self._v_cap_note, hint="(如: 注：)")
        r = self._sep(p, r)
        self._ttk.Label(p, text="三线表", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(0, 4))
        r += 1
        r = self._row_spin(p, r, "顶线粗细:", self._v_tbl_top, lo=0.25, hi=6, step=0.25, unit="磅")
        r = self._row_spin(p, r, "栏目线粗细:", self._v_tbl_hdr, lo=0.25, hi=6, step=0.25, unit="磅")
        r = self._row_spin(p, r, "底线粗细:", self._v_tbl_bot, lo=0.25, hi=6, step=0.25, unit="磅")
        r = self._row_spin(p, r, "表格行距:", self._v_tbl_ls, lo=0.5, hi=3.0, step=0.25, unit="倍")

    def _build_cover_decl(self, p):
        # -- 封面 --
        r = 0
        r = self._row_check(p, r, "启用封面", self._v_cov_en)
        self._ttk.Label(p, text="自定义封面:").grid(row=r, column=0, sticky="w", pady=3)
        cf = self._ttk.Frame(p)
        cf.grid(row=r, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        self._ttk.Entry(cf, textvariable=self._v_custom_cover, width=22).pack(side="left")
        self._ttk.Button(cf, text="浏览", width=5,
                         command=self._browse_custom_cover).pack(side="left", padx=4)
        r += 1
        self._ttk.Label(p, text="（上传已排好版的封面页 .docx，将替代自动生成封面）",
                        foreground="gray").grid(row=r, column=0, columnspan=3, sticky="w", pady=0)
        r += 1
        r = self._sep(p, r)
        r = self._row_entry(p, r, "学校名称:", self._v_school)
        self._ttk.Label(p, text="Logo 文件:").grid(row=r, column=0, sticky="w", pady=3)
        lf = self._ttk.Frame(p)
        lf.grid(row=r, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        self._ttk.Entry(lf, textvariable=self._v_logo, width=22).pack(side="left")
        self._ttk.Button(lf, text="浏览", width=5, command=self._browse_logo).pack(side="left", padx=4)
        r += 1
        r = self._row_entry(p, r, "封面标题:", self._v_covtitle)
        r = self._sep(p, r)
        self._ttk.Label(p, text="信息栏字段:").grid(row=r, column=0, sticky="nw", pady=3)
        self._cov_fields_frame = self._ttk.Frame(p)
        self._cov_fields_frame.grid(row=r, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        self._cov_field_rows = []
        r += 1
        bf = self._ttk.Frame(p)
        bf.grid(row=r, column=1, sticky="w", padx=4, pady=3)
        self._ttk.Button(bf, text="添加", width=6, command=self._add_cov_field).pack(side="left")
        self._ttk.Button(bf, text="删除末行", width=8, command=self._del_cov_field).pack(side="left", padx=4)
        r += 1
        # -- 声明 --
        r = self._sep(p, r)
        r = self._row_check(p, r, "启用声明页", self._v_decl_en)
        self._decl_widgets = []
        for idx, decl in enumerate(DEFAULT_CONFIG.get("declarations", [])):
            r = self._sep(p, r)
            self._ttk.Label(p, text=f"声明 {idx + 1}").grid(row=r, column=0, sticky="w", pady=3)
            r += 1
            tv = self._tk.StringVar(value=decl.get("title", ""))
            self._ttk.Label(p, text="标题:").grid(row=r, column=0, sticky="w", pady=2)
            self._ttk.Entry(p, textvariable=tv, width=42).grid(
                row=r, column=1, columnspan=2, sticky="w", padx=4, pady=2)
            r += 1
            self._ttk.Label(p, text="正文:").grid(row=r, column=0, sticky="nw", pady=2)
            bt = self._scrolledtext.ScrolledText(
                p, width=42, height=4, font=("Microsoft YaHei UI", 9))
            bt.grid(row=r, column=1, columnspan=2, sticky="w", padx=4, pady=2)
            bt.insert("1.0", decl.get("body", ""))
            r += 1
            self._decl_widgets.append({"title": tv, "body": bt, "orig": decl})

    def _build_toc_ref(self, p):
        r = 0
        self._ttk.Label(p, text="目录", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(0, 4))
        r += 1
        r = self._row_spin(p, r, "目录深度:", self._v_tocd, lo=1, hi=4, step=1, unit="级")
        r = self._row_entry(p, r, "二级条目字体:", self._v_tocfont)
        r = self._row_combo(p, r, "二级条目字号:", self._v_tocsz, self.PT_SIZES)
        r = self._row_entry(p, r, "一级条目字体:", self._v_toc_h1font)
        r = self._row_combo(p, r, "一级条目字号:", self._v_toc_h1sz, self.PT_SIZES)
        r = self._row_spin(p, r, "目录行距:", self._v_tocls, lo=1.0, hi=3.0, step=0.25, unit="倍")
        r = self._row_spin(p, r, "条目段前:", self._v_toc_sb, lo=0, hi=5, step=0.5, unit="行")
        r = self._row_spin(p, r, "条目段后:", self._v_toc_sa, lo=0, hi=5, step=0.5, unit="行")
        r = self._sep(p, r)
        self._ttk.Label(p, text="参考文献", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(0, 4))
        r += 1
        r = self._row_spin(p, r, "参考文献缩进:", self._v_refind, lo=0, hi=100, step=1, unit="pt")
        r = self._sep(p, r)
        # special titles
        self._ttk.Label(p, text="特殊标题映射", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(0, 4))
        r += 1
        sf = self._ttk.Frame(p)
        sf.grid(row=r, column=0, columnspan=3, sticky="w", padx=4, pady=3)
        self._st_frame = sf
        self._st_rows = []
        self._ttk.Label(sf, text="匹配").grid(row=0, column=0, padx=2)
        self._ttk.Label(sf, text="显示").grid(row=0, column=1, padx=2)
        self._ttk.Label(sf, text="对齐").grid(row=0, column=2, padx=2)
        r += 1
        bf = self._ttk.Frame(p)
        bf.grid(row=r, column=0, columnspan=3, sticky="w", padx=4, pady=3)
        self._ttk.Button(bf, text="添加", width=6, command=self._add_st).pack(side="left")
        self._ttk.Button(bf, text="删除末行", width=8, command=self._del_st).pack(side="left", padx=4)

    def _add_cov_field(self, label="", width=33):
        tk = self._tk
        row = len(self._cov_field_rows)
        f = self._cov_fields_frame
        lv = tk.StringVar(value=label)
        wv = tk.IntVar(value=width)
        le = self._ttk.Entry(f, textvariable=lv, width=16)
        le.grid(row=row, column=0, padx=(0, 4), pady=1)
        ws = self._ttk.Spinbox(f, from_=5, to=60, textvariable=wv, width=5)
        ws.grid(row=row, column=1, pady=1)
        self._cov_field_rows.append((lv, wv, le, ws))

    def _del_cov_field(self):
        if not self._cov_field_rows:
            return
        _, _, le, ws = self._cov_field_rows.pop()
        le.destroy()
        ws.destroy()

    def _add_st(self, match="", display="", align="center"):
        tk = self._tk
        row = len(self._st_rows) + 1
        f = self._st_frame
        mv = tk.StringVar(value=match)
        dv = tk.StringVar(value=display)
        av = tk.StringVar(value=self._ALIGN_R.get(align, "居中"))
        me = self._ttk.Entry(f, textvariable=mv, width=10)
        me.grid(row=row, column=0, padx=2, pady=1)
        de = self._ttk.Entry(f, textvariable=dv, width=14)
        de.grid(row=row, column=1, padx=2, pady=1)
        ac = self._ttk.Combobox(f, textvariable=av, values=self.ALIGN_LABELS, width=8, state="readonly")
        ac.grid(row=row, column=2, padx=2, pady=1)
        self._st_rows.append((mv, dv, av, me, de, ac))

    def _del_st(self):
        if not self._st_rows:
            return
        _, _, _, me, de, ac = self._st_rows.pop()
        me.destroy()
        de.destroy()
        ac.destroy()

    # ---- bottom bar ----

    def _build_bottom(self, root):
        self._ttk.Separator(root, orient="horizontal").pack(fill="x", padx=8)
        bot = self._ttk.Frame(root, padding=(8, 4))
        bot.pack(fill="x")

        # config row
        cr = self._ttk.Frame(bot)
        cr.pack(fill="x", pady=2)
        self._ttk.Label(cr, text="配置:").pack(side="left")
        self._ttk.Label(cr, textvariable=self._v_cfglbl, foreground="gray").pack(side="left", padx=4)
        self._ttk.Button(cr, text="加载", width=5, command=self._load_config).pack(side="right", padx=2)
        self._ttk.Button(cr, text="保存", width=5, command=self._save_config).pack(side="right", padx=2)
        self._ttk.Button(cr, text="默认(SCAU)", width=10, command=self._reset_defaults).pack(side="right", padx=2)

        # input row
        ir = self._ttk.Frame(bot)
        ir.pack(fill="x", pady=2)
        self._ttk.Label(ir, text="输入:").pack(side="left")
        self._ttk.Entry(ir, textvariable=self._v_in, width=52).pack(side="left", padx=4)
        self._ttk.Button(ir, text="浏览", width=5, command=self._browse_in).pack(side="left")

        # output row
        orr = self._ttk.Frame(bot)
        orr.pack(fill="x", pady=2)
        self._ttk.Label(orr, text="输出:").pack(side="left")
        self._ttk.Entry(orr, textvariable=self._v_out, width=52).pack(side="left", padx=4)
        self._ttk.Button(orr, text="浏览", width=5, command=self._browse_out).pack(side="left")

        # action row
        ar = self._ttk.Frame(bot)
        ar.pack(fill="x", pady=2)
        self._ttk.Checkbutton(ar, text="跳过目录生成（需已安装 Word）",
                              variable=self._v_skip).pack(side="left")
        self._btn = self._ttk.Button(ar, text="开始格式化", command=self._start)
        self._btn.pack(side="right")

        # log
        self._log = self._scrolledtext.ScrolledText(
            bot, width=70, height=8, state="disabled", font=("Microsoft YaHei UI", 9))
        self._log.pack(fill="x", pady=(4, 0))

    # ---- panel switching ----

    def _show_panel(self, name):
        if self._cur_panel:
            self._panels[self._cur_panel].pack_forget()
        self._panels[name].pack(fill="both", expand=True)
        self._cur_panel = name
        canvas = self._panel_canvas.get(name)
        if canvas:
            canvas.yview_moveto(0)
            root = canvas.winfo_toplevel()
            root.bind_all("<MouseWheel>",
                          lambda e, c=canvas: c.yview_scroll(
                              int(-1 * (e.delta / 120)), "units"))

    def _on_cat_select(self, _event=None):
        sel = self._cat_list.curselection()
        if sel:
            self._show_panel(self.CATEGORIES[sel[0]])

    # ---- config ↔ vars ----

    @staticmethod
    def _numval(v):
        """float → int if whole, else float."""
        return int(v) if v == int(v) else v

    def _collect_config(self):
        cfg = copy.deepcopy(DEFAULT_CONFIG)
        # page
        cfg["page"]["margins"]["top"] = self._v_mt.get()
        cfg["page"]["margins"]["bottom"] = self._v_mb.get()
        cfg["page"]["margins"]["left"] = self._v_ml.get()
        cfg["page"]["margins"]["right"] = self._v_mr.get()
        cfg["page"]["gutter"] = self._v_gutter.get()
        cfg["page"]["header_distance"] = self._v_hdist.get()
        cfg["page"]["footer_distance"] = self._v_fdist.get()
        # fonts
        cfg["fonts"]["latin"] = self._v_flat.get()
        cfg["fonts"]["body"] = self._v_fbody.get()
        cfg["fonts"]["h1"] = self._v_fh1.get()
        cfg["fonts"]["h2"] = self._v_fh2.get()
        cfg["fonts"]["h3"] = self._v_fh3.get()
        cfg["fonts"]["h4"] = self._v_fh4.get()
        # sizes
        cfg["sizes"]["body"] = self._numval(float(self._v_sbody.get().replace("pt", "")))
        cfg["sizes"]["h1"] = self._numval(float(self._v_sh1.get().replace("pt", "")))
        cfg["sizes"]["h2"] = self._numval(float(self._v_sh2.get().replace("pt", "")))
        cfg["sizes"]["h3"] = self._numval(float(self._v_sh3.get().replace("pt", "")))
        cfg["sizes"]["h4"] = self._numval(float(self._v_sh4.get().replace("pt", "")))
        cfg["sizes"]["caption"] = self._numval(float(self._v_scap.get().replace("pt", "")))
        cfg["sizes"]["footnote"] = self._numval(float(self._v_sfn.get().replace("pt", "")))
        # headings
        cfg["headings"]["h1"]["bold"] = self._BOLD.get(self._v_h1b.get(), True)
        cfg["headings"]["h1"]["align"] = self._ALIGN.get(self._v_h1a.get(), "left")
        cfg["headings"]["h1"]["space_before"] = self._v_h1sb.get()
        cfg["headings"]["h1"]["space_after"] = self._v_h1sa.get()
        cfg["headings"]["h2"]["bold"] = self._BOLD.get(self._v_h2b.get(), True)
        cfg["headings"]["h2"]["align"] = self._ALIGN.get(self._v_h2a.get(), "left")
        cfg["headings"]["h2"]["space_before"] = self._v_h2sb.get()
        cfg["headings"]["h2"]["space_after"] = self._v_h2sa.get()
        cfg["headings"]["h3"]["bold"] = self._BOLD.get(self._v_h3b.get(), False)
        cfg["headings"]["h3"]["align"] = self._ALIGN.get(self._v_h3a.get(), "left")
        cfg["headings"]["h3"]["space_before"] = self._v_h3sb.get()
        cfg["headings"]["h3"]["space_after"] = self._v_h3sa.get()
        cfg["headings"]["h4"]["bold"] = self._BOLD.get(self._v_h4b.get(), False)
        cfg["headings"]["h4"]["align"] = self._ALIGN.get(self._v_h4a.get(), "left")
        cfg["headings"]["h4"]["space_before"] = self._v_h4sb.get()
        cfg["headings"]["h4"]["space_after"] = self._v_h4sa.get()
        # body
        cfg["body"]["line_spacing"] = self._v_lsp.get()
        cfg["body"]["first_line_indent"] = self._numval(self._v_ind.get())
        cfg["body"]["align"] = self._ALIGN.get(self._v_body_align.get(), "justify")
        cfg["body"]["space_before"] = self._v_body_sb.get()
        cfg["body"]["space_after"] = self._v_body_sa.get()
        # sections (heading numbering patterns)
        cfg["sections"]["chapter_pattern"] = self._v_pat_h1.get()
        cfg["sections"]["h2_pattern"] = self._v_pat_h2.get()
        cfg["sections"]["h3_pattern"] = self._v_pat_h3.get()
        cfg["sections"]["h4_pattern"] = self._v_pat_h4.get()
        cfg["sections"]["renumber_headings"] = self._v_renum.get()
        # captions
        cfg["captions"] = {
            "figure_pattern": self._v_cap_fig.get(),
            "table_pattern": self._v_cap_tbl.get(),
            "subfigure_pattern": self._v_cap_sub.get(),
            "note_pattern": self._v_cap_note.get(),
            "keep_with_next": self._v_cap_kwn.get(),
            "check_numbering": self._v_cap_chk.get(),
        }
        # cover
        cfg["cover"]["enabled"] = self._v_cov_en.get()
        cfg["meta"]["school_name"] = self._v_school.get()
        cfg["cover"]["logo"] = self._v_logo.get()
        cfg["cover"]["title_text"] = self._v_covtitle.get()
        custom_cov = self._v_custom_cover.get().strip()
        if custom_cov:
            cfg["cover"]["custom_docx"] = custom_cov
        cfg["cover"]["fields"] = [
            {"label": lv.get(), "underline_chars": wv.get()}
            for lv, wv, _, _ in self._cov_field_rows
        ]
        # declarations
        if not self._v_decl_en.get():
            cfg["declarations"] = []
        else:
            decls = []
            for dw in self._decl_widgets:
                d = copy.deepcopy(dw["orig"])
                d["title"] = dw["title"].get()
                d["body"] = dw["body"].get("1.0", "end-1c")
                decls.append(d)
            cfg["declarations"] = decls
        # advanced
        cfg["toc"]["depth"] = self._v_tocd.get()
        cfg["toc"]["font"] = self._v_tocfont.get()
        cfg["toc"]["font_size"] = self._numval(float(self._v_tocsz.get().replace("pt", "")))
        cfg["toc"]["h1_font"] = self._v_toc_h1font.get()
        cfg["toc"]["h1_font_size"] = self._numval(float(self._v_toc_h1sz.get().replace("pt", "")))
        cfg["toc"]["line_spacing"] = self._v_tocls.get()
        cfg["toc"]["space_before"] = self._v_toc_sb.get()
        cfg["toc"]["space_after"] = self._v_toc_sa.get()
        cfg["references"]["left_indent"] = self._numval(self._v_refind.get())
        cfg["references"]["first_line_indent"] = -self._numval(self._v_refind.get())
        cfg["table"]["top_border_sz"] = self._numval(self._v_tbl_top.get() * 8)
        cfg["table"]["header_border_sz"] = self._numval(self._v_tbl_hdr.get() * 8)
        cfg["table"]["bottom_border_sz"] = self._numval(self._v_tbl_bot.get() * 8)
        cfg["table"]["line_spacing"] = self._v_tbl_ls.get()
        cfg["footnote"]["line_spacing"] = self._v_fn_ls.get()
        # page_numbers
        cfg["page_numbers"]["front_format"] = self._PGFMT.get(self._v_pgfmt_f.get(), "upperRoman")
        cfg["page_numbers"]["body_format"] = self._PGFMT.get(self._v_pgfmt_b.get(), "decimal")
        cfg["page_numbers"]["front_position"] = self._PGPOS.get(self._v_pn_fpos.get(), "center")
        cfg["page_numbers"]["body_position"] = self._PGPOS.get(self._v_pn_bpos.get(), "center")
        cfg["page_numbers"]["front_start"] = self._v_pn_fstart.get()
        cfg["page_numbers"]["body_start"] = self._v_pn_bstart.get()
        cfg["page_numbers"]["decorator"] = self._v_pn_deco.get()
        cfg["page_numbers"]["font"] = self._v_pn_font.get()
        cfg["page_numbers"]["bold"] = self._v_pn_bold.get()
        cfg["sizes"]["page_number"] = self._numval(float(self._v_pn_size.get().replace("pt", "")))
        # header_footer
        cfg["header_footer"]["enabled"] = self._v_hf_en.get()
        cfg["header_footer"]["scope"] = self._HF_SCOPE.get(self._v_hf_scope.get(), "body")
        cfg["header_footer"]["different_odd_even"] = self._v_hf_diff_oe.get()
        cfg["header_footer"]["first_page_no_header"] = self._v_hf_first_no.get()
        cfg["header_footer"]["odd_page_text"] = "{chapter_title}" if self._v_hf_odd_chap.get() \
            else self._v_hf_odd_text.get()
        cfg["header_footer"]["even_page_text"] = "{chapter_title}" if self._v_hf_even_chap.get() \
            else self._v_hf_even_text.get()
        cfg["header_footer"]["font"] = self._v_hf_font.get()
        cfg["header_footer"]["font_size"] = self._numval(float(self._v_hf_size.get().replace("pt", "")))
        cfg["header_footer"]["bold"] = self._v_hf_bold.get()
        cfg["header_footer"]["odd_page_align"] = self._ALIGN.get(self._v_hf_odd_align.get(), "center")
        cfg["header_footer"]["even_page_align"] = self._ALIGN.get(self._v_hf_even_align.get(), "center")
        cfg["header_footer"]["border_bottom"] = self._v_hf_border.get()
        cfg["header_footer"]["border_bottom_width"] = self._v_hf_bwidth.get()
        cfg["header_footer"]["border_bottom_style"] = self._BORDER_STYLE.get(self._v_hf_bstyle.get(), "single")
        # front_matter
        cfg["front_matter"] = {"mode": self._FM_MODE.get(self._v_fm_mode.get(), "auto")}
        cfg["special_titles"] = [
            {"match": m.get(), "display": d.get(),
             "align": self._ALIGN.get(a.get(), "center")}
            for m, d, a, _, _, _ in self._st_rows
        ]
        return cfg

    def _load_vars_from_config(self, cfg):
        # page
        self._v_mt.set(cfg["page"]["margins"]["top"])
        self._v_mb.set(cfg["page"]["margins"]["bottom"])
        self._v_ml.set(cfg["page"]["margins"]["left"])
        self._v_mr.set(cfg["page"]["margins"]["right"])
        self._v_gutter.set(cfg["page"]["gutter"])
        self._v_hdist.set(cfg["page"]["header_distance"])
        self._v_fdist.set(cfg["page"]["footer_distance"])
        # fonts
        self._v_flat.set(cfg["fonts"]["latin"])
        self._v_fbody.set(cfg["fonts"]["body"])
        self._v_fh1.set(cfg["fonts"]["h1"])
        self._v_fh2.set(cfg["fonts"]["h2"])
        self._v_fh3.set(cfg["fonts"]["h3"])
        self._v_fh4.set(cfg["fonts"]["h4"])
        # sizes
        self._v_sbody.set(str(self._numval(cfg["sizes"]["body"])) + "pt")
        self._v_sh1.set(str(self._numval(cfg["sizes"]["h1"])) + "pt")
        self._v_sh2.set(str(self._numval(cfg["sizes"]["h2"])) + "pt")
        self._v_sh3.set(str(self._numval(cfg["sizes"]["h3"])) + "pt")
        self._v_sh4.set(str(self._numval(cfg["sizes"]["h4"])) + "pt")
        self._v_scap.set(str(self._numval(cfg["sizes"]["caption"])) + "pt")
        self._v_sfn.set(str(self._numval(cfg["sizes"]["footnote"])) + "pt")
        # headings
        self._v_h1b.set(self._BOLD_R.get(cfg["headings"]["h1"]["bold"], "加粗"))
        self._v_h1a.set(self._ALIGN_R.get(cfg["headings"]["h1"]["align"], "左对齐"))
        self._v_h1sb.set(cfg["headings"]["h1"].get("space_before", 0))
        self._v_h1sa.set(cfg["headings"]["h1"].get("space_after", 0))
        self._v_h2b.set(self._BOLD_R.get(cfg["headings"]["h2"]["bold"], "加粗"))
        self._v_h2a.set(self._ALIGN_R.get(cfg["headings"]["h2"]["align"], "左对齐"))
        self._v_h2sb.set(cfg["headings"]["h2"].get("space_before", 0))
        self._v_h2sa.set(cfg["headings"]["h2"].get("space_after", 0))
        self._v_h3b.set(self._BOLD_R.get(cfg["headings"]["h3"]["bold"], "不加粗"))
        self._v_h3a.set(self._ALIGN_R.get(cfg["headings"]["h3"].get("align", "left"), "左对齐"))
        self._v_h3sb.set(cfg["headings"]["h3"].get("space_before", 0))
        self._v_h3sa.set(cfg["headings"]["h3"].get("space_after", 0))
        self._v_h4b.set(self._BOLD_R.get(cfg["headings"]["h4"]["bold"], "不加粗"))
        self._v_h4a.set(self._ALIGN_R.get(cfg["headings"]["h4"].get("align", "left"), "左对齐"))
        self._v_h4sb.set(cfg["headings"]["h4"].get("space_before", 0))
        self._v_h4sa.set(cfg["headings"]["h4"].get("space_after", 0))
        # body
        self._v_lsp.set(cfg["body"]["line_spacing"])
        self._v_ind.set(cfg["body"]["first_line_indent"])
        self._v_body_sb.set(cfg["body"].get("space_before", 0))
        self._v_body_sa.set(cfg["body"].get("space_after", 0))
        # heading numbering patterns
        sec = cfg.get("sections", {})
        self._v_pat_h1.set(sec.get("chapter_pattern", ""))
        self._v_pat_h2.set(sec.get("h2_pattern", ""))
        self._v_pat_h3.set(sec.get("h3_pattern", ""))
        self._v_pat_h4.set(sec.get("h4_pattern", ""))
        self._v_renum.set(sec.get("renumber_headings", True))
        # detect matching preset
        for name, preset in self.HEADING_PRESETS.items():
            if (preset["h1"] == sec.get("chapter_pattern") and
                    preset["h2"] == sec.get("h2_pattern") and
                    preset["h3"] == sec.get("h3_pattern") and
                    preset["h4"] == sec.get("h4_pattern")):
                self._v_hpreset.set(name)
                break
        # captions
        cap = cfg.get("captions", {})
        self._v_cap_fig.set(cap.get("figure_pattern", r"^图\s*\d"))
        self._v_cap_tbl.set(cap.get("table_pattern", r"^(续)?表\s*\d"))
        self._v_cap_sub.set(cap.get("subfigure_pattern", r"^\([a-z]\)"))
        self._v_cap_note.set(cap.get("note_pattern", r"^注[：:]"))
        self._v_cap_kwn.set(cap.get("keep_with_next", True))
        self._v_cap_chk.set(cap.get("check_numbering", True))
        # cover
        self._v_cov_en.set(cfg["cover"]["enabled"])
        self._v_school.set(cfg["meta"]["school_name"])
        self._v_logo.set(cfg["cover"]["logo"])
        self._v_covtitle.set(cfg["cover"]["title_text"])
        # cover fields
        while self._cov_field_rows:
            self._del_cov_field()
        for fld in cfg["cover"].get("fields", []):
            self._add_cov_field(fld.get("label", ""), fld.get("underline_chars", 33))
        # declarations
        decls = cfg.get("declarations", [])
        self._v_decl_en.set(len(decls) > 0)
        for i, dw in enumerate(self._decl_widgets):
            if i < len(decls):
                dw["title"].set(decls[i].get("title", ""))
                dw["body"].delete("1.0", "end")
                dw["body"].insert("1.0", decls[i].get("body", ""))
                dw["orig"] = copy.deepcopy(decls[i])
        # advanced
        self._v_tocd.set(cfg["toc"]["depth"])
        self._v_tocfont.set(cfg["toc"].get("font", cfg["fonts"]["body"]))
        self._v_tocsz.set(str(self._numval(cfg["toc"].get("font_size", cfg["sizes"]["body"]))) + "pt")
        self._v_toc_h1font.set(cfg["toc"].get("h1_font", cfg["fonts"]["h1"]))
        self._v_toc_h1sz.set(str(self._numval(cfg["toc"].get("h1_font_size", cfg["sizes"]["h1"]))) + "pt")
        self._v_tocls.set(cfg["toc"].get("line_spacing", cfg["body"]["line_spacing"]))
        self._v_toc_sb.set(cfg["toc"].get("space_before", 0))
        self._v_toc_sa.set(cfg["toc"].get("space_after", 0))
        self._v_refind.set(cfg["references"]["left_indent"])
        self._v_tbl_top.set(cfg["table"]["top_border_sz"] / 8)
        self._v_tbl_hdr.set(cfg["table"]["header_border_sz"] / 8)
        self._v_tbl_bot.set(cfg["table"]["bottom_border_sz"] / 8)
        self._v_tbl_ls.set(cfg["table"].get("line_spacing", 1.0))
        self._v_fn_ls.set(cfg["footnote"].get("line_spacing", 1.0))
        # front_matter
        fm = cfg.get("front_matter", {})
        self._v_fm_mode.set(self._FM_MODE_R.get(fm.get("mode", "auto"), "自动识别"))
        self._v_body_align.set(self._ALIGN_R.get(cfg["body"].get("align", "justify"), "两端对齐"))
        self._v_pgfmt_f.set(self._PGFMT_R.get(cfg["page_numbers"]["front_format"], "大写罗马"))
        self._v_pgfmt_b.set(self._PGFMT_R.get(cfg["page_numbers"]["body_format"], "阿拉伯数字"))
        # page_numbers position
        pn = cfg["page_numbers"]
        self._v_pn_fpos.set(self._PGPOS_R.get(pn.get("front_position", "center"), "居中"))
        self._v_pn_bpos.set(self._PGPOS_R.get(pn.get("body_position", "center"), "居中"))
        self._v_pn_fstart.set(pn.get("front_start", 1))
        self._v_pn_bstart.set(pn.get("body_start", 1))
        self._v_pn_deco.set(pn.get("decorator", "{page}"))
        self._v_pn_font.set(pn.get("font", ""))
        self._v_pn_bold.set(pn.get("bold", False))
        self._v_pn_size.set(str(self._numval(cfg["sizes"].get("page_number", 10.5))) + "pt")
        # header_footer
        hf = cfg.get("header_footer", {})
        self._v_hf_en.set(hf.get("enabled", False))
        self._v_hf_scope.set(self._HF_SCOPE_R.get(hf.get("scope", "body"), "仅正文"))
        self._v_hf_diff_oe.set(hf.get("different_odd_even", True))
        self._v_hf_first_no.set(hf.get("first_page_no_header", False))
        _odd_raw = hf.get("odd_page_text", "")
        _even_raw = hf.get("even_page_text", "")
        self._v_hf_odd_chap.set("{chapter_title}" in _odd_raw)
        self._v_hf_even_chap.set("{chapter_title}" in _even_raw)
        self._v_hf_odd_text.set("" if _odd_raw == "{chapter_title}" else _odd_raw)
        self._v_hf_even_text.set("" if _even_raw == "{chapter_title}" else _even_raw)
        self._v_hf_font.set(hf.get("font", "宋体"))
        self._v_hf_size.set(str(self._numval(hf.get("font_size", 10.5))) + "pt")
        self._v_hf_bold.set(hf.get("bold", False))
        self._v_hf_odd_align.set(self._ALIGN_R.get(hf.get("odd_page_align", "center"), "居中"))
        self._v_hf_even_align.set(self._ALIGN_R.get(hf.get("even_page_align", "center"), "居中"))
        self._v_hf_border.set(hf.get("border_bottom", True))
        self._v_hf_bwidth.set(hf.get("border_bottom_width", 0.75))
        self._v_hf_bstyle.set(self._BORDER_STYLE_R.get(hf.get("border_bottom_style", "single"), "单线"))
        # special titles
        while self._st_rows:
            self._del_st()
        for st in cfg.get("special_titles", []):
            self._add_st(st.get("match", ""), st.get("display", ""), st.get("align", "center"))

    # ---- save / load / reset ----

    def _save_config(self):
        path = self._filedialog.asksaveasfilename(
            title="保存配置文件", defaultextension=".yaml",
            filetypes=[("YAML 配置", "*.yaml *.yml")])
        if not path:
            return
        try:
            import yaml
        except ImportError:
            self._messagebox.showerror("错误", "需要 pyyaml 库。")
            return
        cfg = self._collect_config()
        with open(path, "w", encoding="utf-8") as f:
            yaml.dump(cfg, allow_unicode=True, default_flow_style=False,
                      sort_keys=False, stream=f)
        self._v_cfglbl.set(os.path.basename(path))

    def _load_config(self):
        path = self._filedialog.askopenfilename(
            title="加载配置文件",
            filetypes=[("YAML 配置", "*.yaml *.yml"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            from thesis_config import load_config
            cfg = load_config(path)
        except Exception as e:
            self._messagebox.showerror("错误", f"加载失败: {e}")
            return
        self._load_vars_from_config(cfg)
        self._v_cfglbl.set(os.path.basename(path))

    def _reset_defaults(self):
        self._load_vars_from_config(copy.deepcopy(DEFAULT_CONFIG))
        self._v_cfglbl.set("默认 (SCAU)")

    # ---- file dialogs ----

    def _browse_logo(self):
        path = self._filedialog.askopenfilename(
            title="选择 Logo 图片",
            filetypes=[("图片", "*.png *.jpg *.jpeg *.bmp"), ("所有文件", "*.*")])
        if path:
            self._v_logo.set(path)

    def _browse_custom_cover(self):
        path = self._filedialog.askopenfilename(
            title="选择封面 docx（单页）",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if path:
            self._v_custom_cover.set(path)

    def _browse_in(self):
        path = self._filedialog.askopenfilename(title="选择论文文件", filetypes=self.FILETYPES)
        if not path:
            return
        self._v_in.set(path)
        stem = os.path.splitext(os.path.basename(path))[0]
        self._v_out.set(os.path.join(os.path.dirname(path), f"{stem}_formatted.docx"))

    def _browse_out(self):
        path = self._filedialog.asksaveasfilename(
            title="保存输出文件", defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")])
        if path:
            self._v_out.set(path)

    # ---- logging ----

    def _append_log(self, text):
        self._msg_q.put(text)

    def _poll(self):
        while not self._msg_q.empty():
            msg = self._msg_q.get_nowait()
            self._log.config(state="normal")
            self._log.insert("end", msg + "\n")
            self._log.see("end")
            self._log.config(state="disabled")
        if self._running:
            self._root.after(100, self._poll)

    # ---- run ----

    def _start(self):
        inp = self._v_in.get().strip()
        out = self._v_out.get().strip()
        if not inp or not os.path.isfile(inp):
            self._messagebox.showerror("错误", "请选择有效的输入文件。")
            return
        if not out:
            self._messagebox.showerror("错误", "请指定输出文件路径。")
            return

        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

        self._btn.config(state="disabled")
        self._running = True
        self._root.after(100, self._poll)

        skip = self._v_skip.get()
        try:
            config = self._collect_config()
        except (ValueError, Exception) as e:
            self._messagebox.showerror("错误", f"参数值无效: {e}")
            self._btn.config(state="normal")
            self._running = False
            return

        def worker():
            try:
                ok = run_format(inp, out, skip, self._append_log, config=config)
                self._append_log("\n--- 格式化完成 ---" if ok else "\n--- 格式化失败 ---")
            except Exception as e:
                self._append_log(f"\n异常: {e}")
            finally:
                self._running = False
                self._root.after(0, lambda: self._btn.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Universal thesis formatter")
    parser.add_argument("--input", help="Input file (.docx/.doc/.txt/.md/.tex)")
    parser.add_argument("--output", help="Output docx (default: <stem>_formatted.docx)")
    parser.add_argument("--config", help="Path to thesis_config.yaml")
    parser.add_argument("--no-postprocess", action="store_true",
                        help="Skip Word COM post-processing")
    parser.add_argument("--dump-config", action="store_true",
                        help="Print default config YAML and exit")
    args = parser.parse_args()

    if args.dump_config:
        print(dump_default_config())
        return

    if not args.input:
        FormatterGUI()
        return

    input_path = os.path.abspath(args.input)
    if not os.path.isfile(input_path):
        print(f"Input not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Resolve config
    cfg, cfg_path = resolve_config(cli_config=args.config, input_path=input_path)

    stem = os.path.splitext(os.path.basename(input_path))[0]
    input_dir = os.path.dirname(input_path)
    output_path = (os.path.abspath(args.output) if args.output
                   else os.path.join(input_dir, f"{stem}_formatted.docx"))

    ok = run_format(input_path, output_path, args.no_postprocess, print,
                    config=cfg, config_path=cfg_path)
    if not ok:
        sys.exit(1)

    if not sys.stdin.isatty() or getattr(sys, "frozen", False):
        input("\n按回车键关闭...")


if __name__ == "__main__":
    main()
