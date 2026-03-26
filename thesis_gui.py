"""Tkinter GUI for the universal thesis formatter."""

import copy
import os
import queue
import threading

from thesis_config import DEFAULT_CONFIG
from thesis_runner import run_format
from thesis_formatter._common import (
    format_line_spacing_value,
    format_paragraph_spacing_value,
    normalize_length_text,
    normalize_line_spacing,
    normalize_line_spacing_mode,
)

# =============================================================================
# MODERN THEME COLOR SYSTEM (Fresh Modern + Swiss Minimal)
# Based on Education industry design guidelines
# =============================================================================
THEME = {
    "primary": "#7A6C61",
    "primary_hover": "#695C52",
    "success": "#7D8A72",
    "secondary": "#9A8C82",
    "accent": "#CDBFB5",
    "accent_soft": "#E8DED7",
    "bg_main": "#F2ECE7",
    "bg_surface": "#E7DDD6",
    "bg_card": "#FAF7F4",
    "bg_panel": "#F3ECE6",
    "border": "#C6B8AE",
    "text_primary": "#3F3834",
    "text_secondary": "#6D655F",
    "text_disabled": "#9D948D",
    "error": "#9A6E68",
    "warning": "#A18668",
    "info": "#8B959B",
}

try:
    import ttkbootstrap as ttkb
    from ttkbootstrap.constants import *
    HAS_TTKBOOTSTRAP = True
except ImportError:
    HAS_TTKBOOTSTRAP = False
    ttkb = None
class FormatterGUI:
    FILETYPES = [
        ("所有支持格式", "*.docx *.doc *.txt *.md *.tex"),
        ("Word 文档", "*.docx *.doc"),
        ("文本/Markdown", "*.txt *.md"),
        ("LaTeX", "*.tex"),
    ]
    CATEGORIES = ["页面", "封面声明", "正文", "标题", "页眉页码", "目录参考", "图表", "单独设置"]
    PANEL_META = {
        "页面": ("页面与前置页", "设置页边距、装订线、页眉页脚距离，以及封面和摘要等前置页处理方式。"),
        "正文": ("正文样式", "调整正文字体、字号、缩进、段落与脚注，让正文版式更稳定。"),
        "标题": ("标题层级", "设置各级标题样式、编号规则和识别模式，避免格式混乱。"),
        "页眉页码": ("页眉与页码", "配置页眉内容、奇偶页策略、页码格式和位置。"),
        "目录参考": ("目录与参考文献", "管理目录条目、参考文献缩进，以及特殊标题映射。"),
        "图表": ("图表与题注", "控制题注字体、识别规则和三线表参数。"),
        "单独设置": ("单独处理", "集中放置只做局部操作的功能，例如单独插入目录或只插入外部封面。"),
        "封面声明": ("封面与声明", "设置学校信息、自定义封面、封面字段和声明页内容。"),
    }
    PT_SIZES = ["9pt", "10.5pt", "12pt", "14pt", "16pt", "18pt", "22pt", "24pt", "26pt", "36pt"]
    ALIGN_LABELS = ["左对齐", "居中", "右对齐", "两端对齐"]
    ALIGN_LABELS_KEEP = ["保持原样", "左对齐", "居中", "右对齐", "两端对齐"]
    _ALIGN = {"左对齐": "left", "居中": "center", "右对齐": "right", "两端对齐": "justify", "保持原样": "keep"}
    _ALIGN_R = {v: k for k, v in _ALIGN.items()}
    BOLD_LABELS = ["加粗", "不加粗", "保持原样"]
    _BOLD = {"加粗": True, "不加粗": False, "保持原样": "keep"}
    _BOLD_R = {True: "加粗", False: "不加粗", "keep": "保持原样"}
    _FM_MODE = {
        "自动识别并格式化": "auto",
        "跳过（保留前置页原格式）": "skip",
        "强制格式化前置页": "format",
    }
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
    _CAPTION_MODE = {"严格动态模式 (推荐)": "dynamic", "稳定模式": "stable"}
    _CAPTION_MODE_R = {v: k for k, v in _CAPTION_MODE.items()}
    CAPTION_MODE_LABELS = list(_CAPTION_MODE.keys())
    _LINE_SPACING_MODE = {
        "单倍行距": {"mode": "multiple", "preset": 1.0},
        "1.5倍行距": {"mode": "multiple", "preset": 1.5},
        "2倍行距": {"mode": "multiple", "preset": 2.0},
        "最小值": {"mode": "at_least"},
        "固定值": {"mode": "exact"},
        "多倍行距": {"mode": "multiple"},
    }
    _LINE_SPACING_MODE_R = {"exact": "固定值", "at_least": "最小值", "multiple": "多倍行距"}
    LINE_SPACING_MODE_LABELS = list(_LINE_SPACING_MODE.keys())
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


    def __init__(self, theme="sandstone"):
        import tkinter as tk
        from tkinter import filedialog, messagebox, scrolledtext, ttk
    
        self._tk = tk
        self._filedialog = filedialog
        self._messagebox = messagebox
        self._scrolledtext = scrolledtext
    
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            pass
    
        theme = "sandstone"
        if HAS_TTKBOOTSTRAP and ttkb:
            self._root = ttkb.Window(themename=theme)
            self._ttk = ttkb
            self._style = ttkb.Style()
            self._theme = theme
            self._has_theme = True
        else:
            self._root = tk.Tk()
            self._ttk = ttk
            self._style = ttk.Style()
            self._theme = "default"
            self._has_theme = False
    
        self._root.title("论文排版工具")
        self._root.resizable(True, True)
        self._configure_window_bounds()
    
        self._msg_q = queue.Queue()
        self._running = False
    
        self._init_vars(tk)
        self._status_var = tk.StringVar(value="等待开始")
        self._header_input_var = tk.StringVar(value="未选择输入文件")
        self._header_output_var = tk.StringVar(value="未设置输出路径")
        self._active_scroll_canvas = None
        self._scroll_target_map = {}
    
        self._configure_modern_style()
    
        main_container = self._ttk.Frame(self._root, style="Workspace.TFrame")
        main_container.pack(fill="both", expand=True)
    
        self._build_header(main_container)
    
        self._main_pane = self._ttk.Panedwindow(main_container, orient="vertical")
        self._main_pane.pack(fill="both", expand=True, padx=0, pady=(0, 12))

        workspace_host = self._ttk.Frame(self._main_pane, style="Workspace.TFrame")
        bottom_host = self._ttk.Frame(self._main_pane, style="Workspace.TFrame")
        self._main_pane.add(workspace_host, weight=5)
        self._main_pane.add(bottom_host, weight=2)

        workspace = self._ttk.Frame(workspace_host, style="Workspace.TFrame")
        workspace.pack(fill="both", expand=True, padx=16, pady=(0, 12))
    
        sidebar_frame = self._ttk.Frame(workspace, style="Card.TFrame", padding=(18, 18, 18, 18))
        sidebar_frame.pack(side="left", fill="y", padx=(0, 12))
    
        sidebar_header = self._ttk.Frame(sidebar_frame, style="Card.TFrame")
        sidebar_header.pack(fill="x", pady=(0, 16))
        self._ttk.Label(sidebar_header, text="配置导航", style="SidebarTitle.TLabel").pack(anchor="w")
        self._ttk.Label(
            sidebar_header,
            text="按分组检查参数，完成后在底部执行格式化。",
            style="Hint.TLabel",
            wraplength=180,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))
    
        self._cat_buttons = {}
        self._sidebar_canvas = tk.Canvas(sidebar_frame, highlightthickness=0, borderwidth=0, bg=THEME["bg_card"], width=220)
        self._sidebar_scroll = self._ttk.Scrollbar(sidebar_frame, orient="vertical", command=self._sidebar_canvas.yview)
        self._sidebar_inner = self._ttk.Frame(self._sidebar_canvas, style="Card.TFrame")
        self._sidebar_inner.bind("<Configure>", lambda e: self._sidebar_canvas.configure(scrollregion=self._sidebar_canvas.bbox("all")))
        sidebar_window = self._sidebar_canvas.create_window((0, 0), window=self._sidebar_inner, anchor="nw")
        self._sidebar_canvas.bind("<Configure>", lambda e, wid=sidebar_window: self._sidebar_canvas.itemconfigure(wid, width=e.width))
        self._sidebar_canvas.configure(yscrollcommand=self._sidebar_scroll.set)
        self._sidebar_canvas.pack(side="left", fill="both", expand=True)
        self._sidebar_scroll.pack(side="right", fill="y", padx=(8, 0))

        self._register_scroll_target(sidebar_frame, self._sidebar_canvas)
        self._register_scroll_target(sidebar_header, self._sidebar_canvas)
        self._register_scroll_target(self._sidebar_canvas, self._sidebar_canvas)
        self._register_scroll_target(self._sidebar_inner, self._sidebar_canvas)

        for cat in self.CATEGORIES:
            btn = tk.Button(
                self._sidebar_inner,
                text=cat,
                font=("Microsoft YaHei UI", 10),
                relief="flat",
                borderwidth=0,
                bg=THEME["bg_card"],
                fg=THEME["text_primary"],
                activebackground=THEME["accent_soft"],
                activeforeground=THEME["text_primary"],
                cursor="hand2",
                anchor="w",
                padx=16,
                pady=10,
                highlightthickness=1,
                highlightbackground=THEME["border"],
                highlightcolor=THEME["border"],
                command=lambda c=cat: self._on_cat_click(c),
            )
            btn.pack(fill="x", pady=4)
            self._cat_buttons[cat] = btn
            self._register_scroll_target(btn, self._sidebar_canvas)
            btn.bind("<Enter>", lambda e, b=btn: self._on_btn_hover(b, True))
            btn.bind("<Leave>", lambda e, b=btn: self._on_btn_hover(b, False))
    
        content_frame = self._ttk.Frame(workspace, style="Card.TFrame")
        content_frame.pack(side="left", fill="both", expand=True)
    
        self._content = self._ttk.Frame(content_frame, style="Panel.TFrame")
        self._content.pack(fill="both", expand=True)
    
        self._panels = {}
        self._panel_canvas = {}
        for name, builder in [
            ("页面", self._build_page),
            ("封面声明", self._build_cover_decl),
            ("正文", self._build_body),
            ("标题", self._build_heading),
            ("页眉页码", self._build_header_pn),
            ("目录参考", self._build_toc_ref),
            ("图表", self._build_caption),
            ("单独设置", self._build_standalone),
        ]:
            wrapper = self._ttk.Frame(self._content, style="Panel.TFrame")
            canvas = tk.Canvas(wrapper, highlightthickness=0, borderwidth=0, bg=THEME["bg_card"])
            vsb = self._ttk.Scrollbar(wrapper, orient="vertical", command=canvas.yview)
            inner = self._ttk.Frame(canvas, padding=(28, 24, 28, 24), style="Panel.TFrame")
            inner.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
            window_id = canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.bind("<Configure>", lambda e, c=canvas, wid=window_id: c.itemconfigure(wid, width=e.width))
            canvas.configure(yscrollcommand=vsb.set)
            canvas.pack(side="left", fill="both", expand=True)
            vsb.pack(side="right", fill="y")
            self._register_scroll_target(canvas, canvas)
            self._register_scroll_target(inner, canvas)
            title, desc = self.PANEL_META.get(name, (name, ""))
            self._build_panel_intro(inner, title, desc)
            form = self._ttk.Frame(inner, style="Panel.TFrame")
            form.pack(fill="both", expand=True, pady=(18, 0))
            self._prepare_panel_grid(form)
            builder(form)
            self._panels[name] = wrapper
            self._panel_canvas[name] = canvas
    
        self._build_bottom(bottom_host)
    
        self._cur_panel = None
        self._show_panel("页面")
        self._update_cat_button_style("页面")
    
        self._v_in.trace_add("write", self._refresh_file_summary)
        self._v_out.trace_add("write", self._refresh_file_summary)
    
        self._load_vars_from_config(copy.deepcopy(DEFAULT_CONFIG))
        self._refresh_file_summary()
        self._root.bind_all("<MouseWheel>", self._on_global_mousewheel, add="+")
        self._root.bind_all("<Button-4>", self._on_global_mousewheel, add="+")
        self._root.bind_all("<Button-5>", self._on_global_mousewheel, add="+")
    
        self._root.mainloop()

    def _configure_window_bounds(self):
        screen_w = self._root.winfo_screenwidth()
        screen_h = self._root.winfo_screenheight()

        min_w = 920 if screen_w >= 1040 else max(820, screen_w - 80)
        min_h = 620 if screen_h >= 760 else max(560, screen_h - 80)
        width = min(1320, max(min_w, screen_w - 96))
        height = min(920, max(min_h, screen_h - 96))
        pos_x = max(0, (screen_w - width) // 2)
        pos_y = max(0, (screen_h - height) // 2)

        self._root.minsize(min_w, min_h)
        try:
            self._root.maxsize(screen_w, screen_h)
        except Exception:
            pass
        self._root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")

    def _set_initial_pane_layout(self):
        if not hasattr(self, "_main_pane"):
            return
        total_h = self._main_pane.winfo_height()
        if total_h <= 1:
            return
        bottom_h = min(280, max(210, total_h // 3))
        try:
            self._main_pane.sashpos(0, total_h - bottom_h)
        except Exception:
            pass

    def _register_scroll_target(self, widget, canvas):
        self._scroll_target_map[str(widget)] = canvas
        widget.bind("<Enter>", lambda _e, c=canvas: self._set_active_scroll_canvas(c), add="+")

    def _set_active_scroll_canvas(self, canvas):
        self._active_scroll_canvas = canvas

    def _resolve_scroll_canvas(self, widget):
        current = widget
        while current is not None:
            canvas = self._scroll_target_map.get(str(current))
            if canvas is not None and canvas.winfo_exists():
                return canvas
            try:
                parent_name = current.winfo_parent()
            except Exception:
                break
            if not parent_name:
                break
            try:
                current = current.nametowidget(parent_name)
            except Exception:
                break
        return None

    def _on_global_mousewheel(self, event):
        target_widget = None
        try:
            target_widget = self._root.winfo_containing(event.x_root, event.y_root)
        except Exception:
            target_widget = None
        if target_widget is None:
            target_widget = getattr(event, "widget", None)
        canvas = self._resolve_scroll_canvas(target_widget)
        if canvas is None or not canvas.winfo_exists():
            return
        self._set_active_scroll_canvas(canvas)
        if getattr(event, "delta", 0):
            delta = -1 if event.delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            delta = -1
        elif getattr(event, "num", None) == 5:
            delta = 1
        else:
            delta = 0
        if delta:
            canvas.yview_scroll(delta, "units")
            return "break"

    def _configure_modern_style(self):
        try:
            self._root.configure(bg=THEME["bg_main"])
        except Exception:
            pass

        style = self._style
        style.configure(".", font=("Microsoft YaHei UI", 9))
        style.configure("Workspace.TFrame", background=THEME["bg_main"])
        style.configure("Card.TFrame", background=THEME["bg_card"], relief="solid", borderwidth=1)
        style.configure("Panel.TFrame", background=THEME["bg_card"], relief="flat", borderwidth=0)
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 17, "bold"), foreground=THEME["text_primary"], background=THEME["bg_card"])
        style.configure("Subtitle.TLabel", font=("Microsoft YaHei UI", 9), foreground=THEME["text_secondary"], background=THEME["bg_card"])
        style.configure("SidebarTitle.TLabel", font=("Microsoft YaHei UI", 11, "bold"), foreground=THEME["text_primary"], background=THEME["bg_card"])
        style.configure("PanelTitle.TLabel", font=("Microsoft YaHei UI", 15, "bold"), foreground=THEME["primary"], background=THEME["bg_card"])
        style.configure("PanelDesc.TLabel", font=("Microsoft YaHei UI", 9), foreground=THEME["text_secondary"], background=THEME["bg_card"])
        style.configure("Meta.TLabel", font=("Microsoft YaHei UI", 9), foreground=THEME["text_secondary"], background=THEME["bg_card"])
        style.configure("Value.TLabel", font=("Microsoft YaHei UI", 9, "bold"), foreground=THEME["text_primary"], background=THEME["bg_card"])
        style.configure("Hint.TLabel", font=("Microsoft YaHei UI", 8), foreground=THEME["text_secondary"], background=THEME["bg_card"])
        style.configure("TLabel", background=THEME["bg_card"], foreground=THEME["text_primary"])
        style.configure("TLabelframe", background=THEME["bg_card"], borderwidth=1, relief="solid", bordercolor=THEME["border"])
        style.configure("TLabelframe.Label", background=THEME["bg_card"], foreground=THEME["text_primary"], font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("TCheckbutton", background=THEME["bg_card"], foreground=THEME["text_primary"])
        style.configure("TButton", padding=(12, 8), background=THEME["accent_soft"], foreground=THEME["text_primary"], borderwidth=1)
        style.map("TButton", background=[("active", THEME["accent"]), ("pressed", THEME["accent"])], foreground=[("disabled", THEME["text_disabled"])])
        style.configure("Primary.Tool.TButton", padding=(14, 9), background=THEME["primary"], foreground=THEME["bg_card"], borderwidth=1)
        style.map("Primary.Tool.TButton", background=[("active", THEME["primary_hover"]), ("pressed", THEME["primary_hover"]), ("disabled", THEME["secondary"])], foreground=[("disabled", THEME["bg_card"])])
        style.configure("Secondary.Tool.TButton", padding=(12, 8), background=THEME["accent_soft"], foreground=THEME["text_primary"], borderwidth=1)
        style.configure("TCombobox", fieldbackground=THEME["bg_card"], background=THEME["bg_card"], foreground=THEME["text_primary"], arrowcolor=THEME["text_primary"], bordercolor=THEME["border"])
        style.map("Secondary.Tool.TButton", background=[("active", THEME["accent"]), ("pressed", THEME["accent"]), ("disabled", THEME["bg_surface"])], foreground=[("disabled", THEME["text_disabled"])])
        style.configure("Tool.TCombobox", fieldbackground=THEME["bg_card"], background=THEME["bg_card"], foreground=THEME["text_primary"], arrowcolor=THEME["text_primary"], bordercolor=THEME["border"])
        style.configure("TEntry", fieldbackground=THEME["bg_card"], foreground=THEME["text_primary"], bordercolor=THEME["border"])
        style.configure("TSpinbox", fieldbackground=THEME["bg_card"], background=THEME["bg_card"], foreground=THEME["text_primary"], arrowcolor=THEME["text_primary"], bordercolor=THEME["border"])
        try:
            style.map("TCheckbutton", background=[("active", THEME["bg_card"])])
            style.map("TCombobox", fieldbackground=[("readonly", THEME["bg_card"])], selectbackground=[("readonly", THEME["accent_soft"])], selectforeground=[("readonly", THEME["text_primary"])])
            style.map("Tool.TCombobox", fieldbackground=[("readonly", THEME["bg_card"])], selectbackground=[("readonly", THEME["accent_soft"])], selectforeground=[("readonly", THEME["text_primary"])])
        except Exception:
            pass

    def _create_button(self, parent, bootstyle=None, **kwargs):
        if "style" not in kwargs:
            kwargs["style"] = "Primary.Tool.TButton" if bootstyle == "primary" else "Secondary.Tool.TButton"
        return self._ttk.Button(parent, **kwargs)

    def _prepare_panel_grid(self, panel):
        panel.grid_columnconfigure(0, minsize=180)
        panel.grid_columnconfigure(1, minsize=280, weight=1)
        panel.grid_columnconfigure(2, minsize=260, weight=1)

    def _build_panel_intro(self, parent, title, description):
        intro = self._ttk.Frame(parent, style="Panel.TFrame")
        intro.pack(fill="x")
        self._ttk.Label(intro, text=title, style="PanelTitle.TLabel").pack(anchor="w")
        self._ttk.Label(intro, text=description, style="PanelDesc.TLabel", wraplength=760, justify="left").pack(anchor="w", pady=(6, 0))


    def _on_btn_hover(self, btn, is_hovering):
        current_cat = self._cur_panel
        if is_hovering:
            if btn != self._cat_buttons.get(current_cat):
                btn.config(bg=THEME["accent_soft"], highlightbackground=THEME["accent"])
        else:
            if btn != self._cat_buttons.get(current_cat):
                btn.config(bg=THEME["bg_card"], highlightbackground=THEME["border"])

    def _on_cat_click(self, category):
        self._show_panel(category)
        self._update_cat_button_style(category)

    def _update_cat_button_style(self, active_cat):
        for cat, btn in self._cat_buttons.items():
            if cat == active_cat:
                btn.config(bg=THEME["primary"], fg=THEME["bg_card"], activebackground=THEME["primary_hover"], activeforeground=THEME["bg_card"], highlightbackground=THEME["primary"])
            else:
                btn.config(bg=THEME["bg_card"], fg=THEME["text_primary"], activebackground=THEME["accent_soft"], activeforeground=THEME["text_primary"], highlightbackground=THEME["border"])


    def _build_header(self, parent):
        header = self._ttk.Frame(parent, style="Card.TFrame", padding=(18, 12, 18, 10))
        header.pack(fill="x", padx=16, pady=(14, 10))

        top_row = self._ttk.Frame(header, style="Card.TFrame")
        top_row.pack(fill="x")
        title_block = self._ttk.Frame(top_row, style="Card.TFrame")
        title_block.pack(side="left", fill="x", expand=True)
        self._ttk.Label(title_block, text="论文排版工具", style="Title.TLabel").pack(anchor="w")
        self._ttk.Label(title_block, text="通用论文参数配置与格式化工作台", style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))


        summary = self._ttk.Frame(header, style="Card.TFrame")
        summary.pack(fill="x", pady=(10, 0))
        for col in (1, 3):
            summary.grid_columnconfigure(col, weight=1)

        def add_info_item(row, column, label, textvariable, emphasis=False):
            base_col = column * 2
            self._ttk.Label(summary, text=f"{label}：", style="Meta.TLabel").grid(
                row=row, column=base_col, sticky="w", pady=2
            )
            self._ttk.Label(
                summary,
                textvariable=textvariable,
                style="Value.TLabel" if emphasis else "Meta.TLabel",
                wraplength=300,
                justify="left",
            ).grid(row=row, column=base_col + 1, sticky="w", padx=(8, 16), pady=2)

        add_info_item(0, 0, "当前配置", self._v_cfglbl, emphasis=True)
        add_info_item(0, 1, "当前状态", self._status_var)
        add_info_item(1, 0, "输入文件", self._header_input_var)
        add_info_item(1, 1, "输出位置", self._header_output_var)
    
    def _quick_format(self):
            if not self._v_in.get():
                self._browse_in()
            if self._v_in.get():
                self._start()

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
        # heading spacing (支持多单位: pt/cm/mm/in/行)
        self._v_h1sb = tk.StringVar(value=format_paragraph_spacing_value(c["headings"]["h1"].get("space_before", 0)))
        self._v_h1sa = tk.StringVar(value=format_paragraph_spacing_value(c["headings"]["h1"].get("space_after", 0)))
        self._v_h2sb = tk.StringVar(value=format_paragraph_spacing_value(c["headings"]["h2"].get("space_before", 0)))
        self._v_h2sa = tk.StringVar(value=format_paragraph_spacing_value(c["headings"]["h2"].get("space_after", 0)))
        self._v_h3sb = tk.StringVar(value=format_paragraph_spacing_value(c["headings"]["h3"].get("space_before", 0)))
        self._v_h3sa = tk.StringVar(value=format_paragraph_spacing_value(c["headings"]["h3"].get("space_after", 0)))
        self._v_h4sb = tk.StringVar(value=format_paragraph_spacing_value(c["headings"]["h4"].get("space_before", 0)))
        self._v_h4sa = tk.StringVar(value=format_paragraph_spacing_value(c["headings"]["h4"].get("space_after", 0)))
        body_ls_mode, body_ls_value = self._split_line_spacing_for_gui(c["body"]["line_spacing"])
        self._v_lsp_mode = tk.StringVar(value=body_ls_mode)
        self._v_lsp = tk.StringVar(value=body_ls_value)
        self._v_ind = tk.DoubleVar(value=c["body"]["first_line_indent"])
        self._v_body_sb = tk.StringVar(value=format_paragraph_spacing_value(c["body"].get("space_before", 0)))
        self._v_body_sa = tk.StringVar(value=format_paragraph_spacing_value(c["body"].get("space_after", 0)))
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
        self._v_cap_mode = tk.StringVar(value=self._CAPTION_MODE_R.get(cap.get("mode", "dynamic"), "严格动态模式 (推荐)"))
        self._v_cap_fig = tk.StringVar(value=cap.get("figure_pattern", r"^图\s*\d"))
        self._v_cap_tbl = tk.StringVar(value=cap.get("table_pattern", r"^(续)?表\s*\d"))
        self._v_cap_sub = tk.StringVar(value=cap.get("subfigure_pattern", r"^\([a-z]\)"))
        self._v_cap_note = tk.StringVar(value=cap.get("note_pattern", r"^注[：:]"))
        self._v_cap_kwn = tk.BooleanVar(value=cap.get("keep_with_next", True))
        self._v_cap_chk = tk.BooleanVar(value=cap.get("check_numbering", True))
        self._v_cap_include_chapter = tk.BooleanVar(value=cap.get("include_chapter", False))
        self._v_cap_restart_chapter = tk.BooleanVar(value=cap.get("restart_per_chapter", False))
        self._v_cap_heading_level = tk.StringVar(value=str(cap.get("chapter_heading_level", 1)))
        self._v_cap_chapter_sep = tk.StringVar(value=cap.get("chapter_separator", "."))
        self._v_cap_caption_sep = tk.StringVar(value=cap.get("caption_separator", ""))
        cap_ls_mode, cap_ls_value = self._split_line_spacing_for_gui(cap.get("line_spacing", c["body"]["line_spacing"]))
        self._v_cap_ls_mode = tk.StringVar(value=cap_ls_mode)
        self._v_cap_ls = tk.StringVar(value=cap_ls_value)
        self._v_cap_font = tk.StringVar(value=cap.get("font", "宋体"))
        self._v_cap_numfont = tk.StringVar(value=cap.get("number_font", "Times New Roman"))
        # cover
        self._v_cov_en = tk.BooleanVar(value=c["cover"]["enabled"])
        self._v_cover_only = tk.BooleanVar(value=c["cover"].get("only_insert", False))
        self._v_pn_only = tk.BooleanVar(value=c["page_numbers"].get("only_insert", False))
        self._v_hf_only = tk.BooleanVar(value=c["header_footer"].get("only_insert", False))
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
        self._v_toc_bold = tk.BooleanVar(value=c["toc"].get("bold", False))
        self._v_toc_only = tk.BooleanVar(value=c["toc"].get("only_insert", False))
        self._v_toc_ex_abs = tk.BooleanVar(value=c["toc"].get("exclude_abstract_headings", False))
        toc_ls_mode, toc_ls_value = self._split_line_spacing_for_gui(c["toc"].get("line_spacing", c["body"]["line_spacing"]))
        self._v_tocls_mode = tk.StringVar(value=toc_ls_mode)
        self._v_tocls = tk.StringVar(value=toc_ls_value)
        self._v_toc_h1font = tk.StringVar(value=c["toc"].get("h1_font", c["fonts"]["h1"]))
        self._v_toc_h1sz = tk.StringVar(value=str(c["toc"].get("h1_font_size", c["sizes"]["h1"])) + "pt")
        self._v_toc_h1bold = tk.BooleanVar(value=c["toc"].get("h1_bold", False))
        self._v_toc_sb = tk.StringVar(value=format_paragraph_spacing_value(c["toc"].get("space_before", 0)))
        self._v_toc_sa = tk.StringVar(value=format_paragraph_spacing_value(c["toc"].get("space_after", 0)))
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
        tbl_ls_mode, tbl_ls_value = self._split_line_spacing_for_gui(c["table"]["line_spacing"])
        self._v_tbl_ls_mode = tk.StringVar(value=tbl_ls_mode)
        self._v_tbl_ls = tk.StringVar(value=tbl_ls_value)
        fn_ls_mode, fn_ls_value = self._split_line_spacing_for_gui(c["footnote"]["line_spacing"])
        self._v_fn_ls_mode = tk.StringVar(value=fn_ls_mode)
        self._v_fn_ls = tk.StringVar(value=fn_ls_value)
        # front_matter
        fm = c.get("front_matter", {})
        self._v_fm_mode = tk.StringVar(
            value=self._FM_MODE_R.get(fm.get("mode", "auto"), "自动识别并格式化")
        )
        # file I/O
        self._v_in = tk.StringVar()
        self._v_out = tk.StringVar()
        self._v_skip = tk.BooleanVar(value=not c["toc"].get("enabled", True))
        self._v_cfglbl = tk.StringVar(value="默认 (SCAU)")

    # ---- row helpers ----

    def _row_spin(self, p, r, lbl, var, lo=0.0, hi=100.0, step=0.1, unit="cm"):
        """带单位的数值输入 + 上下调节按钮 - 现代化样式"""
        import re

        self._ttk.Label(p, text=lbl, font=("Microsoft YaHei UI", 9),
                       foreground=THEME["text_primary"]).grid(row=r, column=0, sticky="w", pady=6)

        # 容器
        cf = self._ttk.Frame(p)
        cf.grid(row=r, column=1, sticky="w", padx=4, pady=3)

        # 解析函数
        def parse_value(v):
            if not v:
                return lo, unit
            v = str(v).strip()

            # "磅" 是 pt 的中文表达
            if "磅" in v:
                v = v.replace("磅", "pt")

            # 检查中文字号（转换为 pt）
            cn_pt_map = {"小四": 12, "四号": 14, "五号": 10.5, "小五": 9}
            if v in cn_pt_map:
                return cn_pt_map[v], "pt"

            # 解析数字+单位
            m = re.match(r"^([\d.]+)([a-zA-Z\u884c]*)$", v)
            if m:
                num_str, u = m.groups()
                try:
                    num = float(num_str)
                    parsed_unit = u or unit
                    return num, parsed_unit
                except:
                    return lo, unit
            try:
                return float(v), unit
            except:
                return lo, unit

        # 格式化函数
        def format_value(num, u):
            # 处理浮点数精度问题：保留合理的小数位数
            if isinstance(num, float):
                # 如果接近整数，显示为整数
                if abs(num - round(num)) < 0.0001:
                    num = round(num)
                # 否则保留最多2位小数
                else:
                    num = round(num, 2)
                    # 去掉末尾无意义的0
                    if num == int(num):
                        num = int(num)
            elif num == int(num):
                num = int(num)
            if u == "行":
                return f"{num}行"
            return f"{num}{u}"

        def set_var_value(num, u):
            if isinstance(var, self._tk.IntVar):
                var.set(int(round(num)))
            elif isinstance(var, self._tk.DoubleVar):
                var.set(float(num))
            else:
                var.set(format_value(num, u))

        # 初始化
        init_num, init_unit = parse_value(var.get())
        current_num = [init_num]
        current_unit = [init_unit]

        # 创建输入框（不绑定 textvariable，手动管理）
        entry = self._ttk.Entry(cf, width=12)
        entry.insert(0, format_value(init_num, init_unit))
        entry.pack(side="left")

        # 按钮点击标志（防止 FocusOut 竞争）
        button_clicking = [False]

        # 当输入框失去焦点或按回车时，解析并保存
        def on_input_complete(event=None):
            if button_clicking[0]:
                return  # 如果正在点击按钮，跳过
            val = entry.get()
            num, u = parse_value(val)
            current_num[0] = num
            current_unit[0] = u
            entry.delete(0, "end")
            entry.insert(0, format_value(num, u))
            set_var_value(num, u)

        entry.bind("<Return>", on_input_complete)
        entry.bind("<FocusOut>", on_input_complete)

        # 上下按钮
        bf = self._ttk.Frame(cf)
        bf.pack(side="left", padx=(2, 0))

        def on_step(delta):
            button_clicking[0] = True
            new_num = current_num[0] + delta
            if lo <= new_num <= hi:
                current_num[0] = new_num
                entry.delete(0, "end")
                entry.insert(0, format_value(new_num, current_unit[0]))
                set_var_value(new_num, current_unit[0])
            # 延迟重置标志，确保 FocusOut 已经完成
            entry.after(100, lambda: button_clicking.__setitem__(0, False))

        # 创建上下按钮 - 使用更简洁的样式
        btn_up = self._create_button(bf, text="增", width=4, command=lambda: on_step(step), bootstyle="secondary-outline")
        btn_up.pack(side="left")
        btn_down = self._create_button(bf, text="减", width=4, command=lambda: on_step(-step), bootstyle="secondary-outline")
        btn_down.pack(side="left")

        # 支持键盘上下键
        def on_key(event):
            if event.keysym == "Up":
                on_step(step)
                return "break"
            elif event.keysym == "Down":
                on_step(-step)
                return "break"
        entry.bind("<Up>", on_key)
        entry.bind("<Down>", on_key)

        return r + 1

    def _row_unit_entry(self, p, r, lbl, var, default_unit="pt", lo=5, hi=72, step=0.5):
        """带单位的输入框 + 上下调节按钮（用于字号）"""
        import re

        self._ttk.Label(p, text=lbl).grid(row=r, column=0, sticky="w", pady=3)

        # 容器
        cf = self._ttk.Frame(p)
        cf.grid(row=r, column=1, sticky="w", padx=4, pady=3)

        # 中文字号映射
        cn_map = {"初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
                  "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
                  "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5}

        # 解析函数（支持"磅"作为pt的同义词）
        def parse_value(v):
            if not v:
                return 12, default_unit
            v = str(v).strip()
            # "磅" 是 pt 的中文表达
            if "磅" in v:
                v = v.replace("磅", "pt")
            # 检查中文字号（支持"小三"和"小三号"两种写法）
            if v in cn_map:
                return cn_map[v], "pt"
            # 去掉末尾的"号"字再检查
            if v.endswith("号"):
                v_without_hao = v[:-1]
                if v_without_hao in cn_map:
                    return cn_map[v_without_hao], "pt"
            m = re.match(r"^([\d.]+)([a-zA-Z]*)$", v)
            if m:
                num_str, u = m.groups()
                try:
                    return float(num_str), u or default_unit
                except:
                    return 12, default_unit
            try:
                return float(v), default_unit
            except:
                return 12, default_unit

        # 格式化函数
        def format_value(num, u):
            # 处理浮点数精度问题
            if isinstance(num, float):
                if abs(num - round(num)) < 0.0001:
                    num = round(num)
                else:
                    num = round(num, 2)
                    if num == int(num):
                        num = int(num)
            elif num == int(num):
                num = int(num)
            return f"{num}{u}"

        # 初始化
        init_num, init_unit = parse_value(var.get())
        current_num = [init_num]
        current_unit = [init_unit]

        # 创建输入框（不绑定 textvariable，手动管理）
        entry = self._ttk.Entry(cf, width=12)
        entry.insert(0, format_value(init_num, init_unit))
        entry.pack(side="left")

        # 按钮点击标志（防止 FocusOut 竞争）
        button_clicking = [False]

        # 当输入框失去焦点或按回车时，解析并保存
        def on_input_complete(event=None):
            if button_clicking[0]:
                return  # 如果正在点击按钮，跳过
            val = entry.get()
            num, u = parse_value(val)
            current_num[0] = num
            current_unit[0] = u
            entry.delete(0, "end")
            entry.insert(0, format_value(num, u))
            var.set(format_value(num, u))

        entry.bind("<Return>", on_input_complete)
        entry.bind("<FocusOut>", on_input_complete)

        # 上下按钮
        bf = self._ttk.Frame(cf)
        bf.pack(side="left", padx=(2, 0))

        def on_step(delta):
            button_clicking[0] = True
            new_num = current_num[0] + delta
            if lo <= new_num <= hi:
                current_num[0] = new_num
                entry.delete(0, "end")
                entry.insert(0, format_value(new_num, current_unit[0]))
                var.set(format_value(new_num, current_unit[0]))
            # 延迟重置标志，确保 FocusOut 已经完成
            entry.after(100, lambda: button_clicking.__setitem__(0, False))

        btn_up = self._create_button(bf, text="增", width=4, command=lambda: on_step(step), bootstyle="secondary-outline")
        btn_up.pack(side="left")
        btn_down = self._create_button(bf, text="减", width=4, command=lambda: on_step(-step), bootstyle="secondary-outline")
        btn_down.pack(side="left")

        # 支持键盘上下键
        def on_key(event):
            if event.keysym == "Up":
                on_step(step)
                return "break"
            elif event.keysym == "Down":
                on_step(-step)
                return "break"
        entry.bind("<Up>", on_key)
        entry.bind("<Down>", on_key)

        return r + 1

    def _row_line_spacing(self, p, r, lbl, mode_var, value_var, hint=None):
        self._ttk.Label(p, text=lbl, font=("Microsoft YaHei UI", 9),
                       foreground=THEME["text_primary"]).grid(row=r, column=0, sticky="w", pady=6)

        cf = self._ttk.Frame(p)
        cf.grid(row=r, column=1, sticky="w", padx=8, pady=6)

        combo = self._ttk.Combobox(
            cf,
            textvariable=mode_var,
            values=self.LINE_SPACING_MODE_LABELS,
            font=("Microsoft YaHei UI", 9),
            width=10,
            state="readonly",
        )
        combo.pack(side="left")

        entry = self._ttk.Entry(cf, textvariable=value_var, width=14, font=("Microsoft YaHei UI", 9))
        entry.pack(side="left", padx=(6, 0), ipady=4)

        def normalize_current(_event=None):
            choice = self._resolve_line_spacing_choice(mode_var.get())
            preset = choice.get("preset")
            if preset is not None:
                value_var.set(format_line_spacing_value({"mode": "multiple", "value": preset}))
                entry.configure(state="disabled")
                return

            entry.configure(state="normal")
            value_var.set(
                self._normalize_line_spacing_value(
                    choice.get("mode", "multiple"),
                    value_var.get(),
                )
            )

        combo.bind("<<ComboboxSelected>>", normalize_current)
        entry.bind("<Return>", normalize_current)
        entry.bind("<FocusOut>", normalize_current)
        normalize_current()

        if hint:
            self._ttk.Label(
                p,
                text=hint,
                foreground=THEME["text_secondary"],
                font=("Microsoft YaHei UI", 8),
                justify="left",
            ).grid(row=r, column=2, sticky="w", pady=6)

        return r + 1

    def _row_entry(self, p, r, lbl, var, w=28, hint=None):
        self._ttk.Label(p, text=lbl, font=("Microsoft YaHei UI", 9),
                       foreground=THEME["text_primary"]).grid(row=r, column=0, sticky="w", pady=6)
        if hint:
            self._ttk.Entry(p, textvariable=var, width=24, font=("Microsoft YaHei UI", 9)).grid(
                row=r, column=1, sticky="ew", padx=8, pady=6, ipady=4)
            self._ttk.Label(p, text=hint, foreground=THEME["text_secondary"], font=("Microsoft YaHei UI", 8), justify="left").grid(
                row=r, column=2, sticky="w", pady=6)
        else:
            self._ttk.Entry(p, textvariable=var, width=w, font=("Microsoft YaHei UI", 9)).grid(
                row=r, column=1, columnspan=2, sticky="ew", padx=8, pady=6, ipady=4)
        return r + 1

    def _row_combo(self, p, r, lbl, var, vals, w=10):
        self._ttk.Label(p, text=lbl, font=("Microsoft YaHei UI", 9),
                       foreground=THEME["text_primary"]).grid(row=r, column=0, sticky="w", pady=6)
        self._ttk.Combobox(p, textvariable=var, values=vals, font=("Microsoft YaHei UI", 9),
                           width=w, state="readonly").grid(row=r, column=1, sticky="w", padx=8, pady=6, ipady=4)
        return r + 1

    def _row_check(self, p, r, lbl, var, command=None):
        self._ttk.Checkbutton(p, text=lbl, variable=var, command=command).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=6)
        return r + 1

    def _sep(self, p, r):
        self._ttk.Separator(p, orient="horizontal").grid(
            row=r, column=0, columnspan=3, sticky="ew", pady=14)
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
        r = self._row_combo(p, r, "前置页处理:", self._v_fm_mode, list(self._FM_MODE.keys()), w=24)
        self._ttk.Label(
            p, text="前置页包括封面、声明和中英文摘要；自动识别会尝试格式化，跳过才会完整保留原格式",
            foreground=THEME["text_secondary"], font=("Microsoft YaHei UI", 8)
        ).grid(row=r, column=0, columnspan=3, sticky="w", padx=18)
        r += 1

    def _build_header_pn(self, p):
        r = 0
        # -- 页眉 --
        self._ttk.Label(p, text="页眉", font=("Microsoft YaHei UI", 11, "bold"),
                       foreground=THEME["primary"]).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(8, 6))
        r += 1
        r = self._row_check(p, r, "启用页眉", self._v_hf_en)
        r = self._row_combo(p, r, "作用范围:", self._v_hf_scope, list(self._HF_SCOPE.keys()))
        r = self._row_check(p, r, "奇偶页不同", self._v_hf_diff_oe)
        r = self._row_check(p, r, "首页不显示页眉", self._v_hf_first_no)
        # odd page (right side)
        r = self._row_entry(p, r, "奇数页(右):", self._v_hf_odd_text)
        r = self._row_check(p, r, "自动显示章标题", self._v_hf_odd_chap)
        r = self._row_combo(p, r, "奇数页对齐:", self._v_hf_odd_align, self.ALIGN_LABELS)
        # even page (left side)
        r = self._row_entry(p, r, "偶数页(左):", self._v_hf_even_text)
        r = self._row_check(p, r, "自动显示章标题", self._v_hf_even_chap)
        r = self._row_combo(p, r, "偶数页对齐:", self._v_hf_even_align, self.ALIGN_LABELS)
        r = self._row_entry(p, r, "页眉字体:", self._v_hf_font)
        r = self._row_unit_entry(p, r, "页眉字号:", self._v_hf_size, default_unit="pt", lo=5, hi=36)
        r = self._row_check(p, r, "页眉文字加粗", self._v_hf_bold)
        r = self._row_check(p, r, "页眉下划线", self._v_hf_border)
        r = self._row_spin(p, r, "下划线粗细:", self._v_hf_bwidth, lo=0.25, hi=3.0, step=0.25, unit="磅")
        r = self._row_combo(p, r, "下划线样式:", self._v_hf_bstyle, list(self._BORDER_STYLE.keys()))
        r = self._sep(p, r)
        # -- 页码 --
        self._ttk.Label(p, text="页码", font=("Microsoft YaHei UI", 11, "bold"),
                       foreground=THEME["primary"]).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(8, 6))
        r += 1
        r = self._row_combo(p, r, "前置页码位置:", self._v_pn_fpos, self.PGPOS_LABELS_SIMPLE)
        r = self._row_combo(p, r, "正文页码位置:", self._v_pn_bpos, self.PGPOS_LABELS)
        r = self._row_combo(p, r, "前置页码格式:", self._v_pgfmt_f, self.PGFMT_LABELS)
        r = self._row_combo(p, r, "正文页码格式:", self._v_pgfmt_b, self.PGFMT_LABELS)
        r = self._row_spin(p, r, "前置起始编号:", self._v_pn_fstart, lo=1, hi=999, step=1, unit="")
        r = self._row_spin(p, r, "正文起始编号:", self._v_pn_bstart, lo=1, hi=999, step=1, unit="")
        r = self._row_entry(p, r, "页码修饰:", self._v_pn_deco, hint="如: - {page} -")
        r = self._row_entry(p, r, "页码字体:", self._v_pn_font, hint="空=跟随西文")
        r = self._row_unit_entry(p, r, "页码字号:", self._v_pn_size, default_unit="pt", lo=5, hi=36)
        r = self._row_check(p, r, "页码加粗", self._v_pn_bold)

    def _build_body(self, p):
        r = 0
        r = self._row_entry(p, r, "西文字体:", self._v_flat)
        r = self._row_entry(p, r, "正文中文字体:", self._v_fbody)
        r = self._row_unit_entry(p, r, "正文字号:", self._v_sbody, default_unit="pt", lo=5, hi=36)
        r = self._row_combo(p, r, "正文对齐:", self._v_body_align, self.ALIGN_LABELS_KEEP)
        r = self._row_spin(p, r, "首行缩进:", self._v_ind, lo=0, hi=100, step=1, unit="pt")
        r = self._row_line_spacing(p, r, "行距:", self._v_lsp_mode, self._v_lsp, hint="可直接选单倍/1.5倍/2倍；多倍如 1.2；固定值/最小值如 20pt")
        r = self._row_spin(p, r, "段前:", self._v_body_sb, lo=0, hi=5, step=0.5, unit="行")
        r = self._row_spin(p, r, "段后:", self._v_body_sa, lo=0, hi=5, step=0.5, unit="行")
        r = self._sep(p, r)
        r = self._row_unit_entry(p, r, "脚注字号:", self._v_sfn, default_unit="pt", lo=5, hi=36)
        r = self._row_line_spacing(p, r, "脚注行距:", self._v_fn_ls_mode, self._v_fn_ls, hint="可直接选单倍/1.5倍/2倍；多倍如 1.25；固定值/最小值如 18pt")

    def _build_heading(self, p):
        r = 0
        sz = self.PT_SIZES

        def _heading_block(r, label, v_font, v_size, v_bold, v_align, v_sb, v_sa):
            self._ttk.Label(p, text=label, font=("Microsoft YaHei UI", 10, "bold"),
                           foreground=THEME["primary"]).grid(
                row=r, column=0, columnspan=3, sticky="w", pady=(8, 4))
            r += 1
            r = self._row_entry(p, r, "  字体:", v_font)
            r = self._row_unit_entry(p, r, "  字号:", v_size, default_unit="pt", lo=5, hi=72)
            r = self._row_combo(p, r, "  加粗:", v_bold, self.BOLD_LABELS)
            r = self._row_combo(p, r, "  对齐:", v_align, self.ALIGN_LABELS_KEEP)
            r = self._row_spin(p, r, "  段前:", v_sb, lo=-1, hi=5, step=0.5, unit="行")
            r = self._row_spin(p, r, "  段后:", v_sa, lo=-1, hi=5, step=0.5, unit="行")
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

        self._ttk.Label(p, text="(-1 = 保持原样)", foreground=THEME["text_secondary"]).grid(
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

    def _build_standalone(self, p):
        r = 0
        self._ttk.Label(p, text="单独处理", font=("Microsoft YaHei UI", 11, "bold"),
                       foreground=THEME["primary"]).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(8, 6))
        r += 1
        r = self._row_check(p, r, "单独插入/更新目录（不改其他格式）", self._v_toc_only, command=self._on_toc_only_toggle)
        r = self._row_check(p, r, "单独插入外部封面（不改正文与其他格式）", self._v_cover_only, command=self._on_cover_only_toggle)
        self._ttk.Label(
            p,
            text="使用“单独插入外部封面”时，要先在“封面声明”里选好自定义封面 .docx。",
            foreground=THEME["text_secondary"],
            font=("Microsoft YaHei UI", 8),
            justify="left",
        ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 6))
        r += 1
        r = self._row_check(p, r, "单独改页码", self._v_pn_only, command=self._on_pn_only_toggle)
        self._ttk.Label(
            p,
            text="只按现有分节调整页码，不自动补分节，尽量少动正文。",
            foreground=THEME["text_secondary"],
            font=("Microsoft YaHei UI", 8),
            justify="left",
        ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 6))
        r += 1
        r = self._row_check(p, r, "单独改页眉", self._v_hf_only, command=self._on_hf_only_toggle)
        self._ttk.Label(
            p,
            text="只按现有分节调整页眉，不自动补分节，尽量少动正文。",
            foreground=THEME["text_secondary"],
            font=("Microsoft YaHei UI", 8),
            justify="left",
        ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 6))
        r += 1
        self._ttk.Label(
            p,
            text="以上单独处理功能互斥，勾选其一时会自动取消其他选项。",
            foreground=THEME["text_secondary"],
            font=("Microsoft YaHei UI", 8),
            justify="left",
        ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 6))
    def _build_caption(self, p):
        r = 0
        r = self._row_unit_entry(p, r, "图表题字号:", self._v_scap, default_unit="pt", lo=5, hi=36)
        r = self._row_line_spacing(p, r, "题注行距:", self._v_cap_ls_mode, self._v_cap_ls, hint="可直接选单倍/1.5倍/2倍；多倍如 1.2；固定值/最小值如 16pt")
        r = self._row_entry(p, r, "题注字体:", self._v_cap_font, hint="(如: 宋体)")
        r = self._row_entry(p, r, "编号字体:", self._v_cap_numfont, hint="(如: Times New Roman)")
        r = self._row_combo(p, r, "题注模式:", self._v_cap_mode, self.CAPTION_MODE_LABELS, w=18)
        note = self._ttk.Label(
            p,
            text="dynamic 仅适用于真正 Heading 样式 + Word 多级列表的规范文档；预检不通过会自动回退 stable。",
            font=("Microsoft YaHei UI", 8),
            foreground=THEME["text_secondary"],
        )
        note.grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 4))
        r += 1
        r = self._row_check(p, r, "图表题防分页 (keep with next)", self._v_cap_kwn)
        r = self._row_check(p, r, "检查图表编号连续性", self._v_cap_chk)
        r = self._row_check(p, r, "图表按章编号（如 图2.1）", self._v_cap_include_chapter)
        r = self._row_check(p, r, "每章重新编号", self._v_cap_restart_chapter)
        r = self._row_combo(p, r, "章节标题级别:", self._v_cap_heading_level, ["1", "2", "3", "4"], w=6)
        r = self._row_entry(p, r, "章号分隔符:", self._v_cap_chapter_sep, hint="(如: . 或 -)")
        r = self._row_entry(p, r, "编号后分隔符:", self._v_cap_caption_sep, hint="(默认留空)")
        r = self._sep(p, r)
        # 题注识别模式说明
        note = self._ttk.Label(p, text="以下为题注识别模式（正则表达式），用于识别原文中的图表题。\n"
                                    "Word 文档中的题注必须符合这些格式才能被正确识别。一般无需修改。",
                              font=("Microsoft YaHei UI", 8), foreground=THEME["text_secondary"])
        note.grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 8))
        r += 1
        r = self._row_entry(p, r, "图题模式:", self._v_cap_fig, hint="(如: 图1)")
        r = self._row_entry(p, r, "表题模式:", self._v_cap_tbl, hint="(如: 表1)")
        r = self._row_entry(p, r, "分图模式:", self._v_cap_sub, hint="(如: (a))")
        r = self._row_entry(p, r, "表注模式:", self._v_cap_note, hint="(如: 注：)")
        r = self._sep(p, r)
        self._ttk.Label(p, text="三线表", font=("Microsoft YaHei UI", 11, "bold"),
                       foreground=THEME["primary"]).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(8, 6))
        r += 1
        r = self._row_spin(p, r, "顶线粗细:", self._v_tbl_top, lo=0.25, hi=6, step=0.25, unit="磅")
        r = self._row_spin(p, r, "栏目线粗细:", self._v_tbl_hdr, lo=0.25, hi=6, step=0.25, unit="磅")
        r = self._row_spin(p, r, "底线粗细:", self._v_tbl_bot, lo=0.25, hi=6, step=0.25, unit="磅")
        r = self._row_line_spacing(p, r, "表格行距:", self._v_tbl_ls_mode, self._v_tbl_ls, hint="可直接选单倍/1.5倍/2倍；多倍如 1.2；固定值/最小值如 14pt")

    def _build_cover_decl(self, p):
        # -- 封面 --
        r = 0
        r = self._row_check(p, r, "启用封面", self._v_cov_en)
        self._ttk.Label(p, text="自定义封面:").grid(row=r, column=0, sticky="w", pady=3)
        cf = self._ttk.Frame(p)
        cf.grid(row=r, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        self._ttk.Entry(cf, textvariable=self._v_custom_cover, width=22).pack(side="left")
        self._create_button(cf, text="浏览", width=5,
                            command=self._browse_custom_cover, bootstyle="secondary-outline").pack(side="left", padx=4)
        r += 1
        self._ttk.Label(p, text="（上传已排好版的封面页 .docx，将替代自动生成封面）",
                        foreground=THEME["text_secondary"]).grid(row=r, column=0, columnspan=3, sticky="w", pady=0)
        r += 1
        r = self._sep(p, r)
        r = self._row_entry(p, r, "学校名称:", self._v_school)
        self._ttk.Label(p, text="Logo 文件:").grid(row=r, column=0, sticky="w", pady=3)
        lf = self._ttk.Frame(p)
        lf.grid(row=r, column=1, columnspan=2, sticky="w", padx=4, pady=3)
        self._ttk.Entry(lf, textvariable=self._v_logo, width=22).pack(side="left")
        self._create_button(lf, text="浏览", width=5, command=self._browse_logo, bootstyle="secondary-outline").pack(side="left", padx=4)
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
        self._create_button(bf, text="添加", width=6, command=self._add_cov_field, bootstyle="secondary-outline").pack(side="left")
        self._create_button(bf, text="删除末行", width=8, command=self._del_cov_field, bootstyle="secondary-outline").pack(side="left", padx=4)
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
        self._ttk.Label(p, text="目录", font=("Microsoft YaHei UI", 11, "bold"),
                       foreground=THEME["primary"]).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(8, 6))
        r += 1
        r = self._row_spin(p, r, "目录深度:", self._v_tocd, lo=1, hi=4, step=1, unit="级")
        r = self._row_entry(p, r, "一级条目字体:", self._v_toc_h1font)
        r = self._row_unit_entry(p, r, "一级条目字号:", self._v_toc_h1sz, default_unit="pt", lo=5, hi=72)
        r = self._row_check(p, r, "一级条目加粗", self._v_toc_h1bold)
        r = self._row_entry(p, r, "二级条目字体:", self._v_tocfont)
        r = self._row_unit_entry(p, r, "二级条目字号:", self._v_tocsz, default_unit="pt", lo=5, hi=36)
        r = self._row_check(p, r, "二级条目加粗", self._v_toc_bold)
        self._ttk.Label(
            p,
            text="说明：当前“二级条目”设置会同时作用于目录二级及以下条目。",
            font=("Microsoft YaHei UI", 8),
            foreground=THEME["text_secondary"],
            justify="left",
        ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 6))
        r += 1
        r = self._row_line_spacing(p, r, "目录行距:", self._v_tocls_mode, self._v_tocls, hint="可直接选单倍/1.5倍/2倍；多倍如 1.2；固定值/最小值如 20pt")
        r = self._row_check(p, r, "摘要/Abstract 不参与目录", self._v_toc_ex_abs)
        self._ttk.Label(
            p,
            text="勾选后会解除摘要/Abstract 的 Heading 样式，并尽量保留原有外观。",
            font=("Microsoft YaHei UI", 8),
            foreground=THEME["text_secondary"],
            justify="left",
        ).grid(row=r, column=0, columnspan=3, sticky="w", pady=(0, 6))
        r += 1
        r = self._row_spin(p, r, "条目段前:", self._v_toc_sb, lo=0, hi=5, step=0.5, unit="行")
        r = self._row_spin(p, r, "条目段后:", self._v_toc_sa, lo=0, hi=5, step=0.5, unit="行")
        r = self._sep(p, r)
        self._ttk.Label(p, text="参考文献", font=("Microsoft YaHei UI", 11, "bold"),
                       foreground=THEME["primary"]).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(8, 6))
        r += 1
        r = self._row_spin(p, r, "参考文献缩进:", self._v_refind, lo=0, hi=100, step=1, unit="pt")
        r = self._sep(p, r)
        # special titles
        self._ttk.Label(p, text="特殊标题映射", font=("Microsoft YaHei UI", 11, "bold"),
                       foreground=THEME["primary"]).grid(
            row=r, column=0, columnspan=3, sticky="w", pady=(8, 6))
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
        self._create_button(bf, text="添加", width=6, command=self._add_st, bootstyle="secondary-outline").pack(side="left")
        self._create_button(bf, text="删除末行", width=8, command=self._del_st, bootstyle="secondary-outline").pack(side="left", padx=4)

    def _add_cov_field(self, label="", width=33):
        tk = self._tk
        row = len(self._cov_field_rows)
        f = self._cov_fields_frame
        lv = tk.StringVar(value=label)
        wv = tk.IntVar(value=width)
        le = self._ttk.Entry(f, textvariable=lv, width=16)
        le.grid(row=row, column=0, padx=(0, 4), pady=1)
        ws = self._ttk.Spinbox(f, from_=5, to=60, textvariable=wv, width=5, style="TSpinbox")
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
        ac = self._ttk.Combobox(f, textvariable=av, values=self.ALIGN_LABELS, width=8, state="readonly", style="Tool.TCombobox")
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
        bottom_panel = self._ttk.Frame(root, style="Card.TFrame", padding=(20, 16, 20, 18))
        bottom_panel.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        top_row = self._ttk.Frame(bottom_panel, style="Card.TFrame")
        top_row.pack(fill="x")

        left_stack = self._ttk.Frame(top_row, style="Card.TFrame")
        left_stack.pack(side="left", fill="both", expand=True)

        config_frame = self._ttk.LabelFrame(left_stack, text=" 配置管理 ")
        config_frame.pack(fill="x")
        config_frame.grid_columnconfigure(1, weight=1)
        self._ttk.Label(config_frame, text="当前配置:", style="Meta.TLabel").grid(row=0, column=0, sticky="w")
        self._ttk.Label(config_frame, textvariable=self._v_cfglbl, style="Value.TLabel").grid(row=0, column=1, sticky="w", padx=(8, 0))
        config_buttons = self._ttk.Frame(config_frame, style="Card.TFrame")
        config_buttons.grid(row=0, column=2, sticky="e")
        self._create_button(config_buttons, text="恢复默认", command=self._reset_defaults, width=8, bootstyle="secondary-outline").pack(side="left")
        self._create_button(config_buttons, text="保存配置", command=self._save_config, width=8, bootstyle="secondary-outline").pack(side="left", padx=6)
        self._create_button(config_buttons, text="加载配置", command=self._load_config, width=8, bootstyle="secondary-outline").pack(side="left")

        file_frame = self._ttk.LabelFrame(left_stack, text=" 文件与输出 ")
        file_frame.pack(fill="x", pady=(12, 0))
        file_frame.grid_columnconfigure(1, weight=1)
        self._ttk.Label(file_frame, text="输入文件:", style="Meta.TLabel").grid(row=0, column=0, sticky="w", pady=4)
        self._ttk.Entry(file_frame, textvariable=self._v_in, font=("Microsoft YaHei UI", 9)).grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=4, ipady=4)
        self._create_button(file_frame, text="浏览", command=self._browse_in, width=6, bootstyle="secondary-outline").grid(row=0, column=2, pady=4)
        self._ttk.Label(file_frame, text="输出文件:", style="Meta.TLabel").grid(row=1, column=0, sticky="w", pady=4)
        self._ttk.Entry(file_frame, textvariable=self._v_out, font=("Microsoft YaHei UI", 9)).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=4, ipady=4)
        self._create_button(file_frame, text="浏览", command=self._browse_out, width=6, bootstyle="secondary-outline").grid(row=1, column=2, pady=4)

        action_frame = self._ttk.LabelFrame(top_row, text=" 执行操作 ")
        action_frame.pack(side="right", padx=(16, 0))
        action_frame.grid_columnconfigure(0, weight=1)
        self._ttk.Label(action_frame, text="运行状态", style="Meta.TLabel").grid(row=0, column=0, sticky="w")
        self._ttk.Label(action_frame, textvariable=self._status_var, style="Value.TLabel").grid(row=1, column=0, sticky="w", pady=(2, 10))
        toggles = self._ttk.Frame(action_frame, style="Card.TFrame")
        toggles.grid(row=2, column=0, sticky="w", pady=(0, 10))
        self._ttk.Checkbutton(toggles, text="跳过目录生成", variable=self._v_skip, command=self._on_skip_toggle).pack(side="left")
        self._btn = self._create_button(action_frame, text="开始格式化", command=self._start, width=14, bootstyle="primary")
        self._btn.grid(row=3, column=0, sticky="ew")
        self._progress_frame = self._ttk.Frame(action_frame, style="Card.TFrame")
        self._progress_frame.grid(row=4, column=0, sticky="ew", pady=(12, 0))
        self._ttk.Label(self._progress_frame, text="正在执行格式化，请稍候。", style="Hint.TLabel").pack(anchor="w", pady=(0, 6))
        progress_kwargs = {"mode": "indeterminate", "length": 220}
        if self._has_theme:
            progress_kwargs["bootstyle"] = "info-striped"
        self._progress = self._ttk.Progressbar(self._progress_frame, **progress_kwargs)
        self._progress.pack(fill="x")
        self._progress_frame.grid_remove()

        log_frame = self._ttk.LabelFrame(bottom_panel, text=" 运行日志 ")
        log_frame.pack(fill="both", expand=True, pady=(16, 0))
        self._ttk.Label(log_frame, text="显示格式化过程、异常信息和完成结果，便于排查论文排版问题。", style="Hint.TLabel").pack(anchor="w", pady=(0, 8))
        self._log = self._scrolledtext.ScrolledText(
            log_frame,
            height=6,
            state="disabled",
            font=("Microsoft YaHei UI", 10),
            relief="flat",
            borderwidth=0,
            background=THEME["bg_panel"],
            foreground=THEME["text_primary"],
            insertbackground=THEME["text_primary"],
            padx=10,
            pady=8,
            spacing1=1,
            spacing3=2,
        )
        self._log.pack(fill="both", expand=True)

    def _shorten_path(self, path, max_len=64):
        if not path:
            return ""
        if len(path) <= max_len:
            return path
        keep = max_len // 2 - 2
        return f"{path[:keep]}...{path[-keep:]}"

    def _refresh_file_summary(self, *_args):
        self._header_input_var.set(self._shorten_path(self._v_in.get().strip()) or "未选择输入文件")
        self._header_output_var.set(self._shorten_path(self._v_out.get().strip()) or "未设置输出路径")

    def _set_status(self, text):
        self._status_var.set(text)

    def _on_theme_change(self, event):
        if not self._has_theme:
            return
        combo = event.widget
        new_theme = combo.get()
        if new_theme and new_theme != self._theme:
            self._style.theme_use(new_theme)
            self._theme = new_theme
            self._configure_modern_style()
            if self._cur_panel:
                self._update_cat_button_style(self._cur_panel)

    # ---- panel switching ----

    def _show_panel(self, name):
        if self._cur_panel:
            self._panels[self._cur_panel].pack_forget()
        self._panels[name].pack(fill="both", expand=True)
        self._cur_panel = name
        canvas = self._panel_canvas.get(name)
        if canvas:
            canvas.yview_moveto(0)
            self._set_active_scroll_canvas(canvas)

    def _on_cat_select(self, _event=None):
        cat_list = getattr(self, "_cat_list", None)
        if not cat_list:
            return
        sel = cat_list.curselection()
        if sel:
            self._on_cat_click(self.CATEGORIES[sel[0]])
    # ---- config ↔ vars ----

    @staticmethod
    def _numval(v):
        """float → int if whole, else float."""
        return int(v) if v == int(v) else v

    @staticmethod
    def _parse_unit_to_pt(value_str, default=12):
        """解析带单位的值，转换为 pt（磅）

        支持: 纯数字 / 数字+pt / 数字+cm / 数字+mm / 数字+in / 中文字号 / 磅
        返回: pt 值的数值（浮点数）
        """
        import re
        if not value_str:
            return default

        s = str(value_str).strip()

        # 中文字号映射
        cn_map = {"初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18,
                  "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9,
                  "六号": 7.5, "小六": 6.5, "七号": 5.5, "八号": 5}
        if s in cn_map:
            return cn_map[s]

        # "磅" 是 pt 的中文表达，先替换
        if "磅" in s:
            s = s.replace("磅", "pt")

        # 解析数字+单位
        m = re.match(r"^([\d.]+)([a-zA-Z]*)$", s)
        if m:
            num_str, unit = m.groups()
            try:
                num = float(num_str)
                if unit == "" or unit == "pt":
                    return num
                elif unit == "cm":
                    return num * 28.3465  # 1cm ≈ 28.35pt
                elif unit == "mm":
                    return num * 2.83465   # 1mm ≈ 2.83pt
                elif unit in ("in", "inch", "inches"):
                    return num * 72         # 1in = 72pt
                else:
                    return num  # 未知单位，原样返回
            except ValueError:
                return default

        # 纯数字
        try:
            return float(s)
        except ValueError:
            return default

    @classmethod
    def _resolve_line_spacing_choice(cls, choice):
        token = str(choice or "").strip()
        legacy = {"多倍": "多倍行距", "固定值": "固定值", "最小值": "最小值"}
        token = legacy.get(token, token)
        if token in cls._LINE_SPACING_MODE:
            return dict(cls._LINE_SPACING_MODE[token])
        return {"mode": normalize_line_spacing_mode(token)}

    @staticmethod
    def _normalize_line_spacing_value(mode, value_str):
        normalized_mode = normalize_line_spacing_mode(mode)
        raw = str(value_str or "").strip()
        if not raw:
            raw = "1.5" if normalized_mode == "multiple" else "20pt"
        if normalized_mode == "multiple":
            try:
                parsed = normalize_line_spacing(raw)
            except (TypeError, ValueError):
                parsed = None
            if parsed and parsed["mode"] == "multiple":
                return format_line_spacing_value(parsed)
            return format_line_spacing_value({"mode": "multiple", "value": 1.5})

        parsed = normalize_line_spacing(raw)
        if parsed["mode"] != "multiple":
            return normalize_length_text(parsed["value"])
        if raw.replace(".", "", 1).isdigit():
            return normalize_length_text(parsed["value"])
        return "20pt"

    @classmethod
    def _split_line_spacing_for_gui(cls, value):
        spec = normalize_line_spacing(value)
        if spec["mode"] == "multiple":
            multiple = float(spec["value"])
            if abs(multiple - 1.0) < 1e-9:
                mode_label = "单倍行距"
            elif abs(multiple - 1.5) < 1e-9:
                mode_label = "1.5倍行距"
            elif abs(multiple - 2.0) < 1e-9:
                mode_label = "2倍行距"
            else:
                mode_label = "多倍行距"
        else:
            mode_label = cls._LINE_SPACING_MODE_R.get(spec["mode"], "多倍行距")
        return mode_label, format_line_spacing_value(spec)

    @classmethod
    def _collect_line_spacing_config(cls, mode_label, value_text):
        choice = cls._resolve_line_spacing_choice(mode_label)
        preset = choice.get("preset")
        if preset is not None:
            return float(preset)

        mode = choice.get("mode", "multiple")
        normalized_value = cls._normalize_line_spacing_value(mode, value_text)
        spec = normalize_line_spacing({"mode": mode, "value": normalized_value})
        if spec["mode"] == "multiple":
            return spec["value"]
        return spec

    @staticmethod
    def _parse_spacing_to_config(value_str, default=0):
        """解析段前段后距离，转换为配置存储格式

        支持: 数字(行) / 数字+行 / 数字+pt / 数字+cm / 数字+mm / 数字+in / 磅
        返回: 配置值（如果是"行"单位则返回行数，否则返回 pt 值）
        """
        import re
        if not value_str:
            return default

        s = str(value_str).strip()

        # "磅" 是 pt 的中文表达，先替换
        if "磅" in s:
            s = s.replace("磅", "pt")

        # 解析数字+单位
        m = re.match(r"^([\d.]+)([a-zA-Z\u884c]*)$", s)
        if m:
            num_str, unit = m.groups()
            try:
                num = float(num_str)
                # "行" 单位 - 返回行数（保持原逻辑）
                if unit == "" or unit == "行":
                    return num
                # 其他单位 - 转换为 pt 值
                elif unit == "pt":
                    return num
                elif unit == "cm":
                    return num * 28.3465
                elif unit == "mm":
                    return num * 2.83465
                elif unit in ("in", "inch", "inches"):
                    return num * 72
                else:
                    return num
            except ValueError:
                return default

        # 纯数字
        try:
            return float(s)
        except ValueError:
            return default

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
        cfg["sizes"]["body"] = self._numval(self._parse_unit_to_pt(self._v_sbody.get()))
        cfg["sizes"]["h1"] = self._numval(self._parse_unit_to_pt(self._v_sh1.get()))
        cfg["sizes"]["h2"] = self._numval(self._parse_unit_to_pt(self._v_sh2.get()))
        cfg["sizes"]["h3"] = self._numval(self._parse_unit_to_pt(self._v_sh3.get()))
        cfg["sizes"]["h4"] = self._numval(self._parse_unit_to_pt(self._v_sh4.get()))
        cfg["sizes"]["caption"] = self._numval(self._parse_unit_to_pt(self._v_scap.get()))
        cfg["sizes"]["footnote"] = self._numval(self._parse_unit_to_pt(self._v_sfn.get()))
        # headings
        cfg["headings"]["h1"]["bold"] = self._BOLD.get(self._v_h1b.get(), True)
        cfg["headings"]["h1"]["align"] = self._ALIGN.get(self._v_h1a.get(), "left")
        cfg["headings"]["h1"]["space_before"] = self._parse_spacing_to_config(self._v_h1sb.get())
        cfg["headings"]["h1"]["space_after"] = self._parse_spacing_to_config(self._v_h1sa.get())
        cfg["headings"]["h2"]["bold"] = self._BOLD.get(self._v_h2b.get(), True)
        cfg["headings"]["h2"]["align"] = self._ALIGN.get(self._v_h2a.get(), "left")
        cfg["headings"]["h2"]["space_before"] = self._parse_spacing_to_config(self._v_h2sb.get())
        cfg["headings"]["h2"]["space_after"] = self._parse_spacing_to_config(self._v_h2sa.get())
        cfg["headings"]["h3"]["bold"] = self._BOLD.get(self._v_h3b.get(), False)
        cfg["headings"]["h3"]["align"] = self._ALIGN.get(self._v_h3a.get(), "left")
        cfg["headings"]["h3"]["space_before"] = self._parse_spacing_to_config(self._v_h3sb.get())
        cfg["headings"]["h3"]["space_after"] = self._parse_spacing_to_config(self._v_h3sa.get())
        cfg["headings"]["h4"]["bold"] = self._BOLD.get(self._v_h4b.get(), False)
        cfg["headings"]["h4"]["align"] = self._ALIGN.get(self._v_h4a.get(), "left")
        cfg["headings"]["h4"]["space_before"] = self._parse_spacing_to_config(self._v_h4sb.get())
        cfg["headings"]["h4"]["space_after"] = self._parse_spacing_to_config(self._v_h4sa.get())
        # body
        cfg["body"]["line_spacing"] = self._collect_line_spacing_config(self._v_lsp_mode.get(), self._v_lsp.get())
        cfg["body"]["first_line_indent"] = self._numval(self._v_ind.get())
        cfg["body"]["align"] = self._ALIGN.get(self._v_body_align.get(), "justify")
        cfg["body"]["space_before"] = self._parse_spacing_to_config(self._v_body_sb.get())
        cfg["body"]["space_after"] = self._parse_spacing_to_config(self._v_body_sa.get())
        # sections (heading numbering patterns)
        cfg["sections"]["chapter_pattern"] = self._v_pat_h1.get()
        cfg["sections"]["h2_pattern"] = self._v_pat_h2.get()
        cfg["sections"]["h3_pattern"] = self._v_pat_h3.get()
        cfg["sections"]["h4_pattern"] = self._v_pat_h4.get()
        cfg["sections"]["renumber_headings"] = self._v_renum.get()
        # captions
        cfg["captions"] = {
            "mode": self._CAPTION_MODE.get(self._v_cap_mode.get(), "dynamic"),
            "figure_pattern": self._v_cap_fig.get(),
            "table_pattern": self._v_cap_tbl.get(),
            "subfigure_pattern": self._v_cap_sub.get(),
            "note_pattern": self._v_cap_note.get(),
            "keep_with_next": self._v_cap_kwn.get(),
            "check_numbering": self._v_cap_chk.get(),
            "include_chapter": self._v_cap_include_chapter.get(),
            "restart_per_chapter": self._v_cap_restart_chapter.get(),
            "chapter_heading_level": int(self._v_cap_heading_level.get() or "1"),
            "chapter_separator": self._v_cap_chapter_sep.get(),
            "caption_separator": self._v_cap_caption_sep.get(),
            "line_spacing": self._collect_line_spacing_config(self._v_cap_ls_mode.get(), self._v_cap_ls.get()),
            "font": self._v_cap_font.get(),
            "number_font": self._v_cap_numfont.get(),
        }
        # cover
        cfg["cover"]["enabled"] = self._v_cov_en.get()
        cfg["cover"]["only_insert"] = self._v_cover_only.get()
        cfg["page_numbers"]["only_insert"] = self._v_pn_only.get()
        cfg["header_footer"]["only_insert"] = self._v_hf_only.get()
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
        cfg["toc"]["enabled"] = not self._v_skip.get()
        cfg["toc"]["only_insert"] = bool(self._v_toc_only.get() and not self._v_skip.get())
        cfg["toc"]["exclude_abstract_headings"] = self._v_toc_ex_abs.get()
        cfg["toc"]["font"] = self._v_tocfont.get()
        cfg["toc"]["font_size"] = self._numval(self._parse_unit_to_pt(self._v_tocsz.get()))
        cfg["toc"]["bold"] = self._v_toc_bold.get()
        cfg["toc"]["h1_font"] = self._v_toc_h1font.get()
        cfg["toc"]["h1_font_size"] = self._numval(self._parse_unit_to_pt(self._v_toc_h1sz.get()))
        cfg["toc"]["h1_bold"] = self._v_toc_h1bold.get()
        cfg["toc"]["line_spacing"] = self._collect_line_spacing_config(self._v_tocls_mode.get(), self._v_tocls.get())
        cfg["toc"]["space_before"] = self._parse_spacing_to_config(self._v_toc_sb.get())
        cfg["toc"]["space_after"] = self._parse_spacing_to_config(self._v_toc_sa.get())
        cfg["references"]["left_indent"] = self._numval(self._v_refind.get())
        cfg["references"]["first_line_indent"] = -self._numval(self._v_refind.get())
        cfg["table"]["top_border_sz"] = self._numval(self._v_tbl_top.get() * 8)
        cfg["table"]["header_border_sz"] = self._numval(self._v_tbl_hdr.get() * 8)
        cfg["table"]["bottom_border_sz"] = self._numval(self._v_tbl_bot.get() * 8)
        cfg["table"]["line_spacing"] = self._collect_line_spacing_config(self._v_tbl_ls_mode.get(), self._v_tbl_ls.get())
        cfg["footnote"]["line_spacing"] = self._collect_line_spacing_config(self._v_fn_ls_mode.get(), self._v_fn_ls.get())
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
        cfg["sizes"]["page_number"] = self._numval(self._parse_unit_to_pt(self._v_pn_size.get()))
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
        cfg["header_footer"]["font_size"] = self._numval(self._parse_unit_to_pt(self._v_hf_size.get()))
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
        self._v_h1sb.set(format_paragraph_spacing_value(cfg["headings"]["h1"].get("space_before", 0)))
        self._v_h1sa.set(format_paragraph_spacing_value(cfg["headings"]["h1"].get("space_after", 0)))
        self._v_h2b.set(self._BOLD_R.get(cfg["headings"]["h2"]["bold"], "加粗"))
        self._v_h2a.set(self._ALIGN_R.get(cfg["headings"]["h2"]["align"], "左对齐"))
        self._v_h2sb.set(format_paragraph_spacing_value(cfg["headings"]["h2"].get("space_before", 0)))
        self._v_h2sa.set(format_paragraph_spacing_value(cfg["headings"]["h2"].get("space_after", 0)))
        self._v_h3b.set(self._BOLD_R.get(cfg["headings"]["h3"]["bold"], "不加粗"))
        self._v_h3a.set(self._ALIGN_R.get(cfg["headings"]["h3"].get("align", "left"), "左对齐"))
        self._v_h3sb.set(format_paragraph_spacing_value(cfg["headings"]["h3"].get("space_before", 0)))
        self._v_h3sa.set(format_paragraph_spacing_value(cfg["headings"]["h3"].get("space_after", 0)))
        self._v_h4b.set(self._BOLD_R.get(cfg["headings"]["h4"]["bold"], "不加粗"))
        self._v_h4a.set(self._ALIGN_R.get(cfg["headings"]["h4"].get("align", "left"), "左对齐"))
        self._v_h4sb.set(format_paragraph_spacing_value(cfg["headings"]["h4"].get("space_before", 0)))
        self._v_h4sa.set(format_paragraph_spacing_value(cfg["headings"]["h4"].get("space_after", 0)))
        # body
        body_ls_mode, body_ls_value = self._split_line_spacing_for_gui(cfg["body"]["line_spacing"])
        self._v_lsp_mode.set(body_ls_mode)
        self._v_lsp.set(body_ls_value)
        self._v_ind.set(cfg["body"]["first_line_indent"])
        self._v_body_sb.set(format_paragraph_spacing_value(cfg["body"].get("space_before", 0)))
        self._v_body_sa.set(format_paragraph_spacing_value(cfg["body"].get("space_after", 0)))
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
        self._v_cap_mode.set(self._CAPTION_MODE_R.get(cap.get("mode", "dynamic"), "严格动态模式 (推荐)"))
        self._v_cap_fig.set(cap.get("figure_pattern", r"^图\s*\d"))
        self._v_cap_tbl.set(cap.get("table_pattern", r"^(续)?表\s*\d"))
        self._v_cap_sub.set(cap.get("subfigure_pattern", r"^\([a-z]\)"))
        self._v_cap_note.set(cap.get("note_pattern", r"^注[：:]"))
        self._v_cap_kwn.set(cap.get("keep_with_next", True))
        self._v_cap_chk.set(cap.get("check_numbering", True))
        self._v_cap_include_chapter.set(cap.get("include_chapter", False))
        self._v_cap_restart_chapter.set(cap.get("restart_per_chapter", False))
        self._v_cap_heading_level.set(str(cap.get("chapter_heading_level", 1)))
        self._v_cap_chapter_sep.set(cap.get("chapter_separator", "."))
        self._v_cap_caption_sep.set(cap.get("caption_separator", ""))
        cap_ls_mode, cap_ls_value = self._split_line_spacing_for_gui(cap.get("line_spacing", cfg["body"]["line_spacing"]))
        self._v_cap_ls_mode.set(cap_ls_mode)
        self._v_cap_ls.set(cap_ls_value)
        self._v_cap_font.set(cap.get("font", "宋体"))
        self._v_cap_numfont.set(cap.get("number_font", "Times New Roman"))
        # cover
        self._v_cov_en.set(cfg["cover"]["enabled"])
        self._v_school.set(cfg["meta"]["school_name"])
        self._v_logo.set(cfg["cover"]["logo"])
        self._v_covtitle.set(cfg["cover"]["title_text"])
        self._v_custom_cover.set(cfg["cover"].get("custom_docx", ""))
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
        self._v_skip.set(not cfg["toc"].get("enabled", True))
        self._v_toc_only.set(cfg["toc"].get("only_insert", False))
        self._v_toc_ex_abs.set(cfg["toc"].get("exclude_abstract_headings", False))
        self._v_cover_only.set(cfg["cover"].get("only_insert", False))
        self._v_pn_only.set(cfg["page_numbers"].get("only_insert", False))
        self._v_hf_only.set(cfg["header_footer"].get("only_insert", False))
        self._v_tocfont.set(cfg["toc"].get("font", cfg["fonts"]["body"]))
        self._v_tocsz.set(str(self._numval(cfg["toc"].get("font_size", cfg["sizes"]["body"]))) + "pt")
        self._v_toc_bold.set(cfg["toc"].get("bold", False))
        self._v_toc_h1font.set(cfg["toc"].get("h1_font", cfg["fonts"]["h1"]))
        self._v_toc_h1sz.set(str(self._numval(cfg["toc"].get("h1_font_size", cfg["sizes"]["h1"]))) + "pt")
        self._v_toc_h1bold.set(cfg["toc"].get("h1_bold", False))
        toc_ls_mode, toc_ls_value = self._split_line_spacing_for_gui(cfg["toc"].get("line_spacing", cfg["body"]["line_spacing"]))
        self._v_tocls_mode.set(toc_ls_mode)
        self._v_tocls.set(toc_ls_value)
        self._v_toc_sb.set(format_paragraph_spacing_value(cfg["toc"].get("space_before", 0)))
        self._v_toc_sa.set(format_paragraph_spacing_value(cfg["toc"].get("space_after", 0)))
        self._normalize_standalone_mode_state()
        self._v_refind.set(cfg["references"]["left_indent"])
        self._v_tbl_top.set(cfg["table"]["top_border_sz"] / 8)
        self._v_tbl_hdr.set(cfg["table"]["header_border_sz"] / 8)
        self._v_tbl_bot.set(cfg["table"]["bottom_border_sz"] / 8)
        tbl_ls_mode, tbl_ls_value = self._split_line_spacing_for_gui(cfg["table"].get("line_spacing", 1.0))
        self._v_tbl_ls_mode.set(tbl_ls_mode)
        self._v_tbl_ls.set(tbl_ls_value)
        fn_ls_mode, fn_ls_value = self._split_line_spacing_for_gui(cfg["footnote"].get("line_spacing", 1.0))
        self._v_fn_ls_mode.set(fn_ls_mode)
        self._v_fn_ls.set(fn_ls_value)
        # front_matter
        fm = cfg.get("front_matter", {})
        self._v_fm_mode.set(
            self._FM_MODE_R.get(fm.get("mode", "auto"), "自动识别并格式化")
        )
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

    def _clear_standalone_modes(self):
        self._v_toc_only.set(False)
        self._v_cover_only.set(False)
        self._v_pn_only.set(False)
        self._v_hf_only.set(False)

    def _activate_standalone_mode(self, mode):
        self._v_skip.set(False)
        self._v_toc_only.set(mode == "toc")
        self._v_cover_only.set(mode == "cover")
        self._v_pn_only.set(mode == "pn")
        self._v_hf_only.set(mode == "hf")

    def _normalize_standalone_mode_state(self):
        if self._v_toc_only.get():
            self._activate_standalone_mode("toc")
        elif self._v_cover_only.get():
            self._activate_standalone_mode("cover")
        elif self._v_pn_only.get():
            self._activate_standalone_mode("pn")
        elif self._v_hf_only.get():
            self._activate_standalone_mode("hf")
        elif self._v_skip.get():
            self._clear_standalone_modes()

    def _on_skip_toggle(self):
        if self._v_skip.get():
            self._clear_standalone_modes()

    def _on_toc_only_toggle(self):
        if self._v_toc_only.get():
            self._activate_standalone_mode("toc")

    def _on_cover_only_toggle(self):
        if self._v_cover_only.get():
            self._activate_standalone_mode("cover")

    def _on_pn_only_toggle(self):
        if self._v_pn_only.get():
            self._activate_standalone_mode("pn")

    def _on_hf_only_toggle(self):
        if self._v_hf_only.get():
            self._activate_standalone_mode("hf")
    # ---- run ----

    def _start(self):
        inp = self._v_in.get().strip()
        out = self._v_out.get().strip()
        if not inp or not os.path.isfile(inp):
            self._set_status("等待开始")
            self._messagebox.showerror("错误", "请选择有效的输入文件。")
            return
        if not out:
            self._set_status("等待开始")
            self._messagebox.showerror("错误", "请指定输出文件路径。")
            return

        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

        self._btn.config(state="disabled")
        self._running = True
        self._set_status("正在格式化")
        self._progress_frame.grid()
        self._progress.start(10)
        self._root.after(100, self._poll)

        skip = self._v_skip.get()
        toc_only = self._v_toc_only.get() and not skip
        cover_only = self._v_cover_only.get()
        pn_only = self._v_pn_only.get()
        hf_only = self._v_hf_only.get()
        self._append_log(f"输入文件: {inp}")
        self._append_log(f"输出文件: {out}")
        if cover_only:
            self._append_log("单独处理: 单独插入外部封面")
        elif pn_only:
            self._append_log("单独处理: 仅更新页码")
        elif hf_only:
            self._append_log("单独处理: 仅更新页眉")
        else:
            if skip:
                toc_status = "跳过生成"
            elif toc_only:
                toc_status = "仅插入/更新目录（不改其他格式）"
            else:
                toc_status = "正常生成"
            self._append_log(f"目录处理: {toc_status}")
        self._append_log("开始执行格式化。")
        try:
            config = self._collect_config()
        except (ValueError, Exception) as e:
            self._messagebox.showerror("错误", f"参数值无效: {e}")
            self._btn.config(state="normal")
            self._running = False
            self._set_status("参数检查失败")
            self._progress.stop()
            self._progress_frame.grid_remove()
            return
        if cover_only and not config.get("cover", {}).get("custom_docx"):
            self._messagebox.showerror("错误", "仅插入外部封面前，请先在“封面声明”中选择自定义封面 .docx。")
            self._btn.config(state="normal")
            self._running = False
            self._set_status("参数检查失败")
            self._progress.stop()
            self._progress_frame.grid_remove()
            return

        def worker():
            final_status = "格式化失败"
            try:
                ok = run_format(inp, out, self._append_log, config=config)
                final_status = "格式化完成" if ok else "格式化失败"
                self._append_log(f"\n--- {final_status} ---")
            except Exception as e:
                final_status = "运行异常"
                self._append_log(f"\n异常: {e}")
            finally:
                self._running = False
                self._root.after(0, lambda s=final_status: self._set_status(s))
                self._root.after(0, lambda: self._btn.config(state="normal"))
                self._root.after(0, lambda: self._progress.stop())
                self._root.after(0, lambda: self._progress_frame.grid_remove())

        threading.Thread(target=worker, daemon=True).start()


def main(theme="sandstone"):
    """Launch the GUI with the fixed default theme."""
    FormatterGUI(theme="sandstone")


if __name__ == "__main__":
    main()




