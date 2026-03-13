"""Configuration loader for the universal thesis formatter.

Loads formatting parameters from YAML config files, with deep-merge
over built-in defaults (SCAU 2024). Users only need to specify values
they want to change; all unspecified values fall back to defaults.
"""

import copy
import os
import sys

try:
    import yaml
    HAS_YAML = True
except ImportError:
    HAS_YAML = False


def _resource_dir():
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


# ---------------------------------------------------------------------------
# Built-in defaults (SCAU 2024) — used when no config file is provided
# ---------------------------------------------------------------------------

_SCAU_DECLARATION_1 = (
    "本人郑重声明：所呈交的毕业论文（设计），是本人在导师的指导下，"
    "独立进行研究工作所取得的成果。除文中已经注明引用的内容外，本论文"
    "不包含任何其他个人或集体已经发表或撰写过的作品成果。对本文的研究"
    "做出重要贡献的个人和集体，均已在文中以明确方式标明。本人完全意识"
    "到本声明的法律结果由本人承担。"
)

_SCAU_DECLARATION_2 = (
    "本人完全了解学校有关保留、使用毕业论文（设计）的规定，同意学校"
    "保留并向国家有关部门或机构送交毕业论文（设计）的复印件和电子版，"
    "允许毕业论文（设计）被查阅和借阅。学校可以将本毕业论文（设计）的"
    "全部或部分内容编入有关数据库进行检索，可以采用影印、缩印或扫描等"
    "复制手段保存和汇编毕业论文（设计）。"
)

DEFAULT_CONFIG = {
    "meta": {
        "school_name": "华南农业大学",
        "config_version": 1,
    },
    "page": {
        "margins": {"top": 2.4, "bottom": 2.4, "left": 2.4, "right": 2.4},
        "gutter": 0.5,
        "header_distance": 1.5,
        "footer_distance": 1.25,
    },
    "fonts": {
        "latin": "Times New Roman",
        "body": "宋体",
        "h1": "黑体",
        "h2": "黑体",
        "h3": "楷体",
        "h4": "楷体",
    },
    "sizes": {
        "body": 12,
        "h1": 14,
        "h2": 12,
        "h3": 12,
        "h4": 12,
        "caption": 10.5,
        "note": 9,
        "footnote": 9,
        "page_number": 10.5,
    },
    "headings": {
        "h1": {"bold": True, "align": "left"},
        "h2": {"bold": True, "align": "left"},
        "h3": {"bold": False, "align": "left"},
        "h4": {"bold": False, "align": "left"},
    },
    "body": {
        "align": "justify",
        "first_line_indent": 24,
        "line_spacing": 1.5,
    },
    "table": {
        "line_spacing": 1.0,
        "top_border_sz": 12,
        "header_border_sz": 8,
        "bottom_border_sz": 12,
    },
    "footnote": {
        "line_spacing": 1.0,
    },
    "captions": {
        "figure_pattern": r"^图\s*\d",
        "table_pattern": r"^(续)?表\s*\d",
        "subfigure_pattern": r"^\([a-z]\)",
        "note_pattern": r"^注[：:]",
        "keep_with_next": True,
        "check_numbering": True,
    },
    "references": {
        "first_line_indent": -24,
        "left_indent": 24,
    },
    "page_numbers": {
        "front_format": "upperRoman",
        "body_format": "decimal",
        "front_start": 1,
        "body_start": 1,
    },
    "toc": {
        "depth": 3,
        "font": "宋体",
        "font_size": 12,
        "line_spacing": 1.5,
    },
    "special_titles": [
        {"match": "摘要", "display": "摘        要", "align": "center"},
        {"match": "目录", "display": "目        录", "align": "center"},
        {"match": "参考文献", "display": "参  考  文  献", "align": "center"},
        {"match": "致谢", "display": "致        谢", "align": "center"},
    ],
    "sections": {
        "chapter_pattern": r"^第\s*\d+\s*章\b",
        "appendix_pattern": r"^附录\s*[A-Z]",
        "h2_pattern": r"^\d+\.\d+\s",
        "h3_pattern": r"^\d+\.\d+\.\d+\s",
        "h4_pattern": r"^\d+\.\d+\.\d+\.\d+\s",
        "special_h1": ["参考文献", "致谢"],
        "renumber_headings": True,
        "cn_keywords_pattern": r"^\s*关键词\s*[：:]",
        "en_abstract_pattern": r"(?i)^\s*Abstract\s*[：:]",
        "en_keywords_pattern": r"(?i)^\s*Key\s*words\s*[：:]",
    },
    "cover": {
        "enabled": True,
        "logo": "scau_logo.png",
        "logo_width_pt": 343.2,
        "logo_height_pt": 96,
        "title_text": "本科毕业论文(或设计)",
        "title_font_size": 36,
        "thesis_title_placeholder": "论文（或设计）题目",
        "thesis_title_font": "黑体",
        "thesis_title_size": 22,
        "fields": [
            {"label": "学    院:", "underline_chars": 34},
            {"label": "专    业:", "underline_chars": 34},
            {"label": "姓    名:", "underline_chars": 34},
            {"label": "学    号:", "underline_chars": 34},
        ],
        "advisor": {
            "label": "指导教师:",
            "underline_chars": 16,
            "title_label": "职称",
            "title_underline_chars": 16,
        },
        "date": {
            "label": "提交日期：",
            "segments": ["年", "月", "日"],
            "segment_underline_chars": 10,
        },
    },
    "declarations": [
        {
            "title": "华南农业大学本科毕业论文（设计）原创性声明",
            "body": _SCAU_DECLARATION_1,
            "signature": "作者签名：                        "
                         "日期：       年     月     日",
        },
        {
            "title": "华南农业大学本科毕业论文（设计）使用授权声明",
            "body": _SCAU_DECLARATION_2,
            "signature": "作者签名：                           "
                         "指导教师签名：                        ",
            "date_line": "日期：       年      月      日      "
                         "日期：     年      月      日",
        },
    ],
    "theme_fonts": {
        "latin": "Times New Roman",
        "hans": "宋体",
    },
}


# ---------------------------------------------------------------------------
# Deep merge
# ---------------------------------------------------------------------------

def _deep_merge(base, override):
    """Recursively merge *override* into a copy of *base*."""
    result = copy.deepcopy(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = copy.deepcopy(v)
    return result


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def load_config(path):
    """Load a YAML config and merge over defaults. Returns merged dict."""
    if not HAS_YAML:
        raise RuntimeError(
            "需要 pyyaml 来加载配置文件。请运行: pip install pyyaml"
        )
    with open(path, "r", encoding="utf-8") as f:
        user = yaml.safe_load(f) or {}
    return _deep_merge(DEFAULT_CONFIG, user)


def resolve_config(cli_config=None, input_path=None):
    """Find and load config using priority: CLI > input dir > exe dir > builtin."""
    candidates = []
    if cli_config:
        candidates.append(os.path.abspath(cli_config))
    if input_path:
        candidates.append(
            os.path.join(os.path.dirname(os.path.abspath(input_path)),
                         "thesis_config.yaml")
        )
    if getattr(sys, "frozen", False):
        candidates.append(
            os.path.join(os.path.dirname(sys.executable),
                         "thesis_config.yaml")
        )
    for path in candidates:
        if os.path.isfile(path):
            return load_config(path), path
    return copy.deepcopy(DEFAULT_CONFIG), None


def resolve_logo_path(cfg, config_path=None):
    """Resolve logo file path relative to config file, exe, or resource dir."""
    logo = cfg.get("cover", {}).get("logo", "")
    if not logo:
        return None
    if os.path.isabs(logo):
        return logo if os.path.isfile(logo) else None
    search = []
    if config_path:
        search.append(os.path.join(os.path.dirname(config_path), logo))
    if getattr(sys, "frozen", False):
        search.append(os.path.join(os.path.dirname(sys.executable), logo))
    search.append(os.path.join(_resource_dir(), logo))
    search.append(os.path.join(_resource_dir(), "defaults", logo))
    for p in search:
        if os.path.isfile(p):
            return p
    return None


def dump_default_config():
    """Return the built-in default config as a YAML string."""
    if not HAS_YAML:
        raise RuntimeError("需要 pyyaml。请运行: pip install pyyaml")
    return yaml.dump(DEFAULT_CONFIG, allow_unicode=True, default_flow_style=False,
                     sort_keys=False)
