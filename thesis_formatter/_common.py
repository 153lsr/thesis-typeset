import os
import re
import sys

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Inches, Pt, RGBColor

_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_JC_MAP = {
    "left": "left",
    "center": "center",
    "right": "right",
    "justify": "both",
}

# 中文字号映射（单位：磅）
_CN_FONT_SIZE_MAP = {
    "初号": 42,
    "小初": 36,
    "一号": 26,
    "小一": 24,
    "二号": 22,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "小四": 12,
    "五号": 10.5,
    "小五": 9,
    "六号": 7.5,
    "小六": 6.5,
    "七号": 5.5,
    "八号": 5,
}


def parse_length(value):
    """解析带单位的长度值，返回对应的 Length 对象

    支持的格式：
    - 数字（默认为 pt）: 12 → Pt(12)
    - pt 后缀: "12pt" → Pt(12)
    - cm 后缀: "2.54cm" → Cm(2.54)
    - mm 后缀: "25.4mm" → Mm(25.4)
    - in/inches 后缀: "1in" / "1inches" → Inches(1)
    - 中文字号: "小四" → Pt(12)

    Args:
        value: 可以是数字、字符串（带单位后缀），或已经是 Pt/Cm/Mm/Inches 对象

    Returns:
        docx.shared.Length 对象（Pt/Cm/Mm/Inches）
    """
    # 如果已经是 Length 对象，直接返回
    if hasattr(value, "pt"):  # Pt/Cm/Mm/Inches 都有 .pt 属性
        return value

    # 如果是纯数字，默认为 pt
    if isinstance(value, (int, float)):
        return Pt(value)

    # 转换为字符串处理
    s = str(value).strip().lower()

    # 尝试匹配中文字号
    if s in _CN_FONT_SIZE_MAP:
        return Pt(_CN_FONT_SIZE_MAP[s])

    # 使用正则提取数字和单位
    match = re.match(r"^([\d.]+)([a-z]*)$", s)
    if not match:
        # 无法解析，尝试作为数字处理
        try:
            return Pt(float(s))
        except (ValueError, TypeError):
            return Pt(0)

    num_str, unit = match.groups()
    try:
        num = float(num_str)
    except ValueError:
        return Pt(0)

    # 根据单位后缀返回对应的 Length 对象
    if unit in ("pt", ""):
        return Pt(num)
    elif unit == "cm":
        return Cm(num)
    elif unit == "mm":
        return Mm(num)
    elif unit in ("in", "inch", "inches"):
        return Inches(num)
    else:
        # 未知单位，默认为 pt
        return Pt(num)


def _resource_dir():
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


# 标题样式映射 (固定的 style_id)
_HEADING_STYLE_IDS = {
    1: "Heading1",
    2: "Heading2",
    3: "Heading3",
    4: "Heading4",
}

# 常见的标题样式名称（多语言）
_HEADING_STYLE_NAMES = {
    1: ["Heading 1", "标题 1", "样式1", "Heading1", "标题1"],
    2: ["Heading 2", "标题 2", "Heading2", "标题2"],
    3: ["Heading 3", "标题 3", "Heading3", "标题3"],
    4: ["Heading 4", "标题 4", "Heading4", "标题4"],
}

_ALL_HEADING_NAMES = set()
for names in _HEADING_STYLE_NAMES.values():
    _ALL_HEADING_NAMES.update(names)


def get_heading_style(doc, level):
    """通过 level (1-4) 获取文档中的标题样式对象，优先使用 style_id 匹配"""
    if level not in _HEADING_STYLE_IDS:
        return None
    style_id = _HEADING_STYLE_IDS[level]

    # 优先通过 style_id 查找
    for style in doc.styles:
        if style.style_id == style_id:
            return style

    # fallback: 通过名称查找
    for name in _HEADING_STYLE_NAMES.get(level, []):
        try:
            return doc.styles[name]
        except KeyError:
            continue
    return None


def get_heading_style_by_id_or_name(doc, level):
    """获取标题样式，返回 (style, style_id, style_name)"""
    style = get_heading_style(doc, level)
    if style:
        return style, style.style_id, style.name
    return None, None, None


def is_heading_style(style, level=None):
    """检查样式是否是标题样式

    Args:
        style: 样式对象或样式名称
        level: 标题级别 (1-4)，如果为 None 则检查是否为任何标题样式
    """
    if style is None:
        return False

    # 如果传入的是字符串（样式名）
    if isinstance(style, str):
        style_name = style
        style_id = None
    else:
        style_name = style.name
        style_id = getattr(style, 'style_id', None)

    # 优先通过 style_id 匹配
    if style_id:
        if level is not None:
            return style_id == _HEADING_STYLE_IDS.get(level)
        return style_id in _HEADING_STYLE_IDS.values()

    # 通过样式名称匹配
    if level is not None:
        return style_name in _HEADING_STYLE_NAMES.get(level, [])
    return style_name in _ALL_HEADING_NAMES


def is_heading(para, level=None):
    """检查段落的样式是否是标题样式

    Args:
        para: 段落对象
        level: 标题级别 (1-4)，如果为 None 则检查是否为任何标题样式
    """
    if not para.style:
        return False
    return is_heading_style(para.style, level)


def get_paragraph_heading_level(para):
    """获取段落的标题级别 (1-4)，如果不是标题则返回 None"""
    if not para.style:
        return None

    style_id = getattr(para.style, 'style_id', None)
    if style_id:
        for level, sid in _HEADING_STYLE_IDS.items():
            if style_id == sid:
                return level

    style_name = para.style.name
    for level, names in _HEADING_STYLE_NAMES.items():
        if style_name in names:
            return level
    return None


def set_rfonts(rpr, east_asia, latin="Times New Roman"):
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), east_asia)


def set_style_font(style, east_asia, size_pt, bold=False, latin="Times New Roman"):
    style.font.name = latin
    style.font.size = size_pt
    if bold is not None:
        style.font.bold = bold
    style.font.color.rgb = RGBColor(0, 0, 0)
    rpr = style.element.get_or_add_rPr()
    set_rfonts(rpr, east_asia, latin)


def set_run_font(run, east_asia, size_pt=None, bold=None, latin="Times New Roman"):
    run.font.name = latin
    if size_pt is not None:
        run.font.size = size_pt
    if bold is not None:
        run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)
    rpr = run._element.get_or_add_rPr()
    set_rfonts(rpr, east_asia, latin)


def set_para_runs_font(para, east_asia, size_pt, bold=None, latin="Times New Roman"):
    for run in para.runs:
        set_run_font(run, east_asia=east_asia, size_pt=size_pt, bold=bold, latin=latin)


def zero_spacing(para):
    pf = para.paragraph_format
    pf.space_before = parse_length(0)
    pf.space_after = parse_length(0)


def set_table_border(cell, edge, sz, val="single", color="000000"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_borders = tc_pr.find(qn("w:tcBorders"))
    if tc_borders is None:
        tc_borders = OxmlElement("w:tcBorders")
        tc_pr.append(tc_borders)
    edge_el = tc_borders.find(qn(f"w:{edge}"))
    if edge_el is None:
        edge_el = OxmlElement(f"w:{edge}")
        tc_borders.append(edge_el)
    edge_el.set(qn("w:val"), val)
    edge_el.set(qn("w:sz"), str(sz))
    edge_el.set(qn("w:color"), color)


def clear_table_border(cell, edge):
    set_table_border(cell, edge, sz=0, val="nil")


def _ensure_keep_next(p_el):
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    if pPr.find(qn("w:keepNext")) is None:
        pPr.append(OxmlElement("w:keepNext"))


def _set_para_spacing(p_el, side, pt_val):
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_el.insert(0, pPr)
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    twips = str(int(pt_val.pt * 20))
    spacing.set(qn("w:" + side), twips)


def _check_caption_numbering(doc, fig_pat, tbl_pat, cfg=None):
    warnings = []
    sec = cfg.get("sections", {}) if cfg else {}
    appendix_re = re.compile(sec.get("appendix_pattern", r"^附录\s*[A-Z]"))
    note_re = re.compile(cfg.get("captions", {}).get("note_pattern", r"^注[：:]")) if cfg else re.compile(r"^注[：:]")
    source_re = re.compile(r"^(资料)?来源\s*[：:]")

    def _parse_body_caption_path(text):
        match = re.search(r"(\d+(?:\s*[.\-]\s*\d+)*)", text)
        if not match:
            return None
        normalized = re.sub(r"\s*([.\-])\s*", r"\1", match.group(1))
        parts = [int(part) for part in re.split(r"[.\-]", normalized) if part]
        if not parts:
            return None
        return tuple(parts) if len(parts) > 1 else parts[0]

    def _format_caption_path(value):
        if isinstance(value, tuple):
            return ".".join(str(part) for part in value)
        return str(value)

    def _element_text(el):
        return "".join((node.text or "") for node in el.iter(qn("w:t"))).strip()

    def _find_prev_meaningful(children, idx):
        probe = idx - 1
        while probe >= 0:
            if children[probe].tag != qn("w:p") or _element_text(children[probe]):
                return probe
            probe -= 1
        return None

    def _find_next_meaningful(children, idx):
        probe = idx + 1
        while probe < len(children):
            if children[probe].tag != qn("w:p") or _element_text(children[probe]):
                return probe
            probe += 1
        return None

    current_appendix = None
    body_figs, body_tbls = [], []
    app_figs = {}
    app_tbls = {}

    for para in doc.paragraphs:
        t = para.text.strip()
        sn = para.style.name if para.style else ""

        if sn in ("Heading 1", "样式1") and appendix_re.match(t):
            m = re.search(r"附录\s*([A-Z])", t)
            if m:
                current_appendix = m.group(1)
            continue

        if current_appendix:
            m = re.match(r"^图([A-Z])(\d+)", t)
            if m:
                app_figs.setdefault(m.group(1), []).append(int(m.group(2)))
                continue
            m = re.match(r"^(续)?表([A-Z])(\d+)", t)
            if m:
                if m.group(1):
                    continue
                app_tbls.setdefault(m.group(2), []).append(int(m.group(3)))
                continue
        else:
            if re.match(fig_pat, t) or re.match(r"^Figure\s*\d", t, re.I):
                path = _parse_body_caption_path(t)
                if path is not None:
                    body_figs.append(path)
            elif re.match(tbl_pat, t) or re.match(r"^Table\s*\d", t, re.I):
                if re.match(r"^续", t):
                    continue
                path = _parse_body_caption_path(t)
                if path is not None:
                    body_tbls.append(path)

    for label, nums in [("图", body_figs), ("表", body_tbls)]:
        if not nums:
            continue
        if any(isinstance(n, tuple) for n in nums):
            grouped = {}
            for item in nums:
                if isinstance(item, tuple):
                    prefix = item[:-1]
                    seq = item[-1]
                    grouped.setdefault(prefix, []).append(seq)
            for prefix, seqs in grouped.items():
                expected = list(range(1, len(seqs) + 1))
                if seqs != expected:
                    prefix_text = ".".join(str(part) for part in prefix)
                    if len(prefix) == 1:
                        warnings.append(
                            f"  警告: 正文第{prefix[0]}章{label}编号不连续 — 期望 {expected}, 实际 {seqs}"
                        )
                    else:
                        warnings.append(
                            f"  警告: 正文编号前缀 {prefix_text} 的{label}编号不连续 — 期望 {expected}, 实际 {seqs}"
                        )
        else:
            expected = list(range(1, len(nums) + 1))
            if nums != expected:
                warnings.append(f"  警告: 正文{label}编号不连续 — 期望 {expected}, 实际 {nums}")
    for letter in sorted(set(list(app_figs.keys()) + list(app_tbls.keys()))):
        for label, store in [("图", app_figs), ("表", app_tbls)]:
            nums = store.get(letter, [])
            if not nums:
                continue
            expected = list(range(1, len(nums) + 1))
            if nums != expected:
                warnings.append(
                    f"  警告: 附录{letter}{label}编号不连续 — 期望 {expected}, 实际 {nums}")

    body_children = list(doc.element.body)
    last_table_number = None
    for idx, el in enumerate(body_children):
        if el.tag != qn("w:p"):
            continue

        text = _element_text(el)
        if not text:
            continue

        is_figure_caption = (
            re.match(fig_pat, text)
            or re.match(r"^Figure\s*\d", text, re.I)
            or re.match(r"^图[A-Z]\d+", text)
        )
        is_table_caption = (
            re.match(tbl_pat, text)
            or re.match(r"^Table\s*\d", text, re.I)
            or re.match(r"^(续)?表[A-Z]\d+", text)
        )

        if is_figure_caption:
            prev_idx = _find_prev_meaningful(body_children, idx)
            prev_el = body_children[prev_idx] if prev_idx is not None else None
            if prev_el is None or prev_el.tag != qn("w:p") or not prev_el.findall(".//" + qn("w:drawing")):
                warnings.append(f'  警告: 图题位置异常 — "{text}" 前未紧邻图片/绘图对象')
            continue

        if not is_table_caption:
            continue

        next_idx = _find_next_meaningful(body_children, idx)
        next_el = body_children[next_idx] if next_idx is not None else None
        if next_el is None or next_el.tag != qn("w:tbl"):
            warnings.append(f'  警告: 表题位置异常 — "{text}" 后未紧跟表格')

        if note_re.match(text) or source_re.match(text):
            continue

        current_number = _parse_body_caption_path(text)
        if re.match(r"^续", text):
            if last_table_number is None:
                warnings.append(f'  警告: 续表未找到可延续的上一表编号 — "{text}"')
            elif current_number != last_table_number:
                warnings.append(
                    f'  警告: 续表编号未延续上一表 — 上一表 {_format_caption_path(last_table_number)}, 当前 {_format_caption_path(current_number)}'
                )
        elif current_number is not None:
            last_table_number = current_number
    return warnings

def is_heading(para, level):
    return para.style and para.style.name == f"Heading {level}"


def contains_cjk(s):
    return any("\u4e00" <= ch <= "\u9fff" for ch in s)


def normalize_cn_keywords(text):
    m = re.match(r"^\s*关键词\s*[：:]\s*(.*)$", text)
    if not m:
        return None
    items = [x.strip(" ；;。,. ") for x in re.split(r"[；;]", m.group(1)) if x.strip(" ；;。,. ")]
    return "关键词：" + "；".join(items)


def cap_token(token):
    if "-" in token:
        parts = token.split("-")
        out = []
        for p in parts:
            if re.search(r"[A-Za-z]", p):
                out.append(p[:1].upper() + p[1:].lower())
            else:
                out.append(p)
        return "-".join(out)
    if re.search(r"[A-Za-z]", token):
        return token[:1].upper() + token[1:].lower()
    return token


def title_case_phrase(s):
    words = [w for w in re.split(r"\s+", s.strip()) if w]
    return " ".join([cap_token(w) for w in words])


def normalize_en_keywords(text):
    m = re.match(r"^\s*Key\s*words\s*:\s*(.*)$", text, flags=re.I)
    if not m:
        return None
    items = [x.strip(" ;；。,. ") for x in re.split(r"[;；]", m.group(1)) if x.strip(" ;；。,. ")]
    items = [title_case_phrase(x) for x in items]
    return "Key words: " + "; ".join(items)

