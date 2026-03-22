import re

from ._common import (
    is_heading, get_paragraph_heading_level,
    _ALIGN_MAP, set_para_runs_font, set_run_font,
    _ALL_HEADING_NAMES, _HEADING_STYLE_IDS,
)
from ._titles import _find_special_display, _get_special_title_map
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


_CN_NUMERALS = "零一二三四五六七八九十"


def _int_to_cn(n):
    if n <= 10:
        return _CN_NUMERALS[n]
    if n < 20:
        return "十" + (_CN_NUMERALS[n - 10] if n > 10 else "")
    tens, ones = divmod(n, 10)
    return (_CN_NUMERALS[tens] if tens > 1 else "") + "十" + \
           (_CN_NUMERALS[ones] if ones else "")


def _renumber_h1_text(text, new_num, pattern):
    if re.search(r"第\s*\d+\s*章", text):
        return re.sub(r"(第\s*)\d+(\s*章)", fr"\g<1>{new_num}\2", text, count=1)
    if re.search(r"(?i)Chapter\s+\d+", text):
        return re.sub(r"(?i)(Chapter\s+)\d+", fr"\g<1>{new_num}", text, count=1)
    if re.match(r"^[一二三四五六七八九十百]+、", text):
        cn = _int_to_cn(new_num)
        return re.sub(r"^[一二三四五六七八九十百]+", cn, text, count=1)
    if re.match(r"^\d+\s", text):
        return re.sub(r"^\d+", str(new_num), text, count=1)
    return text


def _renumber_sub_text(text, prefix):
    if re.match(r"^（[一二三四五六七八九十百]+）", text):
        return re.sub(r"^（[一二三四五六七八九十百]+）", prefix, text, count=1)
    return re.sub(r"^[\d.]+", prefix, text, count=1)


def _replace_para_text(para, new_text):
    if para.runs:
        first_run = para.runs[0]
        font_props = {
            "name": first_run.font.name,
            "size": first_run.font.size,
            "bold": first_run.font.bold,
        }
        ea = first_run.font.element.find(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
        ea_name = ea.get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia") if ea is not None else None
    else:
        font_props = None
        ea_name = None

    para.text = new_text

    if font_props and para.runs:
        r = para.runs[0]
        r.font.name = font_props["name"]
        r.font.size = font_props["size"]
        r.font.bold = font_props["bold"]
        if ea_name:
            rFonts = r.font.element.find(qn("w:rFonts"))
            if rFonts is None:
                rFonts = OxmlElement("w:rFonts")
                r.font.element.insert(0, rFonts)
            rFonts.set(qn("w:eastAsia"), ea_name)


def renumber_headings(doc, cfg, skip_para_ids=None):
    skip_para_ids = set(skip_para_ids or ())
    sec = cfg.get("sections", {})
    chap_pat = re.compile(sec.get("chapter_pattern", r"^第\s*\d+\s*章\b"))
    appendix_pat = re.compile(sec.get("appendix_pattern", r"^附录\s*[A-Z]"))
    h2_pat = re.compile(sec.get("h2_pattern", r"^\d+\.\d+\s"))
    h3_pat = re.compile(sec.get("h3_pattern", r"^\d+\.\d+\.\d+\s"))
    h4_pat = re.compile(sec.get("h4_pattern", r"^\d+\.\d+\.\d+\.\d+\s"))

    st_map = _get_special_title_map(cfg)
    special_set = set(st_map.keys())
    special_set.update(s.replace(" ", "").replace("\u3000", "")
                       for s in sec.get("special_h1", []))

    changes = []
    chap_n = 0
    h2_n = h3_n = h4_n = 0
    in_appendix = False

    for para in doc.paragraphs:
        if id(para._element) in skip_para_ids:
            continue
        level = get_paragraph_heading_level(para)
        t = para.text.strip()
        t_nospace = t.replace(" ", "").replace("\u3000", "")
        if not t:
            continue

        if level == 1:
            if t_nospace in special_set:
                continue
            if appendix_pat.match(t):
                in_appendix = True
                continue
            if in_appendix:
                continue
            if chap_pat.match(t):
                chap_n += 1
                h2_n = h3_n = h4_n = 0
                new_t = _renumber_h1_text(t, chap_n, chap_pat.pattern)
                if new_t != t:
                    changes.append(f"  H1: \"{t}\" → \"{new_t}\"")
                    _replace_para_text(para, new_t)
        elif level == 2 and not in_appendix:
            if h2_pat.match(t):
                h2_n += 1
                h3_n = h4_n = 0
                new_t = _renumber_sub_text(t, f"{chap_n}.{h2_n}")
                if new_t != t:
                    changes.append(f"  H2: \"{t}\" → \"{new_t}\"")
                    _replace_para_text(para, new_t)
        elif level == 3 and not in_appendix:
            if h3_pat.match(t):
                h3_n += 1
                h4_n = 0
                new_t = _renumber_sub_text(t, f"{chap_n}.{h2_n}.{h3_n}")
                if new_t != t:
                    changes.append(f"  H3: \"{t}\" → \"{new_t}\"")
                    _replace_para_text(para, new_t)
        elif level == 4 and not in_appendix:
            if h4_pat.match(t):
                h4_n += 1
                new_t = _renumber_sub_text(t, f"{chap_n}.{h2_n}.{h3_n}.{h4_n}")
                if new_t != t:
                    changes.append(f"  H4: \"{t}\" → \"{new_t}\"")
                    _replace_para_text(para, new_t)

    return changes


def normalize_heading_spacing(doc, cfg, skip_para_ids=None):
    skip_para_ids = set(skip_para_ids or ())
    sec = cfg.get("sections", {})
    chap_re = re.compile(sec.get("chapter_pattern", r"^第\s*\d+\s*章\b"))
    st_map = _get_special_title_map(cfg)
    JIJU = "  "

    for para in doc.paragraphs:
        if id(para._element) in skip_para_ids:
            continue
        level = get_paragraph_heading_level(para)
        t = para.text.strip()
        if not t:
            continue
        t_nospace = t.replace(" ", "").replace("\u3000", "")

        new_t = None
        if level == 1:
            if t_nospace in st_map:
                continue
            m = re.match(r"(第\s*\d+\s*章)\s*(.*)", t)
            if m and m.group(2):
                new_t = m.group(1) + JIJU + m.group(2)
            if new_t is None:
                m = re.match(r"(Chapter\s+\d+)\s+(.*)", t, flags=re.IGNORECASE)
                if m and m.group(2):
                    new_t = m.group(1) + JIJU + m.group(2)
            if new_t is None:
                m = re.match(r"([一二三四五六七八九十百]+、)\s*(.*)", t)
                if m and m.group(2):
                    new_t = m.group(1) + JIJU + m.group(2)
            if new_t is None:
                m = re.match(r"(\d+)\s+(.*)", t)
                if m and m.group(2):
                    new_t = m.group(1) + JIJU + m.group(2)
        elif level in (2, 3, 4):
            m = re.match(r"(（[一二三四五六七八九十百]+）)\s*(.*)", t)
            if m and m.group(2):
                new_t = m.group(1) + JIJU + m.group(2)
            if new_t is None:
                m = re.match(r"([\d.]+)\s*(.*)", t)
                if m and m.group(2):
                    new_t = m.group(1) + JIJU + m.group(2)
            if new_t is None:
                m = re.match(r"(\(\d+\))\s*(.*)", t)
                if m and m.group(2):
                    new_t = m.group(1) + JIJU + m.group(2)

        if new_t is not None and new_t != t:
            _replace_para_text(para, new_t)


_SENTENCE_ENDINGS = set("。！？；")


def auto_assign_heading_styles(doc, cfg, skip_para_ids=None):
    skip_para_ids = set(skip_para_ids or ())
    sec = cfg.get("sections", {})
    chap_re = re.compile(sec.get("chapter_pattern", r"^第\s*\d+\s*章\b"))
    appendix_re = re.compile(sec.get("appendix_pattern", r"^附录\s*[A-Z]"))
    h2_re = re.compile(sec.get("h2_pattern", r"^\d+\.\d+(\s|(?=[\u4e00-\u9fff]))"))
    h3_re = re.compile(sec.get("h3_pattern", r"^\d+\.\d+\.\d+(\s|(?=[\u4e00-\u9fff]))"))
    h4_re = re.compile(sec.get("h4_pattern", r"^\d+\.\d+\.\d+\.\d+(\s|(?=[\u4e00-\u9fff]))"))

    st_map = _get_special_title_map(cfg)
    special_h1_set = set()
    for s in sec.get("special_h1", []):
        special_h1_set.add(s.replace(" ", "").replace("\u3000", ""))
    special_h1_set.update(st_map.keys())

    changes = []
    for para in doc.paragraphs:
        if id(para._element) in skip_para_ids:
            continue
        level = get_paragraph_heading_level(para)
        if level is not None:
            continue
        t = para.text.strip()
        if not t:
            continue
        t_nospace = t.replace(" ", "").replace("\u3000", "")

        target_level = None

        if chap_re.match(t):
            target_level = 1
        elif t_nospace in special_h1_set:
            target_level = 1
        elif appendix_re.match(t):
            target_level = 1
        elif re.match(r"^\d+(\s|(?=[\u4e00-\u9fff]))", t) and not re.match(r"^\d+\.\d+", t):
            if len(t) <= 50 and t[-1] not in _SENTENCE_ENDINGS:
                target_level = 1
        else:
            if len(t) <= 50 and t[-1] not in _SENTENCE_ENDINGS:
                if h4_re.match(t):
                    target_level = 4
                elif h3_re.match(t):
                    target_level = 3
                elif h2_re.match(t):
                    target_level = 2

        if target_level is not None:
            style_id = _HEADING_STYLE_IDS.get(target_level)
            if style_id:
                for style in doc.styles:
                    if style.style_id == style_id:
                        para.style = style
                        changes.append(f"  {style.name}: \"{t}\"")
                        break

    return changes
