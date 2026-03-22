import re

from ._common import is_heading


def _find_special_display(cfg, match_text, raw=False):
    for st in cfg.get("special_titles", []):
        if st["match"] == match_text:
            return st["match"] if raw else st["display"]
    return match_text


def _get_special_title_map(cfg):
    result = {}
    for st in cfg.get("special_titles", []):
        key = st["match"].replace(" ", "").replace("\u3000", "")
        result[key] = st
    return result


def _detect_front_matter(doc, cfg):
    sec = cfg.get("sections", {})
    cn_kw_re = sec.get("cn_keywords_pattern", r"^\s*关键词\s*[：:]")
    en_abs_re = sec.get("en_abstract_pattern", r"(?i)^\s*Abstract\s*[：:]")
    toc_display = _find_special_display(cfg, "目录", raw=True)

    first_h1_idx = None
    for i, para in enumerate(doc.paragraphs):
        if is_heading(para, 1):
            t = para.text.strip().replace(" ", "").replace("\u3000", "")
            if t == toc_display:
                continue
            first_h1_idx = i
            break

    if first_h1_idx is None or first_h1_idx == 0:
        return False

    for para in doc.paragraphs[:first_h1_idx]:
        t = para.text.strip()
        t_nospace = t.replace(" ", "").replace("\u3000", "")
        if t_nospace == "摘要":
            return True
        if re.match(cn_kw_re, t):
            return True
        if re.match(en_abs_re, t):
            return True
    return False
