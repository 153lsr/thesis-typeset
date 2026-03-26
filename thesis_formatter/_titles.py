import re

from ._common import is_heading


_AR_CHAPTER_PATTERN = r"^第\s*\d+\s*章(?:\s|(?=[一-鿿])|$)"
_CN_CHAPTER_PATTERN = r"^第\s*[一二三四五六七八九十百千零两〇]+\s*章(?:\s|(?=[一-鿿])|$)"


def _find_special_display(cfg, match_text, raw=False):
    for st in cfg.get("special_titles", []):
        if st["match"] == match_text:
            return st["match"] if raw else st["display"]
    return match_text


def _get_special_title_map(cfg):
    result = {}
    for st in cfg.get("special_titles", []):
        key = st["match"].replace(" ", "").replace("　", "")
        result[key] = st
    return result


def _normalize_title(text):
    return text.replace(" ", "").replace("　", "")


def _iter_chapter_patterns(sec, text_first=False):
    configured = sec.get("chapter_pattern", _AR_CHAPTER_PATTERN)
    ordered = [_CN_CHAPTER_PATTERN, _AR_CHAPTER_PATTERN, configured] if text_first else [configured, _CN_CHAPTER_PATTERN, _AR_CHAPTER_PATTERN]
    seen = set()
    for pat in ordered:
        if pat and pat not in seen:
            seen.add(pat)
            yield pat


def _matches_chapter_heading(text, sec, text_first=False):
    return any(re.match(pat, text) for pat in _iter_chapter_patterns(sec, text_first=text_first))


def _detect_front_matter(doc, cfg):
    sec = cfg.get("sections", {})
    cn_kw_re = sec.get("cn_keywords_pattern", r"^\s*关键词\s*[：:]")
    en_abs_re = sec.get("en_abstract_pattern", r"(?i)^\s*Abstract\s*[：:]")
    text_first = bool(cfg.get("toc", {}).get("only_insert", False))
    appendix_re = re.compile(sec.get("appendix_pattern", r"^附录\s*[A-Z]"))

    skip_titles = {
        _normalize_title("摘要"),
        _normalize_title(_find_special_display(cfg, "摘要")),
        _normalize_title("Abstract"),
        _normalize_title(_find_special_display(cfg, "目录", raw=True)),
    }
    for entry in cfg.get("special_titles", []):
        match = entry.get("match", "")
        display = entry.get("display", "")
        if match:
            skip_titles.add(_normalize_title(match))
        if display:
            skip_titles.add(_normalize_title(display))
    for title in sec.get("special_h1", []):
        skip_titles.add(_normalize_title(title))

    headings = []
    for i, para in enumerate(doc.paragraphs):
        if not is_heading(para, 1):
            continue
        text = para.text.strip()
        if not text:
            continue
        headings.append((i, text, _normalize_title(text)))

    first_h1_idx = None
    for i, text, normalized in headings:
        if normalized in skip_titles or appendix_re.match(text):
            continue
        if _matches_chapter_heading(text, sec, text_first=text_first):
            first_h1_idx = i
            break

    if first_h1_idx is None:
        for i, text, normalized in headings:
            if normalized in skip_titles or appendix_re.match(text):
                continue
            first_h1_idx = i
            break

    if first_h1_idx is None:
        first_h1_idx = len(doc.paragraphs)
    if first_h1_idx == 0:
        return False

    for para in doc.paragraphs[:first_h1_idx]:
        t = para.text.strip()
        t_nospace = _normalize_title(t)
        if t_nospace == "摘要":
            return True
        if re.match(cn_kw_re, t):
            return True
        if re.match(en_abs_re, t) or re.match(r"(?i)^\s*Abstract\s*$", t):
            return True
    return False
