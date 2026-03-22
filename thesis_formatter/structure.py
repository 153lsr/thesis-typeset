import re
import sys

from ._common import _check_caption_numbering, get_paragraph_heading_level, is_heading_style, _ALL_HEADING_NAMES
from ._titles import _find_special_display, _get_special_title_map


def validate_structure(doc, cfg):
    warnings = []
    paras = doc.paragraphs
    texts = [p.text.strip() for p in paras]
    texts_nospace = [t.replace(" ", "").replace("\u3000", "") for t in texts]

    sec = cfg["sections"]
    st_map = _get_special_title_map(cfg)

    has_cn_abstract = any(t == "\u6458\u8981" for t in texts_nospace)
    cn_kw_pat = sec.get("cn_keywords_pattern", r"\u5173\u952e\u8bcd[\uff1a:]")
    has_cn_keywords = any(re.match(cn_kw_pat, t) for t in texts_nospace)
    en_abs_pat = sec.get("en_abstract_pattern", r"(?i)abstract[\uff1a:]?")
    has_en_abstract = any(re.match(en_abs_pat, t) for t in texts_nospace)
    en_kw_pat = sec.get("en_keywords_pattern", r"(?i)keywords?[\uff1a:]")
    has_en_keywords = any(re.match(en_kw_pat, t.replace(" ", "")) for t in texts)

    if not has_cn_abstract:
        warnings.append("\u7f3a\u5c11\u4e2d\u6587\u6458\u8981\u6807\u9898")
    if not has_cn_keywords:
        warnings.append("\u7f3a\u5c11\u4e2d\u6587\u5173\u952e\u8bcd")
    if not has_en_abstract:
        warnings.append("\u7f3a\u5c11\u82f1\u6587\u6458\u8981 (Abstract)")
    if not has_en_keywords:
        warnings.append("\u7f3a\u5c11\u82f1\u6587\u5173\u952e\u8bcd (Key words)")

    cn_kw_idx = next((i for i, t in enumerate(texts_nospace)
                      if re.match(cn_kw_pat, t)), None)
    en_abs_idx = next((i for i, t in enumerate(texts_nospace)
                       if re.match(en_abs_pat, t)), None)
    if cn_kw_idx is not None and en_abs_idx is not None and cn_kw_idx < en_abs_idx:
        between = [texts[j] for j in range(cn_kw_idx + 1, en_abs_idx) if texts[j]]
        has_en_title = any(re.search(r"[A-Za-z]{4,}", t) for t in between)
        has_affiliation = any(re.search(r"(?i)(university|college|china|institute)", t)
                              for t in between)
        if not has_en_title:
            warnings.append("\u82f1\u6587\u6458\u8981\u9875\u7f3a\u5c11\u82f1\u6587\u9898\u76ee")
        if not has_affiliation:
            warnings.append("\u82f1\u6587\u6458\u8981\u9875\u7f3a\u5c11\u4f5c\u8005\u82f1\u6587\u540d\u4e0e\u5355\u4f4d\u4fe1\u606f")

    chapter_pat = sec.get("chapter_pattern", r"\u7b2c\s*\d+\s*\u7ae0")
    has_chapter_h1 = False
    ref_key = "\u53c2\u8003\u6587\u732e"
    thanks_key = "\u81f4\u8c22"
    if "\u53c2\u8003\u6587\u732e" in st_map:
        ref_key = st_map["\u53c2\u8003\u6587\u732e"]["match"]
    if "\u81f4\u8c22" in st_map:
        thanks_key = st_map["\u81f4\u8c22"]["match"]

    has_refs = any(t == ref_key.replace(" ", "").replace("\u3000", "") for t in texts_nospace)
    has_thanks = any(t == thanks_key.replace(" ", "").replace("\u3000", "") for t in texts_nospace)

    toc_key = _find_special_display(cfg, "\u76ee\u5f55", raw=True)
    has_toc = any(t == toc_key for t in texts_nospace)
    if not has_toc:
        warnings.append("\u7f3a\u5c11\u300c\u76ee\u5f55\u300d\u6807\u9898")

    cap_cfg = cfg.get("captions", {})
    fig_pat = cap_cfg.get("figure_pattern", r"^\u56fe\s*\d")
    tbl_pat = cap_cfg.get("table_pattern", r"^(\u7eed)?\u8868\s*\d")
    has_images = any(
        el.tag.endswith("}blip") for el in doc.element.body.iter()
    )
    has_tables = len(doc.tables) > 0
    has_fig_cap = any(re.match(fig_pat, t) for t in texts)
    has_tbl_cap = any(re.match(tbl_pat, t) for t in texts)
    if has_images and not has_fig_cap:
        warnings.append("\u68c0\u6d4b\u5230\u63d2\u56fe\u4f46\u7f3a\u5c11\u56fe\u9898\uff08\u5982\u300c\u56fe1 xxx\u300d\uff09")
    if has_tables and not has_tbl_cap:
        warnings.append("\u68c0\u6d4b\u5230\u8868\u683c\u4f46\u7f3a\u5c11\u8868\u9898\uff08\u5982\u300c\u88681 xxx\u300d\uff09")

    has_heading_styles = any(
        p.style and is_heading_style(p.style)
        for p in paras if p.text.strip())
    if not has_heading_styles:
        heading_examples = set()
        for s in doc.styles:
            if is_heading_style(s):
                heading_examples.add(s.name)
        examples = ", ".join(list(heading_examples)[:3]) if heading_examples else "Heading 1, Heading 2..."
        warnings.append(f"\u672a\u68c3\u6d4b\u5230\u6807\u9898\u6837\u5f0f\uff08\u8bf7\u786e\u4fdd Word \u4e2d\u5df2\u5bf9\u6807\u9898\u5e94\u7528 {examples} \u6837\u5f0f\uff09")

    for p in paras:
        level = get_paragraph_heading_level(p)
        t = p.text.strip()
        if level == 1 and re.match(chapter_pat, t):
            has_chapter_h1 = True
            break

    if not has_chapter_h1:
        warnings.append("\u672a\u68c3\u6d4b\u5230\u6b63\u6587\u7ae0\u8282\u6807\u9898")
    if not has_refs:
        warnings.append("\u7f3a\u5c11\u300c\u53c2\u8003\u6587\u732e\u300d\u6807\u9898")
    if not has_thanks:
        warnings.append("\u7f3a\u5c11\u300c\u81f4\u8c22\u300d\u6807\u9898")

    appendix_pat = sec.get("appendix_pattern", r"^\u9644\u5f55\s*[A-Z]")
    h1_pat = re.compile(f"({chapter_pat}|{appendix_pat})")
    h2_pat = re.compile(sec.get("h2_pattern", r"^\d+\.\d+\s"))
    h3_pat = re.compile(sec.get("h3_pattern", r"^\d+\.\d+\.\d+\s"))
    h4_pat = re.compile(sec.get("h4_pattern", r"^\d+\.\d+\.\d+\.\d+\s"))

    special_h1_set = set(st_map.keys())
    special_h1_set.update(s.replace(" ", "").replace("\u3000", "")
                          for s in sec.get("special_h1", []))

    for p in paras:
        level = get_paragraph_heading_level(p)
        t = p.text.strip()
        t_nospace = t.replace(" ", "").replace("\u3000", "")
        if not t:
            continue

        if level == 1:
            if t_nospace not in special_h1_set and not h1_pat.match(t):
                warnings.append(f'\u4e00\u7ea7\u6807\u9898\u7f3a\u5c11\u7f16\u53f7: "{t}"')
        elif level == 2:
            if not h2_pat.match(t):
                warnings.append(f'\u4e8c\u7ea7\u6807\u9898\u7f3a\u5c11\u7f16\u53f7: "{t}"')
        elif level == 3:
            if not h3_pat.match(t):
                warnings.append(f'\u4e09\u7ea7\u6807\u9898\u7f3a\u5c11\u7f16\u53f7: "{t}"')
        elif level == 4:
            if not h4_pat.match(t):
                warnings.append(f'\u56db\u7ea7\u6807\u9898\u7f3a\u5c11\u7f16\u53f7: "{t}"')

    if warnings:
        print("=" * 50, file=sys.stderr)
        print("\u7ed3\u6784\u68c0\u67e5\u8b66\u544a:", file=sys.stderr)
        for w in warnings:
            print(f"  \u26a0 {w}", file=sys.stderr)
        print("=" * 50, file=sys.stderr)

    return warnings
