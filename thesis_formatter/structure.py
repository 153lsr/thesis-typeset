import re
import sys

from ._common import _check_caption_numbering, get_paragraph_heading_level, is_heading_style, _ALL_HEADING_NAMES
from ._titles import _find_special_display, _get_special_title_map
from .headings import _compile_section_patterns


def validate_structure(doc, cfg):
    warnings = []
    paras = doc.paragraphs
    texts = [p.text.strip() for p in paras]
    texts_nospace = [t.replace(" ", "").replace("　", "") for t in texts]

    sec = cfg.get("sections", {})
    st_map = _get_special_title_map(cfg)

    has_cn_abstract = any(t == "摘要" for t in texts_nospace)
    cn_kw_pat = sec.get("cn_keywords_pattern", r"关键词[：:]")
    has_cn_keywords = any(re.match(cn_kw_pat, t) for t in texts_nospace)
    en_abs_pat = sec.get("en_abstract_pattern", r"(?i)abstract[：:]?")
    has_en_abstract = any(re.match(en_abs_pat, t) for t in texts_nospace)
    en_kw_pat = sec.get("en_keywords_pattern", r"(?i)keywords?[：:]")
    has_en_keywords = any(re.match(en_kw_pat, t.replace(" ", "")) for t in texts)

    if not has_cn_abstract:
        warnings.append("缺少中文摘要标题")
    if not has_cn_keywords:
        warnings.append("缺少中文关键词")
    if not has_en_abstract:
        warnings.append("缺少英文摘要 (Abstract)")
    if not has_en_keywords:
        warnings.append("缺少英文关键词 (Key words)")

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
            warnings.append("英文摘要页缺少英文题目")
        if not has_affiliation:
            warnings.append("英文摘要页缺少作者英文名与单位信息")

    chapter_pat = sec.get("chapter_pattern", r"第\s*\d+\s*章")
    has_chapter_h1 = False
    ref_key = "参考文献"
    thanks_key = "致谢"
    if "参考文献" in st_map:
        ref_key = st_map["参考文献"]["match"]
    if "致谢" in st_map:
        thanks_key = st_map["致谢"]["match"]

    has_refs = any(t == ref_key.replace(" ", "").replace("　", "") for t in texts_nospace)
    has_thanks = any(t == thanks_key.replace(" ", "").replace("　", "") for t in texts_nospace)

    toc_key = normalize_title(_find_special_display(cfg, "目录", raw=True))
    has_toc = any(t == toc_key for t in texts_nospace)
    if not has_toc:
        warnings.append("缺少「目录」标题")

    cap_cfg = cfg.get("captions", {})
    fig_pat = cap_cfg.get("figure_pattern", r"^图\s*\d")
    tbl_pat = cap_cfg.get("table_pattern", r"^(续)?表\s*\d")
    has_images = any(
        el.tag.endswith("}blip") for el in doc.element.body.iter()
    )
    has_tables = len(doc.tables) > 0
    has_fig_cap = any(re.match(fig_pat, t) for t in texts)
    has_tbl_cap = any(re.match(tbl_pat, t) for t in texts)
    if has_images and not has_fig_cap:
        warnings.append("检测到插图但缺少图题（如「图1 xxx」）")
    if has_tables and not has_tbl_cap:
        warnings.append("检测到表格但缺少表题（如「表1 xxx」）")

    has_heading_styles = any(
        p.style and is_heading_style(p.style)
        for p in paras if p.text.strip())
    if not has_heading_styles:
        heading_examples = set()
        for s in doc.styles:
            if is_heading_style(s):
                heading_examples.add(s.name)
        examples = ", ".join(list(heading_examples)[:3]) if heading_examples else "Heading 1, Heading 2..."
        warnings.append(f"未检测到标题样式（请确保 Word 中已对标题应用 {examples} 样式）")

    for p in paras:
        level = get_paragraph_heading_level(p)
        t = p.text.strip()
        if level == 1 and re.match(chapter_pat, t):
            has_chapter_h1 = True
            break

    if not has_chapter_h1:
        warnings.append("未检测到正文章节标题")
    if not has_refs:
        warnings.append("缺少「参考文献」标题")
    if not has_thanks:
        warnings.append("缺少「致谢」标题")

    appendix_re, h2_pat, h3_pat, h4_pat = _compile_section_patterns(cfg)
    appendix_pat = appendix_re.pattern
    h1_pat = re.compile(f"({chapter_pat}|{appendix_pat})")

    special_h1_set = set(st_map.keys())
    special_h1_set.update(s.replace(" ", "").replace("　", "")
                          for s in sec.get("special_h1", []))

    for p in paras:
        level = get_paragraph_heading_level(p)
        t = p.text.strip()
        t_nospace = t.replace(" ", "").replace("　", "")
        if not t:
            continue

        if level == 1:
            if t_nospace not in special_h1_set and not h1_pat.match(t):
                warnings.append(f'一级标题缺少编号: "{t}"')
        elif level == 2:
            if not h2_pat.match(t):
                warnings.append(f'二级标题缺少编号: "{t}"')
        elif level == 3:
            if not h3_pat.match(t):
                warnings.append(f'三级标题缺少编号: "{t}"')
        elif level == 4:
            if not h4_pat.match(t):
                warnings.append(f'四级标题缺少编号: "{t}"')

    if warnings:
        print("=" * 50, file=sys.stderr)
        print("结构检查警告:", file=sys.stderr)
        for w in warnings:
            print(f"  ⚠ {w}", file=sys.stderr)
        print("=" * 50, file=sys.stderr)

    return warnings
