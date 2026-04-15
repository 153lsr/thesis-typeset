from ._titles import _find_special_display
from ._common import get_paragraph_heading_level, parse_length, line_spacing_to_ooxml, paragraph_spacing_to_ooxml, normalize_title
from .headings import _get_paragraph_outline_level
from .page import find_first_body_heading
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _set_snap_to_grid(ppr, enabled):
    snap = ppr.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        ppr.insert(0, snap)
    snap.set(qn("w:val"), "1" if enabled else "0")


def _set_toc_level_indent(ppr, level):
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    for attr in ("w:left", "w:leftChars", "w:right", "w:rightChars", "w:firstLine", "w:firstLineChars", "w:hanging", "w:hangingChars"):
        key = qn(attr)
        if key in ind.attrib:
            del ind.attrib[key]
    ind.set(qn("w:firstLineChars"), "0")
    if level > 1:
        ind.set(qn("w:leftChars"), str((level - 1) * 100))


def _apply_spacing_value(spacing, side, value):
    spec = paragraph_spacing_to_ooxml(value)
    attr = qn(f"w:{side}")
    lines_attr = qn(f"w:{side}Lines")
    auto_attr = qn(f"w:{side}Autospacing")
    for key in (attr, lines_attr, auto_attr):
        if key in spacing.attrib:
            del spacing.attrib[key]
    if spec["mode"] == "lines":
        spacing.set(lines_attr, spec["value"])
    else:
        spacing.set(attr, spec["value"])


def _ensure_spacing(ppr, line_twips=None, line_rule=None, before_value=None, after_value=None):
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    if line_twips is not None:
        spacing.set(qn("w:line"), line_twips)
    if line_rule is not None:
        spacing.set(qn("w:lineRule"), line_rule)
    if before_value is not None:
        _apply_spacing_value(spacing, "before", before_value)
    if after_value is not None:
        _apply_spacing_value(spacing, "after", after_value)
    return spacing


def _set_on_off_property(rpr, tag_name, enabled):
    el = rpr.find(qn(f"w:{tag_name}"))
    if el is None:
        el = OxmlElement(f"w:{tag_name}")
        rpr.append(el)
    el.set(qn("w:val"), "1" if enabled else "0")


def _is_toc_sdt(sdt, ns):
    gallery = sdt.find(".//w:docPartGallery", ns)
    if gallery is not None and gallery.get(qn("w:val")) == "Table of Contents":
        return True

    for instr in sdt.findall(".//w:instrText", ns):
        if "TOC" in (instr.text or ""):
            return True
    return False


def insert_toc(doc, cfg):
    toc_match = _find_special_display(cfg, "目录", raw=True)
    toc_depth = cfg["toc"]["depth"]
    h1_font = cfg["fonts"]["h1"]
    h1_sz_hp = str(int(cfg["sizes"]["h1"] * 2))
    toc_cfg = cfg["toc"]
    toc_font = toc_cfg.get("font", cfg["fonts"]["body"])
    toc_font_size = toc_cfg.get("font_size", cfg["sizes"]["body"])
    toc_bold = toc_cfg.get("bold", False)
    toc_sz_hp = str(int(toc_font_size * 2))
    toc_ls = toc_cfg.get("line_spacing", cfg["body"]["line_spacing"])
    toc_ls_twips, toc_ls_rule = line_spacing_to_ooxml(toc_ls)
    toc_sb_value = toc_cfg.get("space_before", 0)
    toc_sa_value = toc_cfg.get("space_after", 0)
    latin = cfg["fonts"]["latin"]

    first_h1_el = None
    toc_match_normalized = normalize_title(toc_match)
    for para in list(doc.paragraphs):
        level = get_paragraph_heading_level(para)
        if level is None:
            olvl = _get_paragraph_outline_level(para)
            if olvl is not None:
                level = olvl
        if level != 1:
            if not para.text.strip():
                continue
            t_nospace = normalize_title(para.text.strip())
            if t_nospace != toc_match_normalized:
                continue
            # Match by text even without heading level (e.g., toc_only mode)
            para._element.getparent().remove(para._element)
            continue
        t = para.text.strip().replace(" ", "").replace("\u3000", "")
        if t == toc_match:
            para._element.getparent().remove(para._element)

    first_body_heading = find_first_body_heading(doc, cfg)
    if first_body_heading is not None:
        first_h1_el = first_body_heading._element
    else:
        for para in doc.paragraphs:
            if get_paragraph_heading_level(para) == 1:
                first_h1_el = para._element
                break

    if first_h1_el is None:
        return

    body = doc.element.body

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for sdt in list(body.findall("w:sdt", ns)):
        if _is_toc_sdt(sdt, ns):
            body.remove(sdt)

    toc_display = _find_special_display(cfg, "目录")
    toc_title = OxmlElement("w:p")
    toc_title_ppr = OxmlElement("w:pPr")
    toc_title_jc = OxmlElement("w:jc")
    toc_title_jc.set(qn("w:val"), "center")
    toc_title_ppr.append(toc_title_jc)
    _set_snap_to_grid(toc_title_ppr, False)
    toc_title_spacing = OxmlElement("w:spacing")
    toc_title_spacing.set(qn("w:before"), "0")
    toc_title_spacing.set(qn("w:after"), "0")
    toc_title_spacing.set(qn("w:line"), toc_ls_twips)
    toc_title_spacing.set(qn("w:lineRule"), toc_ls_rule)
    toc_title_ppr.append(toc_title_spacing)
    toc_title.append(toc_title_ppr)

    toc_title_run = OxmlElement("w:r")
    toc_title_rpr = OxmlElement("w:rPr")
    toc_title_rfonts = OxmlElement("w:rFonts")
    toc_title_rfonts.set(qn("w:ascii"), latin)
    toc_title_rfonts.set(qn("w:hAnsi"), latin)
    toc_title_rfonts.set(qn("w:eastAsia"), h1_font)
    toc_title_rpr.append(toc_title_rfonts)
    toc_title_sz = OxmlElement("w:sz")
    toc_title_sz.set(qn("w:val"), h1_sz_hp)
    toc_title_rpr.append(toc_title_sz)
    toc_title_szCs = OxmlElement("w:szCs")
    toc_title_szCs.set(qn("w:val"), h1_sz_hp)
    toc_title_rpr.append(toc_title_szCs)
    toc_title_bold = OxmlElement("w:b")
    toc_title_rpr.append(toc_title_bold)
    toc_title_run.append(toc_title_rpr)
    toc_title_text = OxmlElement("w:t")
    toc_title_text.set(qn("xml:space"), "preserve")
    toc_title_text.text = toc_display
    toc_title_run.append(toc_title_text)
    toc_title.append(toc_title_run)

    toc_field = OxmlElement("w:p")
    toc_field_ppr = OxmlElement("w:pPr")
    _set_snap_to_grid(toc_field_ppr, False)
    _ensure_spacing(toc_field_ppr, line_twips=toc_ls_twips, line_rule=toc_ls_rule, before_value=toc_sb_value, after_value=toc_sa_value)
    toc_field.append(toc_field_ppr)

    run_begin = OxmlElement("w:r")
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run_begin.append(fld_begin)
    toc_field.append(run_begin)

    run_instr = OxmlElement("w:r")
    instr_rpr = OxmlElement("w:rPr")
    instr_rfonts = OxmlElement("w:rFonts")
    instr_rfonts.set(qn("w:ascii"), latin)
    instr_rfonts.set(qn("w:hAnsi"), latin)
    instr_rfonts.set(qn("w:eastAsia"), toc_font)
    instr_rpr.append(instr_rfonts)
    instr_sz = OxmlElement("w:sz")
    instr_sz.set(qn("w:val"), toc_sz_hp)
    instr_rpr.append(instr_sz)
    run_instr.append(instr_rpr)
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = f' TOC \\o "1-{toc_depth}" \\h \\z '
    run_instr.append(instr_text)
    toc_field.append(run_instr)

    run_sep = OxmlElement("w:r")
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run_sep.append(fld_sep)
    toc_field.append(run_sep)

    run_placeholder = OxmlElement("w:r")
    ph_rpr = OxmlElement("w:rPr")
    ph_rfonts = OxmlElement("w:rFonts")
    ph_rfonts.set(qn("w:ascii"), latin)
    ph_rfonts.set(qn("w:hAnsi"), latin)
    ph_rfonts.set(qn("w:eastAsia"), toc_font)
    ph_rpr.append(ph_rfonts)
    ph_sz = OxmlElement("w:sz")
    ph_sz.set(qn("w:val"), toc_sz_hp)
    ph_rpr.append(ph_sz)
    run_placeholder.append(ph_rpr)
    ph_text = OxmlElement("w:t")
    ph_text.text = "请在 Word 中右键此处 → 更新域 → 更新整个目录"
    run_placeholder.append(ph_text)
    toc_field.append(run_placeholder)

    run_end = OxmlElement("w:r")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run_end.append(fld_end)
    toc_field.append(run_end)

    page_break = OxmlElement("w:p")
    pb_run = OxmlElement("w:r")
    br_el = OxmlElement("w:br")
    br_el.set(qn("w:type"), "page")
    pb_run.append(br_el)
    page_break.append(pb_run)

    body.insert(list(body).index(first_h1_el), page_break)
    body.insert(list(body).index(page_break), toc_field)
    body.insert(list(body).index(toc_field), toc_title)


def ensure_toc_styles(doc, cfg):
    toc_cfg = cfg["toc"]
    toc_font = toc_cfg.get("font", cfg["fonts"]["body"])
    toc_font_size = toc_cfg.get("font_size", cfg["sizes"]["body"])
    toc_bold = toc_cfg.get("bold", False)
    toc_sz_hp = str(int(toc_font_size * 2))
    toc_h1_font = toc_cfg.get("h1_font", cfg["fonts"]["h1"])
    toc_h1_font_size = toc_cfg.get("h1_font_size", cfg["sizes"]["h1"])
    toc_h1_bold = toc_cfg.get("h1_bold", False)
    toc_h1_sz_hp = str(int(toc_h1_font_size * 2))
    toc_ls_twips, toc_ls_rule = line_spacing_to_ooxml(toc_cfg.get("line_spacing", cfg["body"]["line_spacing"]))
    latin = cfg["fonts"]["latin"]

    styles_el = doc.styles.element
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    toc_depth = cfg["toc"]["depth"]

    for i in range(1, toc_depth + 1):
        style_id = f"TOC{i}"
        ea = toc_h1_font if i == 1 else toc_font
        sz_hp = toc_h1_sz_hp if i == 1 else toc_sz_hp
        level_bold = toc_h1_bold if i == 1 else toc_bold
        toc_sb_value = toc_cfg.get("space_before", 0)
        toc_sa_value = toc_cfg.get("space_after", 0)
        found = styles_el.find(f'.//w:style[@w:styleId="{style_id}"]', ns)
        if found is not None:
            rpr = found.find("w:rPr", ns)
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                found.append(rpr)
            rfonts = rpr.find("w:rFonts", ns)
            if rfonts is None:
                rfonts = OxmlElement("w:rFonts")
                rpr.append(rfonts)
            for theme in ["w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"]:
                if rfonts.get(qn(theme)) is not None:
                    del rfonts.attrib[qn(theme)]
            rfonts.set(qn("w:ascii"), latin)
            rfonts.set(qn("w:hAnsi"), latin)
            rfonts.set(qn("w:eastAsia"), ea)
            sz = rpr.find("w:sz", ns)
            if sz is None:
                sz = OxmlElement("w:sz")
                rpr.append(sz)
            sz.set(qn("w:val"), sz_hp)
            szCs = rpr.find("w:szCs", ns)
            if szCs is None:
                szCs = OxmlElement("w:szCs")
                rpr.append(szCs)
            szCs.set(qn("w:val"), sz_hp)
            _set_on_off_property(rpr, "b", level_bold)
            _set_on_off_property(rpr, "bCs", level_bold)
            ppr = found.find("w:pPr", ns)
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                found.append(ppr)
            _set_snap_to_grid(ppr, False)
            _ensure_spacing(ppr, line_twips=toc_ls_twips, line_rule=toc_ls_rule, before_value=toc_sb_value, after_value=toc_sa_value)
            _set_toc_level_indent(ppr, i)
            continue

        style_el = OxmlElement("w:style")
        style_el.set(qn("w:type"), "paragraph")
        style_el.set(qn("w:styleId"), style_id)

        name_el = OxmlElement("w:name")
        name_el.set(qn("w:val"), f"toc {i}")
        style_el.append(name_el)

        based = OxmlElement("w:basedOn")
        based.set(qn("w:val"), "Normal")
        style_el.append(based)

        nxt = OxmlElement("w:next")
        nxt.set(qn("w:val"), "Normal")
        style_el.append(nxt)

        ui = OxmlElement("w:uiPriority")
        ui.set(qn("w:val"), "39")
        style_el.append(ui)

        unhide = OxmlElement("w:unhideWhenUsed")
        style_el.append(unhide)

        rpr = OxmlElement("w:rPr")
        rfonts = OxmlElement("w:rFonts")
        rfonts.set(qn("w:ascii"), latin)
        rfonts.set(qn("w:hAnsi"), latin)
        rfonts.set(qn("w:eastAsia"), ea)
        rpr.append(rfonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), sz_hp)
        rpr.append(sz)
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), sz_hp)
        rpr.append(szCs)
        _set_on_off_property(rpr, "b", level_bold)
        _set_on_off_property(rpr, "bCs", level_bold)
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "000000")
        rpr.append(color)
        style_el.append(rpr)

        ppr = OxmlElement("w:pPr")
        _set_snap_to_grid(ppr, False)
        _ensure_spacing(ppr, line_twips=toc_ls_twips, line_rule=toc_ls_rule, before_value=toc_sb_value, after_value=toc_sa_value)
        _set_toc_level_indent(ppr, i)
        style_el.append(ppr)

        styles_el.append(style_el)


