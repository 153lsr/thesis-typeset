from ._titles import _find_special_display
from ._common import get_paragraph_heading_level
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def insert_toc(doc, cfg):
    toc_match = _find_special_display(cfg, "\u76ee\u5f55", raw=True)
    toc_depth = cfg["toc"]["depth"]
    h1_font = cfg["fonts"]["h1"]
    h1_sz_hp = str(int(cfg["sizes"]["h1"] * 2))
    toc_cfg = cfg["toc"]
    toc_font = toc_cfg.get("font", cfg["fonts"]["body"])
    toc_sz_hp = str(int(toc_cfg.get("font_size", cfg["sizes"]["body"]) * 2))
    toc_ls = toc_cfg.get("line_spacing", cfg["body"]["line_spacing"])
    toc_ls_twips = str(int(toc_ls * 240))
    latin = cfg["fonts"]["latin"]

    first_h1_el = None
    for para in doc.paragraphs:
        if get_paragraph_heading_level(para) == 1:
            t = para.text.strip().replace(" ", "").replace("\u3000", "")
            if t == toc_match:
                para._element.getparent().remove(para._element)
                continue
            first_h1_el = para._element
            break

    if first_h1_el is None:
        return

    body = doc.element.body

    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for sdt in list(body.findall("w:sdt", ns)):
        body.remove(sdt)

    toc_display = _find_special_display(cfg, "\u76ee\u5f55")
    toc_title = OxmlElement("w:p")
    toc_title_ppr = OxmlElement("w:pPr")
    toc_title_jc = OxmlElement("w:jc")
    toc_title_jc.set(qn("w:val"), "center")
    toc_title_ppr.append(toc_title_jc)
    toc_title_spacing = OxmlElement("w:spacing")
    toc_title_spacing.set(qn("w:before"), "0")
    toc_title_spacing.set(qn("w:after"), "0")
    toc_title_spacing.set(qn("w:line"), "360")
    toc_title_spacing.set(qn("w:lineRule"), "auto")
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
    toc_field_spacing = OxmlElement("w:spacing")
    toc_field_spacing.set(qn("w:line"), toc_ls_twips)
    toc_field_spacing.set(qn("w:lineRule"), "auto")
    toc_field_ppr.append(toc_field_spacing)
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
    instr_text.text = f' TOC \\o "1-{toc_depth}" \\h \\z \\u '
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
    ph_text.text = "\u8bf7\u5728 Word \u4e2d\u53f3\u952e\u6b64\u5904 \u2192 \u66f4\u65b0\u57df \u2192 \u66f4\u65b0\u6574\u4e2a\u76ee\u5f55"
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
    toc_sz_hp = str(int(toc_cfg.get("font_size", cfg["sizes"]["body"]) * 2))
    toc_h1_font = toc_cfg.get("h1_font", cfg["fonts"]["h1"])
    toc_h1_sz_hp = str(int(toc_cfg.get("h1_font_size", cfg["sizes"]["h1"]) * 2))
    latin = cfg["fonts"]["latin"]

    styles_el = doc.styles.element
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    toc_depth = cfg["toc"]["depth"]

    for i in range(1, toc_depth + 1):
        style_id = f"TOC{i}"
        ea = toc_h1_font if i == 1 else toc_font
        sz_hp = toc_h1_sz_hp if i == 1 else toc_sz_hp
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
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "000000")
        rpr.append(color)
        style_el.append(rpr)

        ppr = OxmlElement("w:pPr")
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:line"), "360")
        spacing.set(qn("w:lineRule"), "auto")
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "0")
        ppr.append(spacing)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLine"), "0")
        if i > 1:
            ind.set(qn("w:left"), str((i - 1) * 240))
        ppr.append(ind)
        style_el.append(ppr)

        styles_el.append(style_el)
