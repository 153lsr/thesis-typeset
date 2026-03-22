import copy
import re

from ._titles import _get_special_title_map
from ._common import get_paragraph_heading_level, is_heading_style
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


_CITE_NUM_RE = re.compile(r'\[(\d+(?:\s*[,\uff0c\-\u2013]\s*\d+)*)\]')
_CITE_AY_OUTER = re.compile(r'[\uff08(](.+?)[\uff09)]')
_CITE_AY_INNER = re.compile(r'(.+?)[,\uff0c]\s*((?:19|20)\d{2}[a-z]?)\s*$')
_REF_NUM_RE = re.compile(r'^\[(\d+)\]\s*')
_REF_TYPE_RE = re.compile(r'\[([A-Z]{1,2}(?:/[A-Z]{1,2})?)\]')
_REF_YEAR_RE = re.compile(r'(?:19|20)\d{2}[a-z]?')
_GBT_VALID_TYPES = {
    "J", "M", "C", "D", "R", "S", "P", "A", "Z", "N",
    "EB/OL", "OL", "DB/OL", "CP/DK", "DB", "CP",
}


def _parse_cite_numbers(inner):
    nums = []
    for part in re.split(r'[,\uff0c]', inner):
        part = part.strip()
        rm = re.match(r'(\d+)\s*[-\u2013]\s*(\d+)', part)
        if rm:
            nums.extend(range(int(rm.group(1)), int(rm.group(2)) + 1))
        elif re.match(r'\d+$', part):
            nums.append(int(part))
    return nums


def _extract_primary_author(author_str):
    return re.split(
        r'\u7b49|[\u548c\u4e0e&,\uff0c]|\s+and\s+|\s+et\s+al', author_str, maxsplit=1
    )[0].strip()


def check_citations(doc, cfg):
    warnings = []
    sec = cfg.get("sections", {})
    st_map = _get_special_title_map(cfg)

    ref_key = "\u53c2\u8003\u6587\u732e"
    if "\u53c2\u8003\u6587\u732e" in st_map:
        ref_key = st_map["\u53c2\u8003\u6587\u732e"]["match"]
    ref_key_norm = ref_key.replace(" ", "").replace("\u3000", "")

    chap_pat = re.compile(sec.get("chapter_pattern", r"^\u7b2c\s*\d+\s*\u7ae0"))

    paras = doc.paragraphs
    ref_start = ref_end = body_start = None

    _boundary_norms = set()
    for st in sec.get("special_titles", []):
        n = st["match"].replace(" ", "").replace("\u3000", "")
        if n != ref_key_norm:
            _boundary_norms.add(n)
    _ap = sec.get("appendix_pattern", r"^\u9644\u5f55\s*[A-Z]?")
    if _ap.endswith("[A-Z]"):
        _ap += "?"
    appendix_re = re.compile(_ap)

    for i, p in enumerate(paras):
        level = get_paragraph_heading_level(p)
        t_strip = p.text.strip()
        t_norm = t_strip.replace(" ", "").replace("\u3000", "")

        if level == 1 and body_start is None and chap_pat.match(t_strip):
            body_start = i
        if level == 1 and t_norm == ref_key_norm:
            ref_start = i + 1
        elif ref_start is not None and ref_end is None:
            if level is not None or (is_heading_style(p.style) and (
                    t_norm in _boundary_norms or appendix_re.match(t_strip))):
                ref_end = i

    if ref_start is None:
        return []
    if ref_end is None:
        ref_end = len(paras)
    if body_start is None:
        body_start = 0

    ref_entries = []
    for i in range(ref_start, ref_end):
        p = paras[i]
        level = get_paragraph_heading_level(p)
        t = p.text.strip()
        t_norm = t.replace(" ", "").replace("\u3000", "")

        if level is not None:
            break
        if t_norm in _boundary_norms or appendix_re.match(t):
            break
        if not t:
            continue

        entry = {"text": t, "idx": i}

        m = _REF_NUM_RE.match(t)
        entry["num"] = int(m.group(1)) if m else None
        t_body = t[m.end():] if m else t

        tm = _REF_TYPE_RE.search(t)
        entry["type"] = tm.group(1) if tm else None

        years = _REF_YEAR_RE.findall(t)
        entry["year"] = years[0] if years else None

        am = re.match(r'(.+?(?:\.[A-Z]\.)*)\.\s*(?=[^A-Z])', t_body)
        if not am:
            am = re.match(r'(.+?)\uff0e', t_body)
        entry["authors"] = am.group(1).strip() if am else t_body[:30].strip()

        ref_entries.append(entry)

    if not ref_entries:
        return []

    num_cites = []
    ay_cites = []
    in_appendix = False

    for i in range(body_start, ref_start - 1):
        p = paras[i]
        level = get_paragraph_heading_level(p)
        t_strip = p.text.strip()

        if level == 1:
            in_appendix = bool(appendix_re.match(t_strip))
        if level is not None or in_appendix:
            continue
        if not t_strip:
            continue

        for m in _CITE_NUM_RE.finditer(t_strip):
            for n in _parse_cite_numbers(m.group(1)):
                num_cites.append((n, i))

        for m in _CITE_AY_OUTER.finditer(t_strip):
            inner = m.group(1)
            for seg in re.split(r'[;\uff1b]', inner):
                seg = seg.strip()
                am = _CITE_AY_INNER.match(seg)
                if am:
                    author = am.group(1).strip()
                    if re.fullmatch(r'[\d\s\-\u2013\u2014\u5e74]+', author):
                        continue
                    ay_cites.append((author, am.group(2).strip(), i))

    style = "numbered" if len(num_cites) >= len(ay_cites) else "author-year"

    if style == "numbered":
        ref_nums = {e["num"]: e for e in ref_entries if e["num"] is not None}

        nums_list = [e["num"] for e in ref_entries if e["num"] is not None]
        if nums_list:
            expected = list(range(nums_list[0], nums_list[0] + len(nums_list)))
            if nums_list != expected:
                gaps = sorted(set(expected) - set(nums_list))
                if gaps:
                    warnings.append(f"\u53c2\u8003\u6587\u732e\u7f16\u53f7\u4e0d\u8fde\u7eed\uff0c\u7f3a\u5c11: {gaps}")
            seen = set()
            for n in nums_list:
                if n in seen:
                    warnings.append(f"\u53c2\u8003\u6587\u732e\u7f16\u53f7\u91cd\u590d: [{n}]")
                seen.add(n)

        first_seen = []
        for n, _ in num_cites:
            if n not in first_seen:
                first_seen.append(n)
        if first_seen and first_seen != sorted(first_seen):
            preview = first_seen[:15]
            warnings.append(
                f"\u6b63\u6587\u5f15\u7528\u7f16\u53f7\u672a\u6309\u9996\u6b21\u51fa\u73b0\u987a\u5e8f\u6392\u5217"
                f"\uff08\u524d{len(preview)}\u4e2a: {preview}\uff09"
            )

        cited_set = {n for n, _ in num_cites}
        ref_set = set(ref_nums.keys())
        diff_cite = sorted(cited_set - ref_set)
        diff_ref = sorted(ref_set - cited_set)
        if diff_cite:
            warnings.append(f"\u6b63\u6587\u5f15\u7528\u4e86\u4f46\u6587\u672b\u65e0\u5bf9\u5e94\u6761\u76ee: {diff_cite}")
        if diff_ref:
            warnings.append(f"\u6587\u672b\u6709\u6761\u76ee\u4f46\u6b63\u6587\u672a\u5f15\u7528: {diff_ref}")

    else:
        unmatched = []
        for author_str, year_str, _ in ay_cites:
            primary = _extract_primary_author(author_str)
            found = any(
                e["year"] and e["year"][:4] == year_str[:4]
                and primary and primary in e["authors"]
                for e in ref_entries
            )
            if not found:
                tag = f"\uff08{author_str}\uff0c{year_str}\uff09"
                if tag not in unmatched:
                    unmatched.append(tag)
        if unmatched:
            warnings.append(
                f"\u6b63\u6587\u5f15\u7528\u4e86\u4f46\u6587\u672b\u65e0\u5339\u914d\u6761\u76ee: {', '.join(unmatched[:15])}"
            )

        ref_ay = set()
        for e in ref_entries:
            if e["year"] and e["authors"]:
                ref_ay.add((_extract_primary_author(e["authors"]), e["year"][:4]))
        cited_ay = set()
        for a, y, _ in ay_cites:
            cited_ay.add((_extract_primary_author(a), y[:4]))
        uncited = ref_ay - cited_ay
        if uncited:
            tags = [f"{a}({y})" for a, y in sorted(uncited)]
            warnings.append(f"\u6587\u672b\u6709\u6761\u76ee\u4f46\u6b63\u6587\u672a\u5f15\u7528: {', '.join(tags[:15])}")

    for e in ref_entries:
        if e["type"] is None:
            warnings.append(f'\u53c2\u8003\u6587\u732e\u7f3a\u5c11\u7c7b\u578b\u6807\u8bc6[J]/[M]/..: "{e["text"][:50]}"')
        elif e["type"] not in _GBT_VALID_TYPES:
            warnings.append(f'\u53c2\u8003\u6587\u732e\u7c7b\u578b\u6807\u8bc6\u4e0d\u89c4\u8303[{e["type"]}]: "{e["text"][:50]}"')
        if not e["year"]:
            warnings.append(f'\u53c2\u8003\u6587\u732e\u7f3a\u5c11\u5e74\u4efd: "{e["text"][:50]}"')

    return warnings


def _make_text_run_el(text, rPr_el=None):
    r = OxmlElement('w:r')
    if rPr_el is not None:
        r.append(copy.deepcopy(rPr_el))
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    return r


def _make_field_runs(instr, display, rPr_el=None):
    els = []
    for ftype in ('begin', None, 'separate', None, 'end'):
        r = OxmlElement('w:r')
        if rPr_el is not None:
            r.append(copy.deepcopy(rPr_el))
        if ftype in ('begin', 'separate', 'end'):
            fc = OxmlElement('w:fldChar')
            fc.set(qn('w:fldCharType'), ftype)
            r.append(fc)
        elif len(els) == 1:
            it = OxmlElement('w:instrText')
            it.set(qn('xml:space'), 'preserve')
            it.text = f' {instr} '
            r.append(it)
        else:
            t = OxmlElement('w:t')
            t.set(qn('xml:space'), 'preserve')
            t.text = display
            r.append(t)
        els.append(r)
    return els


def _parse_cite_structure(inner):
    parts = []
    for seg in re.split(r'([,\uff0c])', inner):
        seg = seg.strip()
        if seg in (',', '\uff0c'):
            if parts:
                parts.append(('sep', ','))
            continue
        rm = re.match(r'(\d+)\s*[-\u2013]\s*(\d+)', seg)
        if rm:
            parts.append(('range', (int(rm.group(1)), int(rm.group(2)))))
        elif re.match(r'\d+$', seg):
            parts.append(('num', int(seg)))
    return parts


def _append_char_segment(p_el, chars):
    if not chars:
        return
    cur_rPr = chars[0][1]
    cur_text = ""
    for ch, rPr in chars:
        if rPr is cur_rPr:
            cur_text += ch
        else:
            if cur_text:
                p_el.append(_make_text_run_el(cur_text, cur_rPr))
            cur_rPr = rPr
            cur_text = ch
    if cur_text:
        p_el.append(_make_text_run_el(cur_text, cur_rPr))


def apply_ref_crosslinks(doc, cfg):
    sec = cfg.get("sections", {})
    st_map = _get_special_title_map(cfg)

    ref_key_norm = "\u53c2\u8003\u6587\u732e"
    if "\u53c2\u8003\u6587\u732e" in st_map:
        ref_key_norm = st_map["\u53c2\u8003\u6587\u732e"]["match"].replace(" ", "").replace("\u3000", "")

    chap_pat = re.compile(sec.get("chapter_pattern", r"^\u7b2c\s*\d+\s*\u7ae0"))
    _ap = sec.get("appendix_pattern", r"^\u9644\u5f55\s*[A-Z]?")
    if _ap.endswith("[A-Z]"):
        _ap += "?"
    appendix_re = re.compile(_ap)

    _boundary_norms = set()
    for st in sec.get("special_titles", []):
        n = st["match"].replace(" ", "").replace("\u3000", "")
        if n != ref_key_norm:
            _boundary_norms.add(n)

    paras = doc.paragraphs
    ref_start = ref_end = body_start = None

    for i, p in enumerate(paras):
        level = get_paragraph_heading_level(p)
        t_strip = p.text.strip()
        t_norm = t_strip.replace(" ", "").replace("\u3000", "")
        if level == 1 and body_start is None and chap_pat.match(t_strip):
            body_start = i
        if level == 1 and t_norm == ref_key_norm:
            ref_start = i + 1
        elif ref_start is not None and ref_end is None:
            if level is not None or (is_heading_style(p.style) and (
                    t_norm in _boundary_norms or appendix_re.match(t_strip))):
                ref_end = i

    if ref_start is None:
        return
    if ref_end is None:
        ref_end = len(paras)
    if body_start is None:
        body_start = 0

    num_count = ay_count = 0
    for i in range(body_start, ref_start - 1):
        t = paras[i].text
        num_count += len(_CITE_NUM_RE.findall(t))
        for m in _CITE_AY_OUTER.finditer(t):
            inner = m.group(1)
            for seg in re.split(r'[;\uff1b]', inner):
                if _CITE_AY_INNER.match(seg.strip()):
                    ay_count += 1
    is_numbered = not (ay_count > 0 and num_count == 0 and ay_count > num_count)

    bm_id = 1000
    bookmark_map = {}

    for i in range(ref_start, ref_end):
        p = paras[i]
        level = get_paragraph_heading_level(p)
        t = p.text.strip()
        if level is not None:
            break
        t_norm = t.replace(" ", "").replace("\u3000", "")
        if t_norm in _boundary_norms or appendix_re.match(t):
            break
        if not t:
            continue

        m = _REF_NUM_RE.match(t)
        if not m:
            continue

        num = int(m.group(1))
        bm_name = f"_Ref{num}"
        bookmark_map[num] = bm_name

        p_el = p._element
        runs = list(p.runs)
        if not runs:
            continue
        chars = []
        for r in runs:
            r_rPr = r._element.find(qn('w:rPr'))
            for ch in (r.text or ""):
                chars.append((ch, r_rPr))
        rPr0 = chars[0][1] if chars else None
        prefix_end = m.end()

        for child in list(p_el):
            if child.tag != qn('w:pPr'):
                p_el.remove(child)

        p_el.append(_make_text_run_el('[', rPr0))
        bm_start = OxmlElement('w:bookmarkStart')
        bm_start.set(qn('w:id'), str(bm_id))
        bm_start.set(qn('w:name'), bm_name)
        p_el.append(bm_start)
        for fel in _make_field_runs('SEQ Ref', str(num), rPr0):
            p_el.append(fel)
        bm_end = OxmlElement('w:bookmarkEnd')
        bm_end.set(qn('w:id'), str(bm_id))
        p_el.append(bm_end)
        p_el.append(_make_text_run_el('] ', rPr0))
        _append_char_segment(p_el, chars[prefix_end:])

        bm_id += 1

    if not bookmark_map:
        return

    if not is_numbered or num_count == 0:
        return
    in_appendix = False
    for i in range(body_start, ref_start - 1):
        p = paras[i]
        level = get_paragraph_heading_level(p)
        t_strip = p.text.strip()
        if level == 1:
            in_appendix = bool(appendix_re.match(t_strip))
        if level is not None or in_appendix or not t_strip:
            continue

        runs = list(p.runs)
        if not runs:
            continue
        chars = []
        for r in runs:
            r_rPr = r._element.find(qn('w:rPr'))
            for ch in (r.text or ""):
                chars.append((ch, r_rPr))
        full_text = "".join(c[0] for c in chars)

        matches = list(_CITE_NUM_RE.finditer(full_text))
        if not matches:
            continue

        has_valid = False
        for mat in matches:
            parts = _parse_cite_structure(mat.group(1))
            all_nums = []
            for pt in parts:
                if pt[0] == 'num':
                    all_nums.append(pt[1])
                elif pt[0] == 'range':
                    all_nums.extend(pt[1])
            if all(n in bookmark_map for n in all_nums):
                has_valid = True
                break
        if not has_valid:
            continue

        p_el = p._element
        for child in list(p_el):
            if child.tag != qn('w:pPr'):
                p_el.remove(child)

        pos = 0
        for mat in matches:
            if mat.start() > pos:
                _append_char_segment(p_el, chars[pos:mat.start()])

            parts = _parse_cite_structure(mat.group(1))
            all_nums = []
            for pt in parts:
                if pt[0] == 'num':
                    all_nums.append(pt[1])
                elif pt[0] == 'range':
                    all_nums.extend(pt[1])
            cite_rPr = chars[mat.start()][1]

            if all(n in bookmark_map for n in all_nums):
                p_el.append(_make_text_run_el('[', cite_rPr))
                for j, pt in enumerate(parts):
                    if pt[0] == 'sep':
                        p_el.append(_make_text_run_el(',', cite_rPr))
                    elif pt[0] == 'num':
                        bm = bookmark_map[pt[1]]
                        for fel in _make_field_runs(f'REF {bm} \\h', str(pt[1]), cite_rPr):
                            p_el.append(fel)
                    elif pt[0] == 'range':
                        bm_s = bookmark_map[pt[1][0]]
                        bm_e = bookmark_map[pt[1][1]]
                        for fel in _make_field_runs(f'REF {bm_s} \\h', str(pt[1][0]), cite_rPr):
                            p_el.append(fel)
                        p_el.append(_make_text_run_el('-', cite_rPr))
                        for fel in _make_field_runs(f'REF {bm_e} \\h', str(pt[1][1]), cite_rPr):
                            p_el.append(fel)
                p_el.append(_make_text_run_el(']', cite_rPr))
            else:
                _append_char_segment(p_el, chars[mat.start():mat.end()])

            pos = mat.end()

        if pos < len(chars):
            _append_char_segment(p_el, chars[pos:])
