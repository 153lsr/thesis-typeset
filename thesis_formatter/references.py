import copy
import re

from ._titles import _get_special_title_map
from ._common import get_paragraph_heading_level, is_heading_style, make_field_runs
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


_CITE_NUM_RE = re.compile(r'\[(\d+(?:\s*[,，\-–]\s*\d+)*)\]')
_CITE_AY_OUTER = re.compile(r'[（(](.+?)[）)]')
_CITE_AY_INNER = re.compile(r'(.+?)[,，]\s*((?:19|20)\d{2}[a-z]?)\s*$')
_REF_NUM_RE = re.compile(r'^\[(\d+)\]\s*')
_REF_TYPE_RE = re.compile(r'\[([A-Z]{1,2}(?:/[A-Z]{1,2})?)\]')
_REF_YEAR_RE = re.compile(r'(?:19|20)\d{2}[a-z]?')
_GBT_VALID_TYPES = {
    "J", "M", "C", "D", "R", "S", "P", "A", "Z", "N",
    "EB/OL", "OL", "DB/OL", "CP/DK", "DB", "CP",
}


def _parse_cite_numbers(inner):
    nums = []
    for part in re.split(r'[,，]', inner):
        part = part.strip()
        rm = re.match(r'(\d+)\s*[-–]\s*(\d+)', part)
        if rm:
            nums.extend(range(int(rm.group(1)), int(rm.group(2)) + 1))
        elif re.match(r'\d+$', part):
            nums.append(int(part))
    return nums


def _extract_primary_author(author_str):
    return re.split(
        r'等|[和与&,，]|\s+and\s+|\s+et\s+al', author_str, maxsplit=1
    )[0].strip()


def check_citations(doc, cfg):
    warnings = []
    sec = cfg.get("sections", {})
    st_map = _get_special_title_map(cfg)

    ref_key = "参考文献"
    if "参考文献" in st_map:
        ref_key = st_map["参考文献"]["match"]
    ref_key_norm = ref_key.replace(" ", "").replace("　", "")

    chap_pat = re.compile(sec.get("chapter_pattern", r"^第\s*\d+\s*章"))

    paras = doc.paragraphs
    ref_start = ref_end = body_start = None

    _boundary_norms = set()
    for st in sec.get("special_titles", []):
        n = st["match"].replace(" ", "").replace("　", "")
        if n != ref_key_norm:
            _boundary_norms.add(n)
    _ap = sec.get("appendix_pattern", r"^附录\s*[A-Z]?")
    if _ap.endswith("[A-Z]"):
        _ap += "?"
    appendix_re = re.compile(_ap)

    for i, p in enumerate(paras):
        level = get_paragraph_heading_level(p)
        t_strip = p.text.strip()
        t_norm = t_strip.replace(" ", "").replace("　", "")

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
        t_norm = t.replace(" ", "").replace("　", "")

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
            am = re.match(r'(.+?)．', t_body)
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

        if level is not None and appendix_re.match(t_strip):
            in_appendix = True
        elif level is not None:
            in_appendix = False
        if level is not None or in_appendix:
            continue
        if not t_strip:
            continue

        for m in _CITE_NUM_RE.finditer(t_strip):
            for n in _parse_cite_numbers(m.group(1)):
                num_cites.append((n, i))

        for m in _CITE_AY_OUTER.finditer(t_strip):
            inner = m.group(1)
            for seg in re.split(r'[;；]', inner):
                seg = seg.strip()
                am = _CITE_AY_INNER.match(seg)
                if am:
                    author = am.group(1).strip()
                    if re.fullmatch(r'[\d\s\-–—年]+', author):
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
                    warnings.append(f"参考文献编号不连续，缺少: {gaps}")
            seen = set()
            for n in nums_list:
                if n in seen:
                    warnings.append(f"参考文献编号重复: [{n}]")
                seen.add(n)

        first_seen = []
        for n, _ in num_cites:
            if n not in first_seen:
                first_seen.append(n)
        if first_seen and first_seen != sorted(first_seen):
            preview = first_seen[:15]
            warnings.append(
                f"正文引用编号未按首次出现顺序排列"
                f"（前{len(preview)}个: {preview}）"
            )

        cited_set = {n for n, _ in num_cites}
        ref_set = set(ref_nums.keys())
        diff_cite = sorted(cited_set - ref_set)
        diff_ref = sorted(ref_set - cited_set)
        if diff_cite:
            warnings.append(f"正文引用了但文末无对应条目: {diff_cite}")
        if diff_ref:
            warnings.append(f"文末有条目但正文未引用: {diff_ref}")

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
                tag = f"（{author_str}，{year_str}）"
                if tag not in unmatched:
                    unmatched.append(tag)
        if unmatched:
            warnings.append(
                f"正文引用了但文末无匹配条目: {', '.join(unmatched[:15])}"
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
            warnings.append(f"文末有条目但正文未引用: {', '.join(tags[:15])}")

    for e in ref_entries:
        if e["type"] is None:
            warnings.append(f'参考文献缺少类型标识[J]/[M]/..: "{e["text"][:50]}"')
        elif e["type"] not in _GBT_VALID_TYPES:
            warnings.append(f'参考文献类型标识不规范[{e["type"]}]: "{e["text"][:50]}"')
        if not e["year"]:
            warnings.append(f'参考文献缺少年份: "{e["text"][:50]}"')

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


def _parse_cite_structure(inner):
    parts = []
    for seg in re.split(r'([,，])', inner):
        seg = seg.strip()
        if seg in (',', '，'):
            if parts:
                parts.append(('sep', ','))
            continue
        rm = re.match(r'(\d+)\s*[-–]\s*(\d+)', seg)
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

    ref_key_norm = "参考文献"
    if "参考文献" in st_map:
        ref_key_norm = st_map["参考文献"]["match"].replace(" ", "").replace("　", "")

    chap_pat = re.compile(sec.get("chapter_pattern", r"^第\s*\d+\s*章"))
    _ap = sec.get("appendix_pattern", r"^附录\s*[A-Z]?")
    if _ap.endswith("[A-Z]"):
        _ap += "?"
    appendix_re = re.compile(_ap)

    _boundary_norms = set()
    for st in sec.get("special_titles", []):
        n = st["match"].replace(" ", "").replace("　", "")
        if n != ref_key_norm:
            _boundary_norms.add(n)

    paras = doc.paragraphs
    ref_start = ref_end = body_start = None

    for i, p in enumerate(paras):
        level = get_paragraph_heading_level(p)
        t_strip = p.text.strip()
        t_norm = t_strip.replace(" ", "").replace("　", "")
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
            for seg in re.split(r'[;；]', inner):
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
        t_norm = t.replace(" ", "").replace("　", "")
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
        for fel in make_field_runs('SEQ Ref', str(num), rPr0):
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
        if level is not None and appendix_re.match(t_strip):
            in_appendix = True
        elif level is not None:
            in_appendix = False
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
                        for fel in make_field_runs(f'REF {bm} \\h', str(pt[1]), cite_rPr):
                            p_el.append(fel)
                    elif pt[0] == 'range':
                        bm_s = bookmark_map[pt[1][0]]
                        bm_e = bookmark_map[pt[1][1]]
                        for fel in make_field_runs(f'REF {bm_s} \\h', str(pt[1][0]), cite_rPr):
                            p_el.append(fel)
                        p_el.append(_make_text_run_el('-', cite_rPr))
                        for fel in make_field_runs(f'REF {bm_e} \\h', str(pt[1][1]), cite_rPr):
                            p_el.append(fel)
                p_el.append(_make_text_run_el(']', cite_rPr))
            else:
                _append_char_segment(p_el, chars[mat.start():mat.end()])

            pos = mat.end()

        if pos < len(chars):
            _append_char_segment(p_el, chars[pos:])
