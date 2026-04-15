"""Post-process formatted docx via Word COM (Python win32com).

Handles:
- Update TOC field
- Fix TOC entry fonts (宋体 + TNR, 小四, not bold)
"""

import argparse
import ctypes
import os
import re
import subprocess
import sys
import threading
from ctypes import wintypes

import pythoncom
import win32com.client as win32

from thesis_formatter._common import parse_length, paragraph_spacing_to_word

wdAlertsNone = 0
wdColorBlack = 0
wdLineSpaceMultiple = 5
msoAutomationSecurityForceDisable = 3


class PostprocessError(RuntimeError):
    """Raised when Word COM post-processing fails."""


class PostprocessTimeoutError(PostprocessError):
    """Raised when Word COM post-processing exceeds the timeout."""


def _get_process_id_from_hwnd(hwnd):
    """Return the process id that owns *hwnd*, or None if unavailable."""
    if not hwnd:
        return None
    pid = wintypes.DWORD()
    thread_id = ctypes.windll.user32.GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid))
    if not thread_id or not pid.value:
        return None
    return int(pid.value)


def _terminate_process(pid, timeout=5):
    """Force-kill a specific process id without touching unrelated Word instances."""
    if not pid:
        return False
    try:
        result = subprocess.run(
            ["taskkill", "/F", "/PID", str(int(pid))],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except Exception:
        return False
    return result.returncode == 0


def _apply_word_spacing(fmt, side, value):
    spec = paragraph_spacing_to_word(value)
    side_cap = side[0].upper() + side[1:]
    line_unit_attr = f"LineUnit{side_cap}"
    space_attr = f"Space{side_cap}"
    if spec["mode"] == "lines":
        # Clear inherited point spacing first; Word otherwise keeps values like 10pt when line-units are zero.
        setattr(fmt, space_attr, 0)
        setattr(fmt, line_unit_attr, float(spec["value"]))
    else:
        setattr(fmt, line_unit_attr, 0)
        setattr(fmt, space_attr, float(spec["value"]))


def _apply_three_line(tbl, top_sz, header_sz, bottom_sz):
    """Apply three-line table borders via Word COM."""
    nr = tbl.Rows.Count
    nc = tbl.Columns.Count
    tbl.Borders.Enable = False
    for ci in range(1, nc + 1):
        try:
            c = tbl.Cell(1, ci)
            b = c.Borders(-1)
            b.LineStyle = 1
            b.LineWidth = top_sz
            b2 = c.Borders(-3)
            b2.LineStyle = 1
            b2.LineWidth = header_sz if nr > 1 else bottom_sz
        except Exception:
            pass
    if nr > 1:
        for ci in range(1, nc + 1):
            try:
                c = tbl.Cell(nr, ci)
                b = c.Borders(-3)
                b.LineStyle = 1
                b.LineWidth = bottom_sz
            except Exception:
                pass


def _split_spanning_tables(doc, config, log):
    """Split tables that span page breaks into separate tables with 续表 captions."""
    wdActiveEndPageNumber = 3

    cap_cfg = config.get("captions", {}) if config else {}
    tbl_pat = re.compile(cap_cfg.get("table_pattern", r"^(续)?表\s*\d"))
    fonts_cfg = config.get("fonts", {}) if config else {}
    ea_font = fonts_cfg.get("body", "宋体")
    lat_font = fonts_cfg.get("latin", "Times New Roman")
    cap_size = (config.get("sizes", {}) if config else {}).get("caption", 10.5)
    tbl_cfg = config.get("table", {}) if config else {}
    top_sz = tbl_cfg.get("top_border_sz", 12)
    header_sz = tbl_cfg.get("header_border_sz", 8)
    bottom_sz = tbl_cfg.get("bottom_border_sz", 12)

    cover_sections = int((config or {}).get("_runtime", {}).get("custom_cover_sections", 0) or 0)

    total = 0
    for _ in range(20):
        did_split = False
        for ti in range(1, doc.Tables.Count + 1):
            try:
                tbl = doc.Tables(ti)
            except Exception:
                continue
            if cover_sections > 0:
                try:
                    sec_idx = tbl.Range.Sections(1).Index
                    if sec_idx <= cover_sections:
                        continue
                except Exception:
                    pass
            nr, nc = tbl.Rows.Count, tbl.Columns.Count
            if nr < 4:
                continue
            try:
                p1 = tbl.Rows(1).Range.Information(wdActiveEndPageNumber)
                pn = tbl.Rows(nr).Range.Information(wdActiveEndPageNumber)
            except Exception:
                continue
            if p1 <= 0 or pn <= 0 or p1 == pn:
                continue

            sr = None
            for ri in range(2, nr):
                try:
                    rp = tbl.Rows(ri).Range.Information(wdActiveEndPageNumber)
                except Exception:
                    continue
                if rp > p1:
                    sr = ri
                    break
            if not sr or sr < 3:
                continue

            cap_text = ""
            try:
                r = doc.Range(tbl.Range.Start, tbl.Range.Start)
                r.MoveStart(4, -1)
                t = r.Text.strip().replace("\r", "").replace("\x07", "")
                if tbl_pat.match(t):
                    cap_text = t if t.startswith("续") else "续" + t
            except Exception:
                pass
            if not cap_text:
                cap_text = "续表"

            log(f"  table {ti}: split row {sr}/{nr}, page {p1}->{pn}")

            try:
                ins = tbl.Range.Duplicate
                ins.Collapse(0)
                ins.InsertAfter(cap_text + "\r")
                ins.Collapse(0)

                overflow = nr - sr + 1
                nt = doc.Tables.Add(ins, 1 + overflow, nc)

                for ci in range(1, nc + 1):
                    try:
                        src = tbl.Cell(1, ci).Range.Duplicate
                        src.MoveEnd(1, -1)
                        nt.Cell(1, ci).Range.FormattedText = src.FormattedText
                    except Exception:
                        pass

                for oi in range(sr, nr + 1):
                    ni = oi - sr + 2
                    for ci in range(1, nc + 1):
                        try:
                            src = tbl.Cell(oi, ci).Range.Duplicate
                            src.MoveEnd(1, -1)
                            nt.Cell(ni, ci).Range.FormattedText = src.FormattedText
                        except Exception:
                            pass

                for ri in range(nr, sr - 1, -1):
                    try:
                        tbl.Rows(ri).Delete()
                    except Exception:
                        pass

                # Remove phantom columns that Word may have added
                attempts = 0
                while nt.Columns.Count > nc and attempts < nc + 5:
                    try:
                        nt.Columns(nt.Columns.Count).Delete()
                    except Exception:
                        break
                    attempts += 1

                try:
                    cr = doc.Range(nt.Range.Start, nt.Range.Start)
                    cr.MoveStart(4, -1)
                    cr.ParagraphFormat.Alignment = 1
                    cr.Font.Name = lat_font
                    cr.Font.NameFarEast = ea_font
                    cr.Font.Size = cap_size
                    cr.Font.Bold = False
                except Exception:
                    pass

                try:
                    nt.PreferredWidthType = tbl.PreferredWidthType
                    nt.PreferredWidth = tbl.PreferredWidth
                except Exception:
                    pass

                _apply_three_line(tbl, top_sz, header_sz, bottom_sz)
                _apply_three_line(nt, top_sz, header_sz, bottom_sz)

                total += 1
                did_split = True
                break
            except Exception as e:
                log(f"  split failed: {e}")
        if not did_split:
            break

    if total:
        log(f"  {total} table(s) split across page breaks.")
    return total


def postprocess(docx_path, timeout=90, config=None, mode="full", log=None):
    if log is None:
        log = print
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        raise PostprocessError(f"File not found: {docx_path}")
    if mode not in {"full", "fields_only"}:
        raise PostprocessError(f"Unsupported postprocess mode: {mode}")

    if config:
        toc_cfg = config.get("toc", {})
        fonts_cfg = config.get("fonts", {})
        sizes_cfg = config.get("sizes", {})
        toc_latin = fonts_cfg.get("latin", "Times New Roman")
        toc_ea = toc_cfg.get("font", fonts_cfg.get("body", "宋体"))
        toc_size = toc_cfg.get("font_size", sizes_cfg.get("body", 12))
        toc_h1_ea = toc_cfg.get("h1_font", fonts_cfg.get("h1", toc_ea))
        toc_h1_size = toc_cfg.get("h1_font_size", sizes_cfg.get("h1", toc_size))
        toc_line_spacing = toc_cfg.get("line_spacing", 1.5)
        toc_space_before_cfg = toc_cfg.get("space_before", 0)
        toc_space_after_cfg = toc_cfg.get("space_after", 0)
    else:
        toc_latin = "Times New Roman"
        toc_ea = "宋体"
        toc_size = 12
        toc_h1_ea = toc_ea
        toc_h1_size = toc_size
        toc_line_spacing = 1.5
        toc_space_before_cfg = 0
        toc_space_after_cfg = 0
    # 构建特殊标题映射：display 原文（含全角空格）→ match 原文（无空格）
    special_toc_map = {}
    if config:
        for st in config.get("special_titles", []):
            match_text = st.get("match", "")
            display_text = st.get("display", match_text)
            if display_text != match_text:
                special_toc_map[display_text] = match_text

    result = {"ok": False, "error": None, "pid": None}
    done_event = threading.Event()

    def worker():
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = wdAlertsNone
            word.AutomationSecurity = msoAutomationSecurityForceDisable
            word.Options.DoNotPromptForConvert = True
            try:
                result["pid"] = _get_process_id_from_hwnd(word.Hwnd)
            except Exception:
                result["pid"] = None

            log("[1/3] Opening document...")
            doc = word.Documents.Open(
                docx_path,
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            log("[1/3] Done.")

            if mode == "full":
                log("[2/3] Updating TOC and fields...")
                for toc in doc.TablesOfContents:
                    toc.Update()
                doc.Fields.Update()
                # 在 TOC Range 内用 Find/Replace 去除特殊标题的全角空格
                if special_toc_map:
                    for toc in doc.TablesOfContents:
                        for display_text, match_text in special_toc_map.items():
                            toc_find = toc.Range.Find
                            toc_find.ClearFormatting()
                            toc_find.Replacement.ClearFormatting()
                            toc_find.Text = display_text
                            toc_find.Replacement.Text = match_text
                            toc_find.Forward = True
                            toc_find.Wrap = 0  # wdFindStop
                            toc_find.MatchCase = True
                            toc_find.MatchWholeWord = False
                            toc_find.Execute(Replace=2)  # wdReplaceAll
                log("[2/3] Done.")

                log(f"[3/3] Fixing TOC fonts")
                seen_toc_styles = set()
                for toc in doc.TablesOfContents:
                    for p in toc.Range.Paragraphs:
                        try:
                            sname = p.Style.NameLocal
                        except Exception:
                            sname = ""
                        level = 0
                        m = re.search(r"(\d+)\s*$", str(sname))
                        if m:
                            level = int(m.group(1))
                        is_level1 = level == 1
                        level_font_size = toc_h1_size if is_level1 else toc_size

                        style_obj = p.Style
                        style_fmt = style_obj.ParagraphFormat
                        style_name = str(sname)
                        p.Range.Font.Name = toc_latin
                        p.Range.Font.NameFarEast = toc_h1_ea if is_level1 else toc_ea
                        p.Range.Font.Size = level_font_size
                        p.Range.Font.Bold = False
                        p.Range.Font.ColorIndex = wdColorBlack
                        try:
                            p.Format.DisableLineHeightGrid = True
                        except Exception:
                            pass
                        try:
                            style_fmt.DisableLineHeightGrid = True
                        except Exception:
                            pass
                        if style_name not in seen_toc_styles:
                            _apply_word_spacing(style_fmt, "before", toc_space_before_cfg)
                            _apply_word_spacing(style_fmt, "after", toc_space_after_cfg)
                            seen_toc_styles.add(style_name)
                        p.Format.LineSpacingRule = style_fmt.LineSpacingRule
                        p.Format.LineSpacing = style_fmt.LineSpacing
                        _apply_word_spacing(p.Format, "before", toc_space_before_cfg)
                        _apply_word_spacing(p.Format, "after", toc_space_after_cfg)
                log("[3/3] Done.")

                _split_spanning_tables(doc, config, log)

            else:
                log("[2/2] Updating fields...")
                doc.Fields.Update()
                log("[2/2] Done.")

            doc.Save()
            doc.Close()
            result["ok"] = True

        except Exception as exc:
            result["error"] = str(exc)
        finally:
            if word:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
            done_event.set()

    t = threading.Thread(target=worker, daemon=True)
    t.start()
    finished = done_event.wait(timeout=timeout)

    if not finished:
        pid = result.get("pid")
        if _terminate_process(pid):
            raise PostprocessTimeoutError(
                f"TIMEOUT after {timeout}s; terminated Word PID {pid}"
            )
        raise PostprocessTimeoutError(
            f"TIMEOUT after {timeout}s; spawned Word PID unavailable, no external Word processes were terminated"
        )

    if result["ok"]:
        log(f"OK {docx_path}")
        return docx_path

    raise PostprocessError(result["error"] or "Unknown Word COM post-processing error")

def main():
    parser = argparse.ArgumentParser(description="Word COM post-processing for thesis docx")
    parser.add_argument("--input", required=True, help="Input docx path")
    parser.add_argument("--timeout", type=int, default=90, help="Timeout in seconds")
    args = parser.parse_args()

    try:
        postprocess(args.input, timeout=args.timeout)
    except PostprocessTimeoutError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)
    except PostprocessError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()





