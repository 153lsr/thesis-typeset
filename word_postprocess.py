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


def postprocess(docx_path, timeout=90, config=None, mode="full"):
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

            print("[1/3] Opening document...", flush=True)
            doc = word.Documents.Open(
                docx_path,
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
            )
            print("[1/3] Done.", flush=True)

            if mode == "full":
                print("[2/3] Updating TOC and fields...", flush=True)
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
                print("[2/3] Done.", flush=True)

                print(f"[3/3] Fixing TOC fonts (L1: {toc_h1_ea} {toc_h1_size}pt, L2+: {toc_ea} {toc_size}pt)...", flush=True)
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
                print("[3/3] Done.", flush=True)
            else:
                print("[2/2] Updating fields...", flush=True)
                doc.Fields.Update()
                print("[2/2] Done.", flush=True)

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
        print(f"OK {docx_path}")
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





