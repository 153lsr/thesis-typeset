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

wdAlertsNone = 0
wdColorBlack = 0
wdLineSpace1pt5 = 1
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


def postprocess(docx_path, timeout=90, config=None):
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        raise PostprocessError(f"File not found: {docx_path}")

    if config:
        toc_cfg = config.get("toc", {})
        fonts_cfg = config.get("fonts", {})
        sizes_cfg = config.get("sizes", {})
        toc_latin = fonts_cfg.get("latin", "Times New Roman")
        toc_ea = toc_cfg.get("font", fonts_cfg.get("body", "宋体"))
        toc_size = toc_cfg.get("font_size", sizes_cfg.get("body", 12))
        toc_h1_ea = toc_cfg.get("h1_font", fonts_cfg.get("h1", toc_ea))
        toc_h1_size = toc_cfg.get("h1_font_size", sizes_cfg.get("h1", toc_size))
    else:
        toc_latin = "Times New Roman"
        toc_ea = "宋体"
        toc_size = 12
        toc_h1_ea = toc_ea
        toc_h1_size = toc_size

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

            print("[2/3] Updating TOC and fields...", flush=True)
            for toc in doc.TablesOfContents:
                toc.Update()
            doc.Fields.Update()
            print("[2/3] Done.", flush=True)

            print(f"[3/3] Fixing TOC fonts (L1: {toc_h1_ea} {toc_h1_size}pt, L2+: {toc_ea} {toc_size}pt)...", flush=True)
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
                    p.Range.Font.Name = toc_latin
                    p.Range.Font.NameFarEast = toc_h1_ea if is_level1 else toc_ea
                    p.Range.Font.Size = toc_h1_size if is_level1 else toc_size
                    p.Range.Font.Bold = False
                    p.Range.Font.ColorIndex = wdColorBlack
                    p.Format.LineSpacingRule = wdLineSpace1pt5
                    p.Format.SpaceBefore = 0
                    p.Format.SpaceAfter = 0
            print("[3/3] Done.", flush=True)

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
