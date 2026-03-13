"""Post-process formatted docx via Word COM (Python win32com).

Handles:
- Update TOC field
- Fix TOC entry fonts (宋体 + TNR, 小四, not bold)
"""

import argparse
import os
import sys
import threading

import pythoncom
import win32com.client as win32

wdAlertsNone = 0
wdColorBlack = 0
wdLineSpace1pt5 = 1
msoAutomationSecurityForceDisable = 3


def postprocess(docx_path, timeout=90, config=None):
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        print(f"File not found: {docx_path}", file=sys.stderr)
        sys.exit(1)

    # Extract font settings from config (or use defaults)
    if config:
        toc_cfg = config.get("toc", {})
        toc_latin = config.get("fonts", {}).get("latin", "Times New Roman")
        toc_ea = toc_cfg.get("font", config.get("fonts", {}).get("body", "宋体"))
        toc_size = toc_cfg.get("font_size", config.get("sizes", {}).get("body", 12))
    else:
        toc_latin = "Times New Roman"
        toc_ea = "宋体"
        toc_size = 12

    result = {"ok": False, "error": None}
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

            print(f"[3/3] Fixing TOC fonts to {toc_ea}+{toc_latin} {toc_size}pt...", flush=True)
            for toc in doc.TablesOfContents:
                for p in toc.Range.Paragraphs:
                    p.Range.Font.Name = toc_latin
                    p.Range.Font.NameFarEast = toc_ea
                    p.Range.Font.Size = toc_size
                    p.Range.Font.Bold = False
                    p.Range.Font.ColorIndex = wdColorBlack
                    p.Format.LineSpacingRule = wdLineSpace1pt5
                    p.Format.SpaceBefore = 0
                    p.Format.SpaceAfter = 0
            print("[3/3] Done.", flush=True)

            doc.Save()
            doc.Close()
            result["ok"] = True

        except Exception as e:
            result["error"] = str(e)
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
        print(f"TIMEOUT after {timeout}s — killing Word", file=sys.stderr)
        os.system("taskkill /F /IM WINWORD.EXE >nul 2>&1")
        sys.exit(2)

    if result["ok"]:
        print(f"OK {docx_path}")
    else:
        print(f"ERROR: {result['error']}", file=sys.stderr)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(description="Word COM post-processing for thesis docx")
    parser.add_argument("--input", required=True, help="Input docx path")
    parser.add_argument("--timeout", type=int, default=90, help="Timeout in seconds")
    args = parser.parse_args()
    postprocess(args.input, timeout=args.timeout)


if __name__ == "__main__":
    main()
