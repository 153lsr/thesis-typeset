"""Formatting pipeline for the universal thesis formatter."""

import os
import shutil
import subprocess
import sys
import tempfile

from preprocess_txt_to_md import preprocess
from thesis_config import resolve_config
from thesis_format_2024 import apply_format
from word_postprocess import postprocess


def find_pandoc():
    """Locate pandoc: exe sibling dir -> _MEIPASS -> PATH."""
    candidates = []
    if getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(sys.executable), "pandoc.exe"))
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    candidates.append(os.path.join(base, "pandoc.exe"))
    for candidate in candidates:
        if os.path.isfile(candidate):
            return candidate
    found = shutil.which("pandoc")
    if found:
        return found
    return None


def convert_doc_to_docx(doc_path, out_docx):
    """Convert .doc to .docx via Word COM."""
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(os.path.abspath(doc_path))
        doc.SaveAs(os.path.abspath(out_docx), 12)
        doc.Close()
    finally:
        if word:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def run_format(input_path, output_path, skip_postprocess, log,
               config=None, config_path=None):
    """Core formatting pipeline. log(str) receives progress messages."""
    ext = os.path.splitext(input_path)[1].lower()
    supported = {".docx", ".doc", ".txt", ".md", ".tex"}
    if ext not in supported:
        log(f"不支持的格式: {ext} (支持: {' '.join(sorted(supported))})")
        return False

    if config is None:
        config, config_path = resolve_config(input_path=input_path)
    school = config.get("meta", {}).get("school_name", "")

    tmp_dir = tempfile.mkdtemp(prefix="thesisfmt_")
    tmp_docx = os.path.join(tmp_dir, "input.docx")

    try:
        if ext == ".docx":
            shutil.copy2(input_path, tmp_docx)
            log("[1/3] 输入为 docx，直接复制。")
        elif ext == ".doc":
            log("[1/3] 通过 Word COM 转换 .doc...")
            convert_doc_to_docx(input_path, tmp_docx)
            log("[1/3] 转换完成。")
        elif ext in (".txt", ".md", ".tex"):
            pandoc = find_pandoc()
            if not pandoc:
                log("错误: 未找到 pandoc。请将 pandoc.exe 放在程序同目录或加入 PATH。")
                return False
            if ext == ".txt":
                log("[1/3] 预处理 txt -> md...")
                tmp_md = os.path.join(tmp_dir, "input.md")
                preprocess(input_path, tmp_md)
                source, fmt_from = tmp_md, "markdown-smart"
            elif ext == ".md":
                source, fmt_from = input_path, "markdown-smart"
            else:
                source, fmt_from = input_path, "latex"
            log(f"[1/3] pandoc 转换中 ({fmt_from} -> docx)...")
            ret = subprocess.run(
                [pandoc, source, f"--from={fmt_from}", "--to=docx", "--standalone", "-o", tmp_docx],
                capture_output=True, text=True,
            )
            if ret.returncode != 0:
                log(f"pandoc 失败:\n{ret.stderr}")
                return False
            log("[1/3] 转换完成。")

        label = f"{school} " if school else ""
        log(f"[2/3] 应用 {label}格式规范...")
        fmt_warnings = apply_format(tmp_docx, output_path, config=config, config_path=config_path) or []
        log("[2/3] 格式化完成。")
        for warning in fmt_warnings:
            log(warning)

        runtime = config.get("_runtime", {}) if config else {}
        force_postprocess = runtime.get("caption_mode_effective") == "dynamic"
        if force_postprocess and skip_postprocess:
            log("[3/3] dynamic 题注模式要求 Word COM 更新域，已忽略跳过后处理。")

        if not skip_postprocess or force_postprocess:
            if force_postprocess:
                log("[3/3] Word COM 后处理（更新目录与动态题注域）...")
            else:
                log("[3/3] Word COM 后处理（更新目录）...")
            try:
                postprocess(output_path, config=config)
                log("[3/3] 后处理完成。")
            except Exception as exc:
                log(f"[3/3] 后处理失败（非致命）: {exc}")
                log("[3/3] 已跳过。可在 Word 中手动更新目录。")
        else:
            log("[3/3] 已跳过目录更新。")

        log(f"\n输出文件: {output_path}")
        return True
    except Exception as exc:
        log(f"\n错误: {exc}")
        return False
    finally:
        if os.path.isdir(tmp_dir):
            shutil.rmtree(tmp_dir, ignore_errors=True)

