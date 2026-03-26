"""Unified CLI / GUI entry point for universal thesis formatter.

Supports: .docx .doc .txt .md .tex
- With --input: CLI mode
- Without args: tkinter GUI
"""

import argparse
import os
import sys

from thesis_config import dump_default_config, resolve_config
from thesis_gui import FormatterGUI
from thesis_runner import run_format


def should_prompt_before_exit():
    """Pause only for interactive frozen executables."""
    if not getattr(sys, "frozen", False):
        return False
    stdin = getattr(sys, "stdin", None)
    if stdin is None:
        return False
    try:
        return bool(stdin.isatty())
    except Exception:
        return False


def main():
    parser = argparse.ArgumentParser(description="Universal thesis formatter")
    parser.add_argument("--input", help="Input file (.docx/.doc/.txt/.md/.tex)")
    parser.add_argument("--output", help="Output docx (default: <stem>_formatted.docx)")
    parser.add_argument("--config", help="Path to thesis_config.yaml")
    parser.add_argument("--toc-only", action="store_true",
                        help="Only insert/update TOC, keep existing document formatting")
    parser.add_argument("--dump-config", action="store_true",
                        help="Print default config YAML and exit")
    args = parser.parse_args()

    if args.dump_config:
        print(dump_default_config())
        return

    if not args.input:
        FormatterGUI()
        return

    input_path = os.path.abspath(args.input)
    if not os.path.isfile(input_path):
        print(f"Input not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    cfg, cfg_path = resolve_config(cli_config=args.config, input_path=input_path)
    if args.toc_only:
        cfg.setdefault("toc", {})
        cfg["toc"]["enabled"] = True
        cfg["toc"]["only_insert"] = True

    stem = os.path.splitext(os.path.basename(input_path))[0]
    input_dir = os.path.dirname(input_path)
    output_path = (os.path.abspath(args.output) if args.output
                   else os.path.join(input_dir, f"{stem}_formatted.docx"))

    ok = run_format(input_path, output_path, print,
                    config=cfg, config_path=cfg_path)
    if not ok:
        sys.exit(1)

    if should_prompt_before_exit():
        try:
            input("\n按回车键关闭...")
        except EOFError:
            pass


if __name__ == "__main__":
    main()
