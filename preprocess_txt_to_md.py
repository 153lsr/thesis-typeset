"""Preprocess thesis plain text into well-structured Markdown for pandoc.

Handles:
- Heading detection (第X章 → H1, X.X → H2, X.X.X → H3, X.X.X.X → H4)
- Special section headings (摘要, 参考文献, 附录, 致谢)
- Abstract/Keywords separation with blank lines
- Reference entries: escape [number] and add blank lines between entries
- Ensure blank lines before/after all headings
"""

import argparse
import re


def detect_heading_level(line):
    """Return (level, line_text) or None if not a heading."""
    stripped = line.strip()
    if not stripped:
        return None

    # H1: 第X章
    if re.match(r"^第\s*\d+\s*章\b", stripped):
        return (1, stripped)

    # H1 special sections
    special_h1 = ["参考文献", "致谢"]
    normalized = stripped.replace(" ", "").replace("\u3000", "")
    for kw in special_h1:
        if normalized == kw:
            return (1, stripped)

    # H1: 附录X (but not standalone 附录)
    if re.match(r"^附录\s*[A-Z]", stripped):
        return (1, stripped)

    # H4: X.X.X.X
    if re.match(r"^\d+\.\d+\.\d+\.\d+\s", stripped):
        return (4, stripped)

    # H3: X.X.X
    if re.match(r"^\d+\.\d+\.\d+\s", stripped):
        return (3, stripped)

    # H2: X.X
    if re.match(r"^\d+\.\d+\s", stripped):
        return (2, stripped)

    return None


def is_cn_abstract_title(line):
    return line.strip().replace(" ", "").replace("\u3000", "") == "摘要"


def is_cn_keywords(line):
    return bool(re.match(r"^\s*关键词\s*[：:]", line))


def is_en_abstract_label(line):
    return bool(re.match(r"^\s*Abstract\s*[：:]?\s*$", line.strip(), re.I))


def is_en_abstract_with_content(line):
    return bool(re.match(r"^\s*Abstract\s*[：:]\s*\S", line.strip(), re.I))


def is_en_keywords(line):
    return bool(re.match(r"^\s*Key\s*words\s*[：:]", line, re.I))


def is_reference_entry(line):
    return bool(re.match(r"^\s*\[\d+\]", line))


def is_table_title(line):
    """Detect table/figure caption lines like 表3-1, 续表3-1, 表4-1."""
    return bool(re.match(r"^(续)?表\s*[\d\-]+", line.strip()))


def is_tab_separated_row(line):
    """Check if a line has tab separators (table data row)."""
    return "\t" in line and len(line.split("\t")) >= 2


def convert_table_block(title_line, data_lines, note_lines):
    """Convert a table title + tab-separated rows + notes into markdown.

    Returns a list of markdown lines.
    """
    result = []
    # Table title as a centered paragraph (will be formatted by thesis_format_2024.py)
    result.append(title_line)
    result.append("")

    if not data_lines:
        return result

    # Determine column count from header
    header_cells = data_lines[0].split("\t")
    num_cols = len(header_cells)

    # Build markdown table
    # Header row
    result.append("| " + " | ".join(c.strip() for c in header_cells) + " |")
    # Separator row
    result.append("| " + " | ".join("---" for _ in range(num_cols)) + " |")

    # Data rows
    for row_line in data_lines[1:]:
        cells = row_line.split("\t")
        # Pad or trim to match column count
        while len(cells) < num_cols:
            cells.append("")
        cells = cells[:num_cols]
        result.append("| " + " | ".join(c.strip() for c in cells) + " |")

    result.append("")

    # Table notes
    for note in note_lines:
        result.append(note)
        result.append("")

    return result


def fix_quotes(text):
    """Fix all double quotes: ASCII → Chinese, then re-pair all as left/right."""
    result = []
    need_left = True
    for ch in text:
        if ch in ('"', '\u201c', '\u201d'):
            result.append('\u201c' if need_left else '\u201d')
            need_left = not need_left
        else:
            result.append(ch)
    return ''.join(result)


def preprocess(input_path, output_path):
    with open(input_path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    # Fix double quotes: ASCII + mismatched Chinese → proper left/right pairs
    lines = [fix_quotes(line) for line in lines]

    output = []
    in_refs = False
    i = 0
    prev_was_blank = True  # start as if preceded by blank line

    while i < len(lines):
        line = lines[i].rstrip("\n\r")
        stripped = line.strip()

        # Empty line
        if not stripped:
            if not prev_was_blank:
                output.append("")
            prev_was_blank = True
            i += 1
            continue

        # Table detection: 表X-Y title followed by tab-separated rows
        if is_table_title(stripped):
            title = stripped
            data_rows = []
            notes = []
            i += 1
            # Collect tab-separated data rows
            while i < len(lines):
                row_line = lines[i].rstrip("\n\r")
                row_stripped = row_line.strip()
                if not row_stripped:
                    i += 1
                    break
                if is_tab_separated_row(row_line):
                    data_rows.append(row_stripped)
                elif re.match(r"^注[：:]", row_stripped):
                    notes.append(row_stripped)
                elif re.match(r"^相关系数", row_stripped):
                    # Special row like correlation coefficients - treat as note
                    notes.append(row_stripped)
                else:
                    break
                i += 1
            # Convert to markdown table
            if not prev_was_blank:
                output.append("")
            md_lines = convert_table_block(title, data_rows, notes)
            output.extend(md_lines)
            prev_was_blank = True
            continue

        # Chinese abstract title: "摘要" on its own line
        if is_cn_abstract_title(stripped) and not detect_heading_level(stripped):
            output.append("摘要")
            output.append("")
            prev_was_blank = True
            i += 1
            continue

        # Keywords lines
        if is_cn_keywords(stripped):
            if not prev_was_blank:
                output.append("")
            output.append(stripped)
            output.append("")
            prev_was_blank = True
            i += 1
            continue

        if is_en_keywords(stripped):
            if not prev_was_blank:
                output.append("")
            # Normalize to "Key words: ..."
            m = re.match(r"^\s*Key\s*words\s*[：:]\s*(.*)", stripped, re.I)
            if m:
                output.append("Key words: " + m.group(1))
            else:
                output.append(stripped)
            output.append("")
            prev_was_blank = True
            i += 1
            continue

        # Abstract label (standalone or with content)
        if is_en_abstract_label(stripped):
            if not prev_was_blank:
                output.append("")
            # Look ahead for abstract body on next line
            if i + 1 < len(lines) and lines[i + 1].strip():
                body = lines[i + 1].strip()
                output.append("Abstract: " + body)
                i += 2
            else:
                output.append("Abstract:")
                i += 1
            output.append("")
            prev_was_blank = True
            continue

        if is_en_abstract_with_content(stripped):
            if not prev_was_blank:
                output.append("")
            m = re.match(r"^\s*Abstract\s*[：:]\s*(.*)", stripped, re.I)
            output.append("Abstract: " + (m.group(1) if m else ""))
            output.append("")
            prev_was_blank = True
            i += 1
            continue

        # Reference entries
        if is_reference_entry(stripped):
            if not in_refs:
                in_refs = True
            # Escape leading [number] for markdown
            escaped = re.sub(r"^\[(\d+)\]", r"\\[\1\\]", stripped)
            if not prev_was_blank:
                output.append("")
            output.append(escaped)
            output.append("")
            prev_was_blank = True
            i += 1
            continue
        else:
            in_refs = False

        # Heading detection
        heading = detect_heading_level(stripped)
        if heading:
            level, text = heading
            prefix = "#" * level
            # Skip standalone "附录" (redundant if followed by 附录A/B/C)
            if text.replace(" ", "").replace("\u3000", "") == "附录":
                i += 1
                continue
            if not prev_was_blank:
                output.append("")
            output.append(f"{prefix} {text}")
            output.append("")
            prev_was_blank = True
            i += 1
            continue

        # Normal paragraph
        if not prev_was_blank:
            output.append("")
        output.append(stripped)
        prev_was_blank = False
        i += 1

    with open(output_path, "w", encoding="utf-8") as f:
        f.write("\n".join(output) + "\n")

    print(f"Preprocessed: {input_path} -> {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Preprocess thesis txt to markdown")
    parser.add_argument("--input", required=True, help="Input txt file")
    parser.add_argument("--output", required=True, help="Output md file")
    args = parser.parse_args()
    preprocess(args.input, args.output)


if __name__ == "__main__":
    main()
