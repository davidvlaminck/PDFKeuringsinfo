"""Convert the Vincotte foutcodes PDF into an Excel workbook.

This script uses ``pdftotext -bbox-layout`` to extract the table layout and
produces two sheets:

- ``foutcodes``: best-effort parsed table with code, type, section and flags
- ``secties``: detected section headers for auditability

Usage:
    python Vincotte/convert_vincotte_pdf_to_excel.py \
        --pdf "Vincotte/Volledige lijst foutcodes Vincotte.pdf" \
        --output "Vincotte/Volledige lijst foutcodes Vincotte.xlsx"
"""

from __future__ import annotations

import argparse
import html
import re
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd


SHORT_X_THRESHOLD = 260.0
FLAG_X_THRESHOLD = 500.0
HH_X_THRESHOLD = 521.0  # observed flag clusters around ~510 and ~532

HEADER_MARKERS = (
    "TYPEZINNEN ELECTRICITY",
    "Questionnaires Electricity",
    "huishoudelijk en niet-huishoudelijk",
    "SHORT PF (NL)",
    "LONG PF (NL)",
    "Nr.",
    "(I)",
    "(O)",
    "(N)",
)

FOOTER_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\b.*\bVinotte\b.*\b\d+\s*/\s*\d+\b")
CODE_RE = re.compile(r"^(\d{4}):")
CODE_WORD_RE = re.compile(r"^\d{4}$")
TYPE_WORD_RE = re.compile(r"^[ION]$")


@dataclass(frozen=True)
class Word:
    x: float
    text: str


@dataclass(frozen=True)
class VisualLine:
    page: int
    y: float
    words: tuple[Word, ...]


@dataclass
class ParsedRow:
    page: int
    row_order: int
    section: str
    code: str
    type: str
    short_pf_nl: str
    long_pf_nl: str
    hh: str
    nhh: str
    raw_text: str


@dataclass
class SectionEntry:
    page: int
    row_order: int
    section: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--pdf",
        default=str(Path(__file__).with_name("Volledige lijst foutcodes Vincotte.pdf")),
        help="Path to the source PDF.",
    )
    parser.add_argument(
        "--output",
        default=str(Path(__file__).with_name("Volledige lijst foutcodes Vincotte.xlsx")),
        help="Path to the output Excel workbook.",
    )
    return parser.parse_args()


def run_pdftotext_bbox_layout(pdf_path: Path) -> str:
    with tempfile.TemporaryDirectory() as tmpdir:
        html_path = Path(tmpdir) / "out.html"
        subprocess.run(
            ["pdftotext", "-bbox-layout", str(pdf_path), str(html_path)],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return html_path.read_text(encoding="utf-8", errors="replace")


def extract_pages(html_text: str) -> list[str]:
    return re.findall(r'<page width="[0-9.]+" height="[0-9.]+">(.*?)</page>', html_text, re.S)


def decode_word_text(text: str) -> str:
    return html.unescape(text)


def extract_visual_lines(page_html: str, page_number: int) -> list[VisualLine]:
    lines: list[tuple[float, float, list[Word]]] = []
    for m in re.finditer(
        r'<line xMin="([0-9.]+)" yMin="([0-9.]+)" xMax="([0-9.]+)" yMax="([0-9.]+)">(.*?)</line>',
        page_html,
        re.S,
    ):
        y = float(m.group(2))
        x = float(m.group(1))
        body = m.group(5)
        words: list[Word] = []
        for wm in re.finditer(
            r'<word xMin="([0-9.]+)" yMin="([0-9.]+)" xMax="([0-9.]+)" yMax="([0-9.]+)">(.*?)</word>',
            body,
            re.S,
        ):
            words.append(Word(float(wm.group(1)), decode_word_text(wm.group(5))))
        if words:
            lines.append((y, x, words))

    # Keep physical order within a page, but merge fragments that share the same baseline.
    lines.sort(key=lambda item: (item[0], item[1]))
    merged: list[VisualLine] = []
    current_words: list[Word] = []
    current_y: float | None = None
    for y, _, words in lines:
        if current_y is None or abs(y - current_y) < 0.05:
            current_words.extend(words)
            current_y = y
        else:
            merged.append(VisualLine(page_number, current_y, tuple(sorted(current_words, key=lambda w: w.x))))
            current_words = list(words)
            current_y = y
    if current_words:
        merged.append(VisualLine(page_number, current_y if current_y is not None else 0.0, tuple(sorted(current_words, key=lambda w: w.x))))
    return merged


def is_header_or_footer(text: str) -> bool:
    if not text.strip():
        return True
    if any(marker in text for marker in HEADER_MARKERS):
        return True
    if "Vinotte" in text and re.search(r"\b\d{2}/\d{2}/\d{4}\b", text):
        return True
    if FOOTER_RE.search(text):
        return True
    return False


def is_section_line(words: Iterable[Word]) -> str | None:
    nonx_words = [w for w in words if w.text != "x" and not CODE_WORD_RE.fullmatch(w.text) and not TYPE_WORD_RE.fullmatch(w.text)]
    if not nonx_words:
        return None
    short_text = " ".join(w.text for w in nonx_words if w.x < SHORT_X_THRESHOLD).strip()
    long_text = " ".join(w.text for w in nonx_words if SHORT_X_THRESHOLD <= w.x < FLAG_X_THRESHOLD).strip()
    x_count = sum(1 for w in words if w.text == "x")
    if x_count >= 1 and short_text and short_text == long_text:
        return short_text
    return None


def extract_code(words: Iterable[Word]) -> str | None:
    for w in words:
        if w.x < 100.0:
            if CODE_WORD_RE.fullmatch(w.text):
                return w.text
            m = CODE_RE.match(w.text)
            if m:
                return m.group(1)
    for w in words:
        m = CODE_RE.match(w.text)
        if m:
            return m.group(1)
    return None


def extract_type(words: Iterable[Word]) -> str | None:
    for w in words:
        if w.x < 140.0 and TYPE_WORD_RE.fullmatch(w.text):
            return w.text
    return None


def clean_join(parts: list[str]) -> str:
    text = " ".join(part.strip() for part in parts if part and part.strip())
    text = re.sub(r"\s+", " ", text).strip()
    return text


def clip_after_next_code_prefix(text: str, current_code: str) -> str:
    # When a row span contains the next row's code prefix, keep only the first row's text.
    prefix_matches = list(re.finditer(r"\b(\d{4}):", text))
    if not prefix_matches:
        return text
    for match in prefix_matches:
        if match.group(1) != current_code:
            return text[: match.start()].strip()
    return text


def parse_row_span(lines: list[VisualLine], page_number: int, row_order: int, current_section: str) -> tuple[ParsedRow | None, SectionEntry | None, str]:
    raw_words: list[str] = []
    short_words: list[str] = []
    long_words: list[str] = []
    x_flags: list[float] = []
    code: str | None = None
    row_type: str | None = None

    for line in lines:
        line_text = clean_join([w.text for w in line.words])
        if is_header_or_footer(line_text):
            continue

        section_text = is_section_line(line.words)
        if section_text:
            current_section = section_text
            continue

        raw_words.extend([w.text for w in line.words if w.text != "x"])
        if code is None:
            code = extract_code(line.words)
        if row_type is None:
            row_type = extract_type(line.words)

        for w in line.words:
            if w.text == "x":
                x_flags.append(w.x)
            elif CODE_WORD_RE.fullmatch(w.text) and w.x < 100.0:
                continue
            elif TYPE_WORD_RE.fullmatch(w.text) and w.x < 140.0:
                continue
            elif w.x < SHORT_X_THRESHOLD:
                short_words.append(w.text)
            elif w.x < FLAG_X_THRESHOLD:
                long_words.append(w.text)

    if code is None:
        return None, None, current_section

    short_text = clip_after_next_code_prefix(clean_join(short_words), code)
    long_text = clip_after_next_code_prefix(clean_join(long_words), code)
    raw_text = clip_after_next_code_prefix(clean_join(raw_words), code)

    hh = "x" if any(x < HH_X_THRESHOLD for x in x_flags) else ""
    nhh = "x" if any(x >= HH_X_THRESHOLD for x in x_flags) else ""

    return (
        ParsedRow(
            page=page_number,
            row_order=row_order,
            section=current_section,
            code=code,
            type=row_type or "",
            short_pf_nl=short_text,
            long_pf_nl=long_text,
            hh=hh,
            nhh=nhh,
            raw_text=raw_text,
        ),
        None,
        current_section,
    )


def parse_pdf(pdf_path: Path) -> tuple[list[ParsedRow], list[SectionEntry]]:
    html_text = run_pdftotext_bbox_layout(pdf_path)
    page_bodies = extract_pages(html_text)

    rows: list[ParsedRow] = []
    sections: list[SectionEntry] = []
    current_section = ""
    row_order = 0

    for page_number, page_html in enumerate(page_bodies, start=1):
        lines = extract_visual_lines(page_html, page_number)

        # Row starts are identified by the short-description prefix like "1008:".
        start_indices = [
            idx
            for idx, line in enumerate(lines)
            if any(CODE_RE.match(word.text) for word in line.words) or any(CODE_RE.match(word.text) for word in line.words if word.x < SHORT_X_THRESHOLD)
        ]

        # Page-level sections before the first row should still be captured.
        prev_start = 0
        if not start_indices:
            continue

        for i, start_idx in enumerate(start_indices):
            end_idx = start_indices[i + 1] if i + 1 < len(start_indices) else len(lines)
            # First row on a page can have content that visually precedes the start line.
            span_start = 0 if i == 0 else start_idx
            span_lines = lines[span_start:end_idx]

            # Check for section lines and update the active section.
            for line in span_lines:
                section_text = is_section_line(line.words)
                if section_text:
                    current_section = section_text
                    sections.append(SectionEntry(page=page_number, row_order=row_order, section=section_text))

            parsed_row, _, current_section = parse_row_span(span_lines, page_number, row_order, current_section)
            if parsed_row:
                rows.append(parsed_row)
                row_order += 1

    return rows, sections


def write_workbook(rows: list[ParsedRow], sections: list[SectionEntry], output_path: Path) -> None:
    df = pd.DataFrame([row.__dict__ for row in rows])
    section_df = pd.DataFrame([section.__dict__ for section in sections])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="foutcodes")
        section_df.to_excel(writer, index=False, sheet_name="secties")

        # Light formatting: freeze the header row and add filters.
        wb = writer.book
        for ws_name in ["foutcodes", "secties"]:
            ws = wb[ws_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value is not None:
                        max_len = max(max_len, len(str(cell.value)))
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 80)


def main() -> None:
    args = parse_args()
    pdf_path = Path(args.pdf).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    rows, sections = parse_pdf(pdf_path)
    write_workbook(rows, sections, output_path)

    print(f"Wrote {len(rows)} rows and {len(sections)} section markers to {output_path}")


if __name__ == "__main__":
    main()


