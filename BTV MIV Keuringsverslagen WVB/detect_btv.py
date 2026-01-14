"""Simpel script om PDF's te scannen en te raden of ze van BTV zijn.

Gebruik:
  python3 "BTV MIV Keuringsverslagen WVB/detect_btv.py" "BTV MIV Keuringsverslagen WVB/pdfs/*.pdf"
  python3 "BTV MIV Keuringsverslagen WVB/detect_btv.py" "BTV MIV Keuringsverslagen WVB/pdfs" -r

Excel aanvullen (Lijst miv.xlsx):
  python3 "BTV MIV Keuringsverslagen WVB/detect_btv.py" \
    --update-excel "BTV MIV Keuringsverslagen WVB/Lijst miv.xlsx" \
    --pdf-dir "BTV MIV Keuringsverslagen WVB/pdfs" \
    --excel-out "BTV MIV Keuringsverslagen WVB/Lijst miv.filled.xlsx"

Geen installatie/egg nodig: dit draait rechtstreeks.

Dependency:
- pypdf
- openpyxl (enkel nodig voor --update-excel)

Installeer in een venv:
  python3 -m venv .venv
  . .venv/bin/activate
  python -m pip install -U pip pypdf openpyxl
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import re
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Iterable

from datetime import datetime

from pypdf import PdfReader


# -------------------- types --------------------


@dataclass(frozen=True)
class DetectionResult:
    path: str
    is_btv: bool
    confidence: float
    matched_rules: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    metadata: dict[str, Any] = field(default_factory=dict)
    text_stats: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class PdfText:
    pages_scanned: int
    extracted_chars: int
    text: str
    metadata: dict[str, str]


# -------------------- rules --------------------


_TOKEN_RULES: list[tuple[str, float, re.Pattern[str]]] = [
    ("token:btv", 3.0, re.compile(r"\bbtv\b", re.IGNORECASE)),
    ("token:keuringsverslag", 2.0, re.compile(r"\bkeuringsverslag\b", re.IGNORECASE)),
    ("token:keuringsattest", 2.0, re.compile(r"\bkeuringsattest\b", re.IGNORECASE)),
    ("token:keuring", 0.5, re.compile(r"\bkeuring\b", re.IGNORECASE)),
    ("token:miv", 1.0, re.compile(r"\bmiv\b", re.IGNORECASE)),
    ("token:inbreuken", 0.5, re.compile(r"\binbreuk(en)?\b", re.IGNORECASE)),
    ("token:opmerkingen", 0.5, re.compile(r"\bopmerking(en)?\b", re.IGNORECASE)),
]

_META_RULES: list[tuple[str, float, re.Pattern[str]]] = [
    ("meta:btv", 1.0, re.compile(r"\bbtv\b", re.IGNORECASE)),
]


def score_text(text: str) -> tuple[float, list[str]]:
    matches: list[str] = []
    score = 0.0
    for name, points, pat in _TOKEN_RULES:
        if pat.search(text):
            score += points
            matches.append(name)
    return score, matches


def score_metadata(meta: dict[str, str]) -> tuple[float, list[str]]:
    matches: list[str] = []
    score = 0.0
    for k, v in meta.items():
        if not v:
            continue
        for name, points, pat in _META_RULES:
            if pat.search(str(v)):
                score += points
                matches.append(f"{name}:{k}")
    return score, matches


def confidence_from_score(raw_score: float) -> float:
    if raw_score <= 0:
        return 0.0
    return min(raw_score / 5.0, 1.0)


# -------------------- pdf reading --------------------


def read_pdf_text(path: str, max_pages: int = 3) -> PdfText:
    reader = PdfReader(path)

    meta: dict[str, str] = {}
    try:
        if reader.metadata:
            for k, v in reader.metadata.items():
                kk = str(k).lstrip("/")
                meta[kk] = "" if v is None else str(v)
    except Exception:
        meta = {}

    n_pages = len(reader.pages)
    pages_to_scan = min(max_pages, n_pages)

    chunks: list[str] = []
    extracted = 0
    for i in range(pages_to_scan):
        page = reader.pages[i]
        txt = page.extract_text() or ""
        chunks.append(txt)
        extracted += len(txt)

    joined = "\n".join(chunks)
    return PdfText(
        pages_scanned=pages_to_scan,
        extracted_chars=extracted,
        text=joined,
        metadata=meta,
    )


# -------------------- detector --------------------


def detect_btv(path: str, *, threshold: float = 0.7, max_pages: int = 3) -> DetectionResult:
    errors: list[str] = []
    matched: list[str] = []
    meta: dict[str, str] = {}
    text_stats: dict[str, object] = {"pages_scanned": 0, "extracted_chars": 0}

    try:
        pdf = read_pdf_text(path, max_pages=max_pages)
        meta = pdf.metadata
        text_stats = {"pages_scanned": pdf.pages_scanned, "extracted_chars": pdf.extracted_chars}

        s_text, m_text = score_text(pdf.text)
        s_meta, m_meta = score_metadata(meta)
        matched.extend(m_text)
        matched.extend(m_meta)

        raw = s_text + s_meta
        conf = confidence_from_score(raw)
        is_btv = conf >= threshold

        return DetectionResult(
            path=path,
            is_btv=is_btv,
            confidence=conf,
            matched_rules=matched,
            errors=errors,
            metadata=meta,
            text_stats=text_stats,
        )
    except Exception as e:
        errors.append(f"{type(e).__name__}: {e}")
        return DetectionResult(
            path=path,
            is_btv=False,
            confidence=0.0,
            matched_rules=matched,
            errors=errors,
            metadata=meta,
            text_stats=text_stats,
        )


# -------------------- excel update --------------------


_MIV_FROM_FILENAME_PATTERNS: list[re.Pattern[str]] = [
    # e.g. "Keuringsattest MIV635.Teller.pdf" of "Keuring MIV008.2.pdf"
    re.compile(r"\bMIV\s*(\d{3,4})(?:\.(\d+))?\b", re.IGNORECASE),
    # e.g. "MIV418 SAT1" (geen punt in bestandsnaam, maar in Excel bestaat het)
    re.compile(r"\bMIV\s*(\d{3,4})\s*SAT\s*(\d+)", re.IGNORECASE),
    # e.g. "MIV4361" wat MIV436.1 betekent (gezien in dataset)
    re.compile(r"\bMIV\s*(\d{3})(\d)\b", re.IGNORECASE),
]


def normalize_miv_code(value: object) -> str | None:
    """Normalize Excel cell values (MIVxxx kolom) into a mapping key.

    Examples:
    - "MIV418" -> "MIV418"
    - "MIV418 SAT1" -> "MIV418.1"
    - "MIV041 SAT" -> "MIV041.1" (implicit SAT = .1)
    - "MIV418.2" -> "MIV418.2"
    """
    if value is None:
        return None

    # Strip and normalize whitespace, including NBSP.
    s = str(value)
    s = s.replace("\u00A0", " ")  # non-breaking space
    s = s.strip()
    if not s:
        return None
    s = re.sub(r"\s+", " ", s)

    # normalize: MIV418 SAT1 -> MIV418.1
    m = re.search(r"\bMIV\s*(\d{3,4})\s*SAT\s*(\d+)\b", s, re.IGNORECASE)
    if m:
        return f"MIV{m.group(1)}.{int(m.group(2))}"

    # normalize: MIV041 SAT -> MIV041.1 (implicit)
    m = re.search(r"\bMIV\s*(\d{3,4})\s*SAT\b", s, re.IGNORECASE)
    if m:
        return f"MIV{m.group(1)}.1"

    # normalize: MIV418.1 (keep)
    m = re.search(r"\bMIV\s*(\d{3,4})(?:\.(\d+))?\b", s, re.IGNORECASE)
    if m:
        base = f"MIV{m.group(1)}"
        if m.group(2):
            return f"{base}.{int(m.group(2))}"
        return base
    return None


def _pdf_is_sat_variant(filename: str) -> bool:
    """Heuristic: variants that should not win over the base PDF when choosing between multiple files."""
    s = filename.lower()
    return any(tok in s for tok in ("teller", "tellerkast"))


def miv_key_from_pdf_filename(filename: str) -> str | None:
    base = Path(filename).stem

    # Prefer explicit patterns.
    for pat in _MIV_FROM_FILENAME_PATTERNS[:2]:
        m = pat.search(base)
        if not m:
            continue
        num = m.group(1)
        sat = m.group(2)
        k = f"MIV{num}"
        if sat:
            return f"{k}.{int(sat)}"
        return k

    # Fallback: compact form MIV4361 -> MIV436.1 (only for 3 digits + 1 digit)
    m = _MIV_FROM_FILENAME_PATTERNS[2].search(base)
    if m:
        return f"MIV{int(m.group(1))}.{int(m.group(2))}"

    return None


def choose_best_excel_row(rows: list[int], ws: Any, miv_cell_col: int, pdf_name: str) -> int:
    """Pick the best row among duplicates for the given PDF.

    - For .1 PDFs, prefer an Excel row whose MIV field contains 'SAT' or '.1'
    - For base PDFs, prefer row without 'SAT' and without a dot
    """
    key = miv_key_from_pdf_filename(pdf_name)
    wants_dot1 = bool(key and key.endswith(".1"))

    def rank(row: int) -> int:
        v = ws.cell(row, miv_cell_col).value
        s = str(v).upper() if v is not None else ""
        is_sat = "SAT" in s
        has_dot = "." in s
        is_dot1 = ".1" in s
        if wants_dot1:
            return 0 if (is_sat or is_dot1) else 2
        return 0 if ((not is_sat) and (not has_dot)) else 2

    return sorted(rows, key=rank)[0]


def is_pdf_text_readable(path: str, *, max_pages: int = 1, min_chars: int = 20) -> tuple[bool, int, list[str]]:
    """Heuristic: considered 'readable' if we can extract enough text chars from first page(s)."""
    errors: list[str] = []
    try:
        pdf = read_pdf_text(path, max_pages=max_pages)
        return pdf.extracted_chars >= min_chars, pdf.extracted_chars, errors
    except Exception as e:
        errors.append(f"{type(e).__name__}: {e}")
        return False, 0, errors


_DATE_DMY_RE = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")


def _parse_dmy_date(s: str) -> datetime | None:
    m = _DATE_DMY_RE.search(s)
    if not m:
        return None
    try:
        return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
    except ValueError:
        return None


def extract_keuringsdatum(text: str) -> datetime | None:
    """Extract 'Datum(s) van controle: dd/mm/yyyy' from PDF text."""
    # normalize whitespace to make regex simpler
    t = text.replace("\u00A0", " ")
    t = re.sub(r"[\t\r]+", " ", t)
    # Try strict label
    m = re.search(
        r"Datum\(s\)\s+van\s+controle\s*:\s*(\d{1,2}/\d{1,2}/\d{4})",
        t,
        re.IGNORECASE,
    )
    if m:
        return _parse_dmy_date(m.group(1))
    # Fallback: any date near 'Datum(s) van controle'
    m = re.search(r"Datum\(s\)\s+van\s+controle\b(.{0,80})", t, re.IGNORECASE | re.DOTALL)
    if m:
        return _parse_dmy_date(m.group(1))
    return None


def extract_keuringsresultaat(text: str) -> str | None:
    """Extract result from PDF text.

    Returns one of: 'conform', 'conform met opmerkingen', 'niet conform'
    """
    t = text.replace("\u00A0", " ")
    t_low = t.lower()

    # Strong signals (often on page 1 header)
    if re.search(r"\binstallatie\s+is\s+niet\s+conform\b", t_low):
        return "niet conform"
    if re.search(r"\binstallatie\s+is\s+conform\b", t_low):
        return "conform"
    if re.search(r"\binstallatie\s+is\s+conform\s+met\s+opmerkingen\b", t_low):
        return "conform met opmerkingen"

    # Prefer scanning a window after the label.
    m = re.search(
        r"Besluit\s+van\s+de\s+controle\s*:\s*(.{0,2000})",
        t,
        re.IGNORECASE | re.DOTALL,
    )
    window = m.group(1).lower() if m else t_low

    # Order matters: 'conform met opmerkingen' contains 'conform'
    if "conform met opmerkingen" in window:
        return "conform met opmerkingen"

    if "niet conform" in window:
        return "niet conform"

    # Many documents phrase 'voldoet aan' for conform.
    if re.search(r"\bvoldoet\b", window) or re.search(r"\bconform\b", window):
        return "conform"

    return None


def extract_inbreuken(text: str) -> str | None:
    """Extract the 'vastgestelde inbreuken' paragraph/section.

    Returns a cleaned, single-line string (can contain multiple inbreuken separated by ' | ').
    """
    t = text.replace("\u00A0", " ")
    # capture a window after the header; stop at a next common section header
    m = re.search(
        r"vastgestelde\s+inbreuk(?:en)?\s*:?\s*(.{0,3000})",
        t,
        re.IGNORECASE | re.DOTALL,
    )
    if not m:
        return None
    chunk = m.group(1)

    # stop at next likely header
    stop = re.search(
        r"\n\s*(besluit\s+van\s+de\s+controle|opmerkingen|controle\s*gegevens|datum\(s\)\s+van\s+controle)\b",
        chunk,
        re.IGNORECASE,
    )
    if stop:
        chunk = chunk[: stop.start()]

    # normalize whitespace
    lines = [ln.strip() for ln in chunk.splitlines() if ln.strip()]
    if not lines:
        return None

    # Heuristic: keep only meaningful lines, join
    # If lines look like bullets/numbered items, keep all.
    cleaned: list[str] = []
    for ln in lines:
        # strip bullet-like prefixes
        ln2 = re.sub(r"^[-•\u2022\s]+", "", ln)
        ln2 = re.sub(r"^\(?\d+\)?[\.)\s]+", "", ln2)
        cleaned.append(ln2.strip())

    cleaned = [c for c in cleaned if c]
    if not cleaned:
        return None
    return " | ".join(cleaned)


def update_excel(
    excel_path: str,
    pdf_dir: str,
    *,
    excel_out: str | None,
    sheet_name: str | None,
    miv_column: str = "MIVxxx",
    pdf_col_name: str = "PDF-bestand",
    readable_col_name: str = "PDF inleesbaar",
    keuringsdatum_col_name: str = "Keuringsdatum",
    keuringsresultaat_col_name: str = "Keuringsresultaat",
    inbreuken_col_name: str = "Vastgestelde inbreuken",
    max_pages: int = 1,
    min_chars: int = 20,
) -> int:
    try:
        import openpyxl  # type: ignore
    except Exception as e:
        print(
            "openpyxl ontbreekt. Installeer met: python -m pip install openpyxl\n" f"Details: {e}",
            file=sys.stderr,
        )
        return 2

    xlsx = Path(excel_path)
    if not xlsx.exists():
        print(f"Excel niet gevonden: {excel_path}", file=sys.stderr)
        return 2

    pdf_root = Path(pdf_dir)
    if not pdf_root.exists() or not pdf_root.is_dir():
        print(f"PDF map niet gevonden: {pdf_dir}", file=sys.stderr)
        return 2

    wb = openpyxl.load_workbook(xlsx)
    ws = wb[sheet_name] if sheet_name else wb.active

    # find header row that contains miv_column
    header_row = None
    header_to_col: dict[str, int] = {}
    for r in range(1, min(50, ws.max_row) + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if not any(v is not None and str(v).replace("\u00A0", " ").strip() != "" for v in row_vals):
            continue
        header_to_col = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            key = str(v).replace("\u00A0", " ").strip()
            if not key:
                continue
            header_to_col[key] = c

        if miv_column in header_to_col:
            header_row = r
            break

    if not header_row:
        print(f"Kon header met kolom '{miv_column}' niet vinden.", file=sys.stderr)
        return 2

    def ensure_column(col_name: str) -> int:
        nonlocal header_to_col
        if col_name in header_to_col:
            return header_to_col[col_name]
        new_col = ws.max_column + 1
        ws.cell(header_row, new_col).value = col_name
        header_to_col[col_name] = new_col
        return new_col

    pdf_col = ensure_column(pdf_col_name)
    readable_col = ensure_column(readable_col_name)
    keuringsdatum_col = ensure_column(keuringsdatum_col_name)
    keuringsresultaat_col = ensure_column(keuringsresultaat_col_name)
    inbreuken_col = ensure_column(inbreuken_col_name)

    # Optional: used to silence "missing" rows in report.
    keuringsattest_col = header_to_col.get("Keuringsattest beschikbaar?")

    # index excel rows by MIV key
    miv_key_to_rows: dict[str, list[int]] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        key = normalize_miv_code(ws.cell(r, header_to_col[miv_column]).value)
        if not key:
            continue
        miv_key_to_rows.setdefault(key, []).append(r)

    # scan PDFs
    pdf_files = sorted(pdf_root.rglob("*.pdf"))

    unmapped_pdfs: list[str] = []
    ambiguous_pdfs: list[str] = []

    for pdf_path in pdf_files:
        key = miv_key_from_pdf_filename(pdf_path.name)
        if not key:
            unmapped_pdfs.append(str(pdf_path))
            continue

        rows = miv_key_to_rows.get(key)
        if not rows:
            unmapped_pdfs.append(str(pdf_path))
            continue

        row = choose_best_excel_row(rows, ws, header_to_col[miv_column], pdf_path.name)
        if len(rows) > 1:
            ambiguous_pdfs.append(f"{pdf_path} -> {key} (rows={rows}, chosen={row})")

        readable, extracted_chars, errs = is_pdf_text_readable(
            str(pdf_path), max_pages=max_pages, min_chars=min_chars
        )

        ws.cell(row, pdf_col).value = pdf_path.name
        ws.cell(row, readable_col).value = "JA" if readable else "NEE"
        if errs:
            ws.cell(row, pdf_col).value = f"{pdf_path.name} (ERROR: {errs[0]})"
        elif not readable:
            ws.cell(row, pdf_col).value = f"{pdf_path.name} (chars={extracted_chars})"

        # Only parse date/result when readable and no extraction error.
        if readable and not errs:
            try:
                pdf_text = read_pdf_text(str(pdf_path), max_pages=max(2, max_pages)).text
                dt = extract_keuringsdatum(pdf_text)
                res = extract_keuringsresultaat(pdf_text)
                inbr = extract_inbreuken(pdf_text)
                if dt:
                    ws.cell(row, keuringsdatum_col).value = dt.date()
                if res:
                    ws.cell(row, keuringsresultaat_col).value = res
                if inbr:
                    ws.cell(row, inbreuken_col).value = inbr
            except Exception:
                # Keep Excel update robust; if parsing fails, leave cells empty.
                pass

    # list excel rows that did not get a pdf mapped
    unmapped_excel: list[str] = []
    for key, rows in sorted(miv_key_to_rows.items()):
        # mapped?
        if any(ws.cell(r, pdf_col).value is not None and str(ws.cell(r, pdf_col).value).strip() != "" for r in rows):
            continue

        # ignore if all rows for this key have Keuringsattest beschikbaar? == NEE
        if keuringsattest_col is not None:
            all_no = True
            for r in rows:
                v = ws.cell(r, keuringsattest_col).value
                s = str(v).strip().upper() if v is not None else ""
                if s != "NEE":
                    all_no = False
                    break
            if all_no:
                continue

        unmapped_excel.append(f"{key} (rows={rows})")

    out_path = Path(excel_out) if excel_out else xlsx
    wb.save(out_path)

    print(f"Excel bijgewerkt: {out_path}")
    print(f"Toegevoegd/gevuld kolommen: '{pdf_col_name}', '{readable_col_name}'")
    print(
        f"Extra kolommen: '{keuringsdatum_col_name}', '{keuringsresultaat_col_name}', '{inbreuken_col_name}'"
    )

    if ambiguous_pdfs:
        print("\nAmbigue mapping (meerdere Excel rijen voor dezelfde MIV key):")
        for line in ambiguous_pdfs:
            print(f"- {line}")

    if unmapped_pdfs:
        print("\nPDFs niet kunnen mappen:")
        for p in unmapped_pdfs:
            print(f"- {os.path.relpath(p)}")

    if unmapped_excel:
        print("\nExcel-keuringen zonder gemapte PDF:")
        for line in unmapped_excel:
            print(f"- {line}")

    return 0


# -------------------- CLI helpers --------------------


def iter_paths(inputs: list[str], recursive: bool) -> Iterable[str]:
    for raw in inputs:
        expanded = glob.glob(raw)
        candidates = expanded if expanded else [raw]
        for c in candidates:
            p = Path(c)
            if p.is_dir():
                if recursive:
                    yield from (str(f) for f in p.rglob("*.pdf"))
                else:
                    yield from (str(f) for f in p.glob("*.pdf"))
            else:
                yield str(p)


def main() -> int:
    ap = argparse.ArgumentParser()

    # mode selection
    ap.add_argument("--update-excel", help="Pad naar Lijst miv.xlsx om aan te vullen")
    ap.add_argument("--pdf-dir", help="Map met PDF's (wordt recursief gescand)")
    ap.add_argument("--excel-out", help="Output pad voor bijgewerkte Excel (default: overschrijven)")
    ap.add_argument("--sheet", help="Sheetnaam (default: actieve sheet)")
    ap.add_argument("--min-chars", type=int, default=20, help="Min. aantal tekstchars om 'inleesbaar' te zijn")

    # original detector args
    ap.add_argument("paths", nargs="*", help="PDF file(s), directory, or glob(s)")
    ap.add_argument("--threshold", type=float, default=0.7)
    ap.add_argument("--max-pages", type=int, default=3)
    ap.add_argument("-r", "--recursive", action="store_true")
    ap.add_argument("--json", action="store_true", help="Output as JSON")
    args = ap.parse_args()

    if args.update_excel:
        if not args.pdf_dir:
            print("--pdf-dir is verplicht samen met --update-excel", file=sys.stderr)
            return 2
        return update_excel(
            args.update_excel,
            args.pdf_dir,
            excel_out=args.excel_out,
            sheet_name=args.sheet,
            max_pages=min(args.max_pages, 3),
            min_chars=args.min_chars,
        )

    if not args.paths:
        ap.print_help(sys.stderr)
        return 2

    results = [
        detect_btv(p, threshold=args.threshold, max_pages=args.max_pages)
        for p in iter_paths(args.paths, args.recursive)
    ]

    if args.json:
        print(json.dumps([asdict(r) for r in results], ensure_ascii=False, indent=2))
        return 0

    for r in results:
        yn = "yes" if r.is_btv else "no"
        rel = os.path.relpath(r.path)
        try:
            print(f"{yn}\t{r.confidence:.2f}\t{rel}")
        except BrokenPipeError:
            try:
                sys.stdout.close()
            except Exception:
                pass
            return 0

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except BrokenPipeError:
        raise SystemExit(0)
