"""Extractie van keuringsdata uit PDF's in VincottePlatform/ExtractPDF/pdfs/

Output: CSV met per PDF:
- pdf_filename
- pdf_readable (bool)
- installation (gevonden installatienummer of leeg)
- mapping_diagnostics
- keuringsdatum_raw
- keuringsdatum_iso (YYYY-MM-DD)
- resultaat_raw
- resultaat_normalized
- inbreuken

Gebruik:
    python extract_pdf_reports.py --pdf-dir ./pdfs --out extraction_results.csv

Afhankelijkheden (optioneel): pdfplumber, PyPDF2, pytesseract (voor OCR).
Het script probeert eerst pdfplumber, dan PyPDF2; als geen van beide beschikbaar is, geeft het instructie.
"""
from __future__ import annotations

import argparse
import csv
import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

logger = logging.getLogger(__name__)

# Try import pdf extraction libraries (best-effort, not required at import time)
try:
    import pdfplumber  # type: ignore
    _HAS_PDFPLUMBER = True
except Exception:
    _HAS_PDFPLUMBER = False

try:
    from PyPDF2 import PdfReader  # type: ignore
    _HAS_PYPDF2 = True
except Exception:
    _HAS_PYPDF2 = False

# Optional pdfminer.six (best-effort)
try:
    from pdfminer.high_level import extract_text as _pdfminer_extract_text  # type: ignore
    _HAS_PDFMINER = True
except Exception:
    _HAS_PDFMINER = False


def extract_text_from_pdf(path: Path, max_pages: int = 5) -> Tuple[str, bool]:
    """Return (text, readable_flag).

    readable_flag is True when extracted text is non-empty and likely not a pure scan.
    """
    parts = []

    # Try pdfminer first (it can be better on some PDFs)
    if _HAS_PDFMINER:
        try:
            # pdfminer can extract whole document or page ranges; to keep parity with others,
            # we attempt to extract the first `max_pages` pages by specifying page_numbers.
            # Note: page_numbers is 0-based in pdfminer
            page_numbers = list(range(max_pages))
            txt = _pdfminer_extract_text(str(path), page_numbers=page_numbers) or ""
            if txt.strip():
                parts.append(txt)
        except Exception:
            logger.exception("pdfminer failed for %s", path)

    # Try pdfplumber
    if _HAS_PDFPLUMBER:
        try:
            with pdfplumber.open(path) as pdf:
                for i, p in enumerate(pdf.pages):
                    if i >= max_pages:
                        break
                    page_text = p.extract_text() or ""
                    if page_text.strip():
                        parts.append(page_text)
        except Exception:
            logger.exception("pdfplumber failed for %s", path)

    # Try PyPDF2 as well and append any additional text
    if _HAS_PYPDF2:
        try:
            r = PdfReader(str(path))
            for i, p in enumerate(r.pages):
                if i >= max_pages:
                    break
                try:
                    page_text = p.extract_text() or ""
                except Exception:
                    page_text = ""
                if page_text.strip():
                    parts.append(page_text)
        except Exception:
            logger.exception("PyPDF2 failed for %s", path)

    # Consolidate extracted text (avoid duplicate consecutive parts)
    if parts:
        # join with double newlines between pages to preserve paragraphs
        text = "\n\n".join(dict.fromkeys(parts))
    else:
        text = ""

    # Heuristic for readability: consider both alnum chars and word count
    alnum_chars = len(re.sub(r"[^A-Za-z0-9]", "", text))
    words = len(re.findall(r"\w+", text))
    # Relaxed thresholds: accept as readable if some text is present
    readable = (alnum_chars >= 30) or (words >= 6)
    return text.strip(), readable


def _normalize_installation_from_filename(name: str) -> Tuple[Optional[str], str]:
    """Try various heuristics (per spec) to extract an installation number from a filename.

    Returns (installation_or_None, diagnostics_string)
    """
    base = Path(name).stem.upper().strip()

    # SAT handling: if SAT present, try to map preceding MIV### to MIV###.1
    if re.search(r"\bSAT\b|\bSAT\d\b", base):
        m = re.search(r"(MIV\d{3,4})", base)
        if m:
            return m.group(1) + ".1", "mapped via SAT rule"

    # Direct MIV match
    m = re.search(r"\b(MIV\d{3,4}(?:\.\d)?)\b", base)
    if m:
        return m.group(1), "mapped via direct MIV match"

    # Prefix rule like A/C/G/WO/WW + 4 digits (allow spaces)
    m = re.search(r"\b(?:A|C|G|WO|WW)\s*\d{4}\b", base)
    if m:
        val = re.sub(r"\s+", "", m.group(0))
        return val, "mapped via prefix rule"

    # Alt pattern: 3 digits, letter A/C/G, 1 digit (allow spaces)
    m = re.search(r"(\d{3}\s*[ACG]\s*\d)", base)
    if m:
        norm = re.sub(r"\s+", "", m.group(1))
        return norm, "mapped via alt pattern (3-digit letter digit)"

    # Token heuristic
    tokens = re.split(r"[\s\-_.]+", base)
    # Only accept tokens that match known installation patterns to avoid mapping unrelated tokens
    for t in tokens:
        if re.match(r"^(?:MIV\d{3,4}(?:\.\d)?|(?:A|C|G|WO|WW)\d{4}|\d{3}[ACG]\d)$", t):
            return t, "mapped via token heuristic"

    return None, "no install found"


def extract_keuringsdatum(text: str) -> Tuple[Optional[str], Optional[str]]:
    """Return (raw_match, iso_date) or (None, None).

    Looks for patterns like 'Datum(s) van controle:18/06/2020' and general dates DD/MM/YYYY.
    """
    if not text:
        return None, None

    # Common pattern
    m = re.search(r"Datum\(s\) van controle[:\s]*([0-3]?\d/[01]?\d/\d{4})", text, flags=re.IGNORECASE)
    if not m:
        m = re.search(r"Datum van controle[:\s]*([0-3]?\d/[01]?\d/\d{4})", text, flags=re.IGNORECASE)
    if m:
        raw = m.group(1).strip()
        try:
            dt = datetime.strptime(raw, "%d/%m/%Y").date()
            return raw, dt.isoformat()
        except Exception:
            return raw, None

    # fallback: find first DD/MM/YYYY anywhere
    m = re.search(r"([0-3]?\d/[01]?\d/\d{4})", text)
    if m:
        raw = m.group(1)
        try:
            dt = datetime.strptime(raw, "%d/%m/%Y").date()
            return raw, dt.isoformat()
        except Exception:
            return raw, None

    return None, None


def extract_result(text: str) -> Tuple[Optional[str], Optional[str]]:
    """Return (raw_result_block, normalized_result).

    Strategy: find 'Besluit van de controle' and capture next 1-2 lines. Normalize to known values.
    """
    if not text:
        return None, None

    m = re.search(r"Besluit van de controle[:\s]*\n?(.{0,400})", text, flags=re.IGNORECASE | re.DOTALL)
    raw = None
    if m:
        snippet = m.group(1).strip()
        # take up to first double newline or 200 chars
        snippet = re.split(r"\n\s*\n", snippet)[0]
        raw = snippet.split("\n")[0].strip()
    else:
        # fallback: try to find lines starting with 'Besluit' anywhere
        m2 = re.search(r"\n([^\n]{0,200}Besluit[^\n]{0,200})\n", text, flags=re.IGNORECASE)
        if m2:
            raw = m2.group(1).strip()

    if not raw:
        # no clean snippet; try to detect any of the known keywords
        low = text.lower()
        if "niet-conform" in low or "niet conform" in low:
            return "", "niet-conform"
        if "conform met opmerkingen" in low:
            return "", "conform met opmerkingen"
        if "conform" in low:
            return "", "conform"
        return None, None

    normalized = None
    lr = raw.lower()
    if "conform met opmerkingen" in lr:
        normalized = "conform met opmerkingen"
    elif "niet-conform met inbreuken" in lr or "niet conform met inbreuken" in lr:
        normalized = "niet-conform met inbreuken"
    elif "niet-conform" in lr or "niet conform" in lr:
        normalized = "niet-conform"
    elif "conform" in lr:
        normalized = "conform"

    return raw, normalized


def extract_inbreuken(text: str) -> Optional[str]:
    if not text:
        return None
    m = re.search(r"vastgestelde inbreuken[^:]*[:\n]\s*(.+?)(?:\n\s*\n|$)", text, flags=re.IGNORECASE | re.DOTALL)
    if m:
        block = m.group(1).strip()
        # compress whitespace and newlines
        block = re.sub(r"\s+", " ", block)
        return block
    return None


# New: try to find installation numbers inside the PDF text when filename fails
def _normalize_installation_from_text(text: str) -> Tuple[Optional[str], str]:
    if not text:
        return None, "no text"
    t = text.upper()

    # Look for MIV patterns first (may contain .1 etc.)
    m = re.search(r"\b(MIV\s*\d{3,4}(?:\s*\.\s*\d)?)\b", t)
    if m:
        val = re.sub(r"\s+", "", m.group(1))
        return val, "mapped from text (MIV)"

    # Prefix A/C/G/WO/WW + 4 digits
    m = re.search(r"\b(?:A|C|G|WO|WW)\s*\d{4}\b", t)
    if m:
        val = re.sub(r"\s+", "", m.group(0))
        return val, "mapped from text (prefix)"

    # 3 digits letter 1 digit pattern
    m = re.search(r"(\d{3}\s*[ACG]\s*\d)", t)
    if m:
        val = re.sub(r"\s+", "", m.group(1))
        return val, "mapped from text (3-digit letter digit)"

    return None, "no install in text"


def _improve_result_detection(text: str) -> Tuple[Optional[str], Optional[str]]:
    """Augment extract_result: detect uppercase summary lines like 'DE INSTALLATIE IS NIET CONFORM' and other variants."""
    if not text:
        return None, None
    # First try the existing extract_result logic
    raw, norm = extract_result(text)
    if raw or norm:
        return raw, norm

    # Look for uppercase headline-style lines
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        up = s.upper()
        if "INSTALLATIE" in up and "NIET CONFORM" in up:
            return s, "niet-conform"
        if "INSTALLATIE" in up and "CONFORM" in up and "NIET" not in up:
            return s, "conform"

    # fallback: search for short phrases anywhere
    low = text.lower()
    if "conform met opmerkingen" in low:
        return "conform met opmerkingen", "conform met opmerkingen"
    if "niet-conform met inbreuken" in low or "niet conform met inbreuken" in low:
        return "niet-conform met inbreuken", "niet-conform met inbreuken"
    if "niet-conform" in low or "niet conform" in low:
        return "niet-conform", "niet-conform"
    if "conform" in low:
        return "conform", "conform"

    return None, None


def _improve_inbreuken_detection(text: str) -> Optional[str]:
    if not text:
        return None
    # reuse existing extraction
    got = extract_inbreuken(text)
    if got:
        return got
    # fallback: search for 'inbreuk' word and gather surrounding sentences
    low = text.lower()
    idx = low.find('inbreuk')
    if idx == -1:
        idx = low.find('inbreuken')
    if idx == -1:
        return None
    # take 200 chars around first occurrence
    start = max(0, idx - 100)
    end = min(len(text), idx + 300)
    snippet = text[start:end]
    snippet = re.sub(r"\s+", " ", snippet)
    return snippet.strip()


def process_all(pdf_dir: Path, out_csv: Path, max_pages: int = 5) -> int:
    pdf_dir = pdf_dir.resolve()
    if not pdf_dir.exists() or not pdf_dir.is_dir():
        logger.error("pdf-dir not found: %s", pdf_dir)
        return 2

    rows = []

    for p in sorted(pdf_dir.iterdir()):
        if not p.is_file() or p.suffix.lower() != ".pdf":
            continue
        logger.info("Processing %s", p.name)
        text, readable = extract_text_from_pdf(p, max_pages=max_pages)

        install, diag = _normalize_installation_from_filename(p.name)
        # If we didn't find an installation from filename, try the extracted text
        if not install:
            install_text, diag_text = _normalize_installation_from_text(text)
            if install_text:
                install = install_text
                diag = f"{diag_text}; fallback from filename"

        keurings_raw, keurings_iso = extract_keuringsdatum(text)
        # improved result detection
        resultaat_raw, resultaat_norm = _improve_result_detection(text)
        # improved inbreuken detection
        inbreuken = _improve_inbreuken_detection(text)

        # If extraction found fields, be more permissive: mark readable
        if (not readable) and (keurings_raw or resultaat_raw or inbreuken):
            readable = True

        row = {
            "pdf_filename": p.name,
            "pdf_readable": str(readable),
            "installation": install or "",
            "mapping_diagnostics": diag,
            "keuringsdatum_raw": keurings_raw or "",
            "keuringsdatum_iso": keurings_iso or "",
            "resultaat_raw": resultaat_raw or "",
            "resultaat_normalized": resultaat_norm or "",
            "inbreuken": inbreuken or "",
        }
        rows.append(row)

    # Write CSV
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(
            fh,
            fieldnames=[
                "pdf_filename",
                "pdf_readable",
                "installation",
                "mapping_diagnostics",
                "keuringsdatum_raw",
                "keuringsdatum_iso",
                "resultaat_raw",
                "resultaat_normalized",
                "inbreuken",
            ],
        )
        writer.writeheader()
        for r in rows:
            writer.writerow(r)

    logger.info("Wrote %d rows to %s", len(rows), out_csv)
    return 0


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Extract keuringsdata from PDF reports in this folder.")
    p.add_argument("--pdf-dir", type=Path, default=Path("./pdfs"), help="Directory with PDFs (default: ./pdfs)")
    p.add_argument("--out", type=Path, default=Path("extraction_results.csv"), help="Output CSV path")
    p.add_argument("--max-pages", type=int, default=5, help="Max pages to read per PDF (default 5)")
    return p


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-8s %(message)s")
    args = _build_parser().parse_args()

    base_dir = Path(__file__).resolve().parent
    pdf_dir = (base_dir / "pdfs") if str(args.pdf_dir) == "./pdfs" else args.pdf_dir
    # Determine output CSV path: if the user left the default filename, place it next to the script.
    out_arg = Path(args.out)
    if out_arg == Path("extraction_results.csv"):
        out_csv = base_dir / out_arg
    else:
        # For any explicitly provided path, respect it (resolve relative to cwd if not absolute)
        out_csv = out_arg if out_arg.is_absolute() else out_arg.resolve()

    # Warn if no pdf extraction backend available
    if not _HAS_PDFPLUMBER and not _HAS_PYPDF2:
        logger.error(
            "No PDF extraction library available. Install 'pdfplumber' or 'PyPDF2' in the venv."
        )

    raise SystemExit(process_all(pdf_dir, out_csv, max_pages=args.max_pages))
