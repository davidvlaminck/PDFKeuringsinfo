"""
Verify PDF document contents: inspection result (RED/ORANGE/GREEN/GEEN_KEURING) and inspection date.

This module extracts and classifies the compliance verdict from PDF inspection documents
(BESLUIT section) and performs date extraction from various document sections.
Reports on label mismatches, date mismatches, and content duplicates.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib
import logging
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any

import pandas as pd
from tqdm import tqdm

logger = logging.getLogger(__name__)

DEFAULT_SINCE_DATE = "1900-01-01"
DEFAULT_INPUT_DIR = Path("/home/davidlinux/Documenten/AWV/Keuringen")
DEFAULT_INPUT_FILENAME = "MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx"
DEFAULT_INPUT_PATH = DEFAULT_INPUT_DIR / DEFAULT_INPUT_FILENAME


@dataclass(frozen=True)
class DecisionExtraction:
    """Classified compliance verdict from BESLUIT section."""
    verdict: str
    confidence: str
    reason: str
    section_excerpt: str


@dataclass(frozen=True)
class InspectionDateExtraction:
    """Extracted inspection date from PDF document."""
    inspection_date: dt.date | None
    fallback_date: dt.date | None
    confidence: str
    reason: str


def _normalize_text(value: Any) -> str:
    """Normalize text value: trim, deduplicate whitespace."""
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    return re.sub(r"\s+", " ", text)


# Illegal XML 1.0 characters that openpyxl cannot write to .xlsx
_ILLEGAL_XML_CHARS_RE = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f"
    r"\ud800-\udfff\ufffe\uffff]"
)


def _sanitize_for_excel(value: Any) -> Any:
    """Strip illegal XML characters from strings so openpyxl can write them."""
    if isinstance(value, bytes):
        value = value.decode("utf-8", errors="replace")
    if not isinstance(value, str):
        return value
    return _ILLEGAL_XML_CHARS_RE.sub("", value)


def _sanitize_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Sanitize all values and headers so Excel export does not fail on illegal chars."""
    sanitized = df.copy()
    sanitized.columns = [_sanitize_for_excel(str(col)) for col in sanitized.columns]
    for col in sanitized.columns:
        sanitized[col] = sanitized[col].map(_sanitize_for_excel)
    return sanitized


def _find_first_illegal_xml_value(df: pd.DataFrame) -> tuple[int, str, str] | None:
    """Return first (row_index, column_name, snippet) containing illegal XML characters."""
    for ridx, row in df.iterrows():
        for col in df.columns:
            value = row[col]
            text = value.decode("utf-8", errors="replace") if isinstance(value, bytes) else value
            if isinstance(text, str) and _ILLEGAL_XML_CHARS_RE.search(text):
                return ridx, str(col), text[:120]
    return None


def _normalize_site_id(value: Any) -> str:
    """Normalize site ID: trim, remove .0 suffix if numeric."""
    text = _normalize_text(value)
    if text.endswith(".0") and text.replace(".", "", 1).isdigit():
        return text[:-2]
    return text


def _normalize_label(value: Any) -> str:
    """Normalize label to standard form: RED, ORANGE, GREEN, or GEEN_KEURING."""
    token = _normalize_text(value).upper().replace(" ", "")
    if not token:
        return ""

    red_tokens = {
        "RED",
        "ROOD",
        "R",
        "NIETCONFORM",
        "NONCONFORM",
        "NC",
    }
    orange_tokens = {
        "ORANGE",
        "ORANJE",
        "O",
        "CONFORMMETOPMERKINGEN",
        "CMO",
    }
    green_tokens = {
        "GREEN",
        "GROEN",
        "G",
        "CONFORM",
        "OK",
    }
    no_inspection_tokens = {
        "GEEN_KEURING",
        "GEENKEURING",
        "GEENKEURINGSVERSLAG",
        "DOCUMENTATIE",
    }

    if token in red_tokens:
        return "RED"
    if token in orange_tokens:
        return "ORANGE"
    if token in green_tokens:
        return "GREEN"
    if token in no_inspection_tokens:
        return "GEEN_KEURING"
    return ""


def _extract_pdf_text(pdf_path: Path) -> str:
    """Extract all text from PDF file, concatenating pages."""
    PdfReader = _get_pdf_reader()
    reader = PdfReader(str(pdf_path))
    parts: list[str] = []
    for page in reader.pages:
        try:
            parts.append(page.extract_text() or "")
        except Exception:
            parts.append("")
    return "\n".join(parts)


@lru_cache(maxsize=1)
def _get_pdf_reader():
    """Dynamically import pypdf.PdfReader. Cached to avoid repeated imports."""
    try:
        pdf_module = importlib.import_module("pypdf")
        return getattr(pdf_module, "PdfReader")
    except Exception as exc:  # pragma: no cover - runtime env specific
        raise RuntimeError("Dependency 'pypdf' ontbreekt. Run eerst: uv sync") from exc


def _extract_besluit_section(pdf_text: str) -> str:
    """Extract the BESLUIT section from the PDF text.

    Handles two known layout types:
    - Type 1: "BESLUIT" (bold/underlined header, bullet points with dash)
    - Type 2: "Besluit" (mixed-case header, followed by checkmark "V" and short text)
    """
    if not pdf_text:
        return ""

    # Case-insensitive match on "BESLUIT" as a standalone word/heading.
    match = re.search(r"(?im)^\s*BESLUIT\s*$", pdf_text)
    if not match:
        # Fallback: BESLUIT anywhere (e.g. inline on a line with other text)
        match = re.search(r"(?i)\bBESLUIT\b", pdf_text)
    if not match:
        return ""

    tail = pdf_text[match.end():]

    # Stop at the next all-caps section heading or page boundary.
    stop_match = re.search(
        r"\n\s*(?:[A-Z][A-Z0-9 /().\-]{4,})\s*\n",
        tail,
    )
    if stop_match and stop_match.start() > 20:
        section = tail[: stop_match.start()]
    else:
        section = tail[:2500]

    section = re.sub(r"[\r\t]+", " ", section)
    section = re.sub(r"\n+", "\n", section)
    return section.strip()


# ---------------------------------------------------------------------------
# Classification Patterns
#
# Evaluation order:
#   1. RED   – "voldoet niet aan" / "niet in overeenstemming"
#   2. GREEN base (voldoet aan) + remarks → ORANGE
#   3. GREEN base without remarks → GREEN
#   4. UNKNOWN fallback
# ---------------------------------------------------------------------------

# Patterns that unambiguously signal non-compliance.
_RED_PATTERNS = [
    r"voldoet\s+niet\s+aan",
    r"niet\s+in\s+overeenstemming",
    r"niet\s+conform",
    r"niet\s+teruggevonden",
    r"niet\s+gevonden",
    r"niet\s+worden\s+gecontroleerd",
    r"niet\s+gecontroleerd",
    r"geen\s+keuring\s+uitgevoerd",
    r"nog\s+niets\s+om\s+te\s+keuren",
    r"niet\s+toegankelijk",
    r"bord\s+leeg",
    r"geen\s+bord\s+gevonden",
    r"kon\s+niet\s+uitgevoerd\s+worden",
    r"geen\s+ingevulde\s+gegevens",
    r"geen\s+besluit\s+genomen",
    r"geen\s+controle\s+uitgevoerd",
    r"geen\s+controle\s+mogelijk",
    r"controle\s+niet\s+mogelijk",
    r"niet\s+bereikbaar",
    r"geen\s+kast\s+aanwezig",
    # Encoded variants seen in some PDFs:
    # "NIET WORDEN GECONTROLEERD" and "kon niet worden gecontroleerd"
    r"OJFU[\x00-\x1f\s]*XPSEFO[\x00-\x1f\s]*HFDPOUSPMFFSE",
    r"\]WRc[\x00-\x1f\s]*f\^aQR\][\x00-\x1f\s]*URP\^\]ca\^ZRRaQ",
]

_ORANGE_OVERRIDE_PATTERNS = [
    r"cabine\s+is\s+reeds\s+afgebroken",
]

# Patterns for the "voldoet aan" base (compliance confirmed).
# Covers Type 1 ("hierboven vermelde voorschriften") and
# Type 2 ("voorschriften van Boek 2 …").
_GREEN_BASE_PATTERNS = [
    r"voldoet\s+aan\s+de\s+(?:hierboven\s+vermelde\s+)?voorschriften",
    r"voldoet\s+aan\s+de\s+voorschriften\s+van\b",   # Type 2: "Boek 2 …"
    r"\bvolledig\s+conform\b",
]

# Patterns that signal remarks / recommendations alongside compliance → ORANGE.
_REMARK_PATTERNS = [
    r"opmerk(?:ing(?:en)?)?",
    r"aanbevel(?:ing(?:en)?)?",
    r"dient\s+gevolg\s+gegeven",
    r"voorbehoud",
]

_NO_REMARKS_PATTERNS = [
    r"opmerk(?:ing(?:en)?)?\s*:\s*(?:-|\u2022)?\s*nihil\b",
    r"geen\s+opmerk(?:ing(?:en)?)?\b",
]

_NIHIL_INBREUKEN_PATTERN = r"vastgestelde\s+inbreuken\s*:\s*(?:-|\u2022)?\s*nihil\b"

# Patterns that appear as boilerplate in the checkbox-style template and are NOT informative.
# In this layout both "VOLDOET AAN" and "VOLDOET NIET AAN" always appear; only one is checked.
_CHECKBOX_TEMPLATE_BOILERPLATE_RED = frozenset({
    r"voldoet\s+niet\s+aan",
    r"niet\s+in\s+overeenstemming",
})
# Some PDFs use a custom font encoding where extracted text looks shifted/garbled.
_ENCODED_INBREUKEN_PATTERN = r"[W7]BTUHFTUFMEF[\x00-\x1f]*JOCSFVLFO"
_ENCODED_NIHIL_PATTERN = r"@WVWZ|OJIJM"
_ENCODED_OPMERKINGEN_PATTERN = r"0QNFSLJOHFO|PQNFSLJOHFO"
_ENCODED_OPMERKINGEN_NIHIL_PATTERN = (
    r"(?:0QNFSLJOHFO|PQNFSLJOHFO)[\x00-\x1f\s:;,\-]*(?:@WVWZ|OJIJM)"
)

_NO_INSPECTION_PATTERNS = [
    r"\bsynthesedocument\b",
    r"bijstand\s+in\s+het\s+kader\s+van\s+de\s+risicoanalyse",
    r"\bsamenvatting\s+zonder\s+conclusie\b",
    r"documentatie\s+arei",
]


def _has_any_pattern(normalized_text: str, patterns: list[str]) -> bool:
    """Check if any pattern matches the text."""
    if not normalized_text:
        return False
    return any(re.search(pat, normalized_text, flags=re.IGNORECASE) for pat in patterns)


def _has_remarks_marker(normalized_text: str) -> bool:
    """Check if text contains remark/recommendation markers."""
    if not normalized_text:
        return False
    for pat in _REMARK_PATTERNS:
        if re.search(pat, normalized_text, flags=re.IGNORECASE):
            return True
    return False


def _has_no_remarks_marker(normalized_text: str) -> bool:
    """Check if text explicitly states no remarks (nihil)."""
    if not normalized_text:
        return False
    for pat in _NO_REMARKS_PATTERNS:
        if re.search(pat, normalized_text, flags=re.IGNORECASE):
            return True
    return False


def _is_checkbox_template(normalized_besluit: str) -> bool:
    """Return True when the BESLUIT section contains both "VOLDOET AAN" and "VOLDOET NIET AAN".

    This is a fixed checkbox template where both options are always printed and only one is
    visually checked (☑/□).  Text extraction cannot distinguish which box is ticked so the
    presence of "VOLDOET NIET AAN" must not be treated as a RED signal in this layout.
    """
    has_voldoet = bool(
        re.search(
            r"de\s+elektrische\s+installatie\s+voldoet\s+aan\s+de",
            normalized_besluit,
            re.IGNORECASE,
        )
    )
    has_voldoet_niet = bool(
        re.search(
            r"de\s+elektrische\s+installatie\s+voldoet\s+niet\s+aan\s+de",
            normalized_besluit,
            re.IGNORECASE,
        )
    )
    return has_voldoet and has_voldoet_niet


def _extract_pre_besluit_section(full_text: str) -> str:
    """Return the part of the document that precedes the BESLUIT heading.

    This typically contains the INBREUKEN / OPMERKINGEN sections.
    """
    match = re.search(r"(?im)^\s*BESLUIT\s*$", full_text)
    if not match:
        match = re.search(r"(?i)\bBESLUIT\b", full_text)
    if not match:
        return full_text
    return full_text[: match.start()]


def _is_no_inspection_document(normalized_full: str) -> bool:
    """Check if document is a summary/documentation without actual inspection result."""
    if not normalized_full:
        return False
    return any(re.search(pat, normalized_full, flags=re.IGNORECASE) for pat in _NO_INSPECTION_PATTERNS)


def _hash_pdf_text(pdf_text: str) -> str:
    """Compute SHA256 hash of normalized PDF text for duplicate detection."""
    normalized = _normalize_text(pdf_text).casefold()
    if not normalized:
        return ""
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


_DATE_TOKEN_RE = r"(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*((?:19|20)\d{2}|\d{2,3})"
_DATE_TOKEN_NAMED_RE = r"(?P<day>\d{1,2})\s*/\s*(?P<month>\d{1,2})\s*/\s*(?P<year>(?:19|20)\d{2}|\d{2,3})"
_DATE_TOKEN_COMPILED = re.compile(_DATE_TOKEN_NAMED_RE)
_DATE_RANGE_SEP_RE = r"(?:tot|to)"
_NOISE_DATE_LINE_RE = re.compile(r"(?i)uitgiftedatum|koninklijk\s+besluit\s+van")


def _infer_year_from_truncated_token(
    year_token: str,
    *,
    visit_date: dt.date | None,
    day: int,
    month: int,
) -> int | None:
    """Infer 4-digit year from token (may be 2, 3, or 4 digits) using visit_date as hint."""
    if len(year_token) == 4:
        return int(year_token)

    if len(year_token) == 2:
        base = 2000 + int(year_token)
        if visit_date and visit_date.day == day and visit_date.month == month and abs(visit_date.year - base) <= 1:
            return visit_date.year
        return base

    if len(year_token) != 3:
        return None

    if visit_date and visit_date.day == day and visit_date.month == month:
        return visit_date.year

    candidates: set[int] = set()
    for pos in range(4):
        for digit in "0123456789":
            built = f"{year_token[:pos]}{digit}{year_token[pos:]}"
            if len(built) != 4:
                continue
            year_int = int(built)
            if 1900 <= year_int <= 2099:
                candidates.add(year_int)

    if not candidates:
        return None
    if visit_date:
        return min(candidates, key=lambda y: (abs(y - visit_date.year), abs(y - 2000)))
    return min(candidates, key=lambda y: abs(y - 2000))


def _parse_dayfirst_date_token(
    token: str,
    *,
    visit_date: dt.date | None,
) -> dt.date | None:
    """Parse date token in DD/MM/YYYY format (possibly truncated year)."""
    m = re.search(_DATE_TOKEN_RE, token)
    if not m:
        return None

    day = int(m.group(1))
    month = int(m.group(2))
    year_token = m.group(3)
    year = _infer_year_from_truncated_token(year_token, visit_date=visit_date, day=day, month=month)
    if year is None:
        return None

    try:
        return dt.date(year, month, day)
    except ValueError:
        return None


def _extract_inspection_date(pdf_text: str, *, visit_date: dt.date | None) -> InspectionDateExtraction:
    """Extract inspection date from PDF document.

    Tries multiple patterns:
    1. Interventiegegevens section (date range or single date)
    2. Uitgevoerd op section
    3. BESLUIT section (single date)
    4. Fallback: all dates in document (if unique)
    """
    if not pdf_text or not pdf_text.strip():
        return InspectionDateExtraction(
            inspection_date=None,
            fallback_date=None,
            confidence="LOW",
            reason="Geen leesbare tekst voor datumextractie",
        )

    patterns: list[tuple[str, str, str]] = [
        (
            r"(?is)interventiegegevens[\s\S]{0,420}?\bdatum\b\s*[:\-]?\s*(?:van|vanaf)?\s*"
            + _DATE_TOKEN_RE
            + r"\s*"
            + _DATE_RANGE_SEP_RE
            + r"\s*"
            + _DATE_TOKEN_RE,
            "Interventiegegevens range-datum gevonden",
            "HIGH",
        ),
        (
            r"(?is)\buitgevoerd\s+op\b\s*[:\-]?\s*(?:vanaf\s*)?"
            + _DATE_TOKEN_RE
            + r"(?:\s*\d{1,2}:\d{2})?(?:\s*[\-–]\s*\d{1,2}:\d{2})?\s*"
            + _DATE_RANGE_SEP_RE
            + r"\s*"
            + _DATE_TOKEN_RE,
            "Uitgevoerd op range-datum gevonden",
            "HIGH",
        ),
        (
            r"(?is)interventiegegevens[\s\S]{0,420}?\bdatum\b\s*[:\-]?\s*(?:van|vanaf)?\s*"
            + _DATE_TOKEN_RE,
            "Interventiegegevens datum gevonden",
            "HIGH",
        ),
        (
            r"(?is)\buitgevoerd\s+op\b\s*[:\-]?\s*(?:vanaf\s*)?" + _DATE_TOKEN_RE,
            "Uitgevoerd op datum gevonden",
            "HIGH",
        ),
    ]

    for pattern, reason, confidence in patterns:
        m = re.search(pattern, pdf_text, flags=re.IGNORECASE)
        if not m:
            continue

        groups = m.groups()
        # First date groups are always indexes 0..2, optional second date 3..5.
        primary_token = "/".join([groups[0], groups[1], groups[2]])
        primary_date = _parse_dayfirst_date_token(primary_token, visit_date=visit_date)
        fallback_date = None
        if len(groups) >= 6:
            fallback_token = "/".join([groups[3], groups[4], groups[5]])
            fallback_date = _parse_dayfirst_date_token(fallback_token, visit_date=visit_date)

        if primary_date:
            return InspectionDateExtraction(primary_date, fallback_date, confidence, reason)
        if fallback_date:
            return InspectionDateExtraction(
                fallback_date,
                None,
                "MEDIUM",
                reason + " (fallback op einddatum)",
            )

    besluit_text = _extract_besluit_section(pdf_text)
    if besluit_text:
        m_besluit = re.search(r"(?is)\bdatum\b\s*[:\-]?\s*" + _DATE_TOKEN_RE, besluit_text, flags=re.IGNORECASE)
        if m_besluit:
            token = "/".join([m_besluit.group(1), m_besluit.group(2), m_besluit.group(3)])
            parsed = _parse_dayfirst_date_token(token, visit_date=visit_date)
            if parsed:
                return InspectionDateExtraction(parsed, None, "HIGH", "Datum gevonden in BESLUIT-sectie")

    filtered_lines = [line for line in pdf_text.splitlines() if not _NOISE_DATE_LINE_RE.search(line)]
    filtered_text = "\n".join(filtered_lines)
    date_hits = []
    for m in _DATE_TOKEN_COMPILED.finditer(filtered_text):
        token = f"{m.group('day')}/{m.group('month')}/{m.group('year')}"
        parsed = _parse_dayfirst_date_token(token, visit_date=visit_date)
        if parsed is not None:
            date_hits.append(parsed)

    unique_hits = sorted(set(date_hits))
    if len(unique_hits) == 1:
        return InspectionDateExtraction(unique_hits[0], None, "MEDIUM", "Unieke datum in document")
    if len(unique_hits) > 1:
        return InspectionDateExtraction(None, None, "LOW", "Meerdere datums in document, geen eenduidige inspectiedatum")

    return InspectionDateExtraction(None, None, "LOW", "Geen datum gevonden in PDF")


def _classify_besluit_text(section_text: str, full_text: str) -> DecisionExtraction:
    """Classify the BESLUIT section into RED / ORANGE / GREEN / UNKNOWN.

    Known document types:
    - Type 1 (Vincotte standard): BESLUIT header, dash-bullet sentences.
      RED:    "voldoet niet aan … voorschriften"
      ORANGE: "voldoet aan … voorschriften" + "opmerkingen/aanbevelingen"
      GREEN:  "voldoet aan … voorschriften" (no remarks bullet)
    - Type 2 (compact layout): "Besluit" header, "V" checkmark, short sentence.
      GREEN:  "voldoet aan de voorschriften van Boek 2 …"
    """
    target = section_text if section_text else full_text
    normalized = _normalize_text(target).lower()
    normalized_full = _normalize_text(full_text).lower()

    if _is_no_inspection_document(normalized_full):
        return DecisionExtraction(
            verdict="GEEN_KEURING",
            confidence="HIGH",
            reason="Documenttype zonder keuringsbesluit (samenvatting/documentatie)",
            section_excerpt=_normalize_text(target)[:400],
        )

    if _has_any_pattern(normalized, [r"geen\s+keuring\s+uitgevoerd"]) or _has_any_pattern(
        normalized_full, [r"geen\s+keuring\s+uitgevoerd"]
    ):
        if (_has_remarks_marker(normalized) and not _has_no_remarks_marker(normalized)) or (
            _has_remarks_marker(normalized_full) and not _has_no_remarks_marker(normalized_full)
        ):
            return DecisionExtraction(
                verdict="ORANGE",
                confidence="HIGH",
                reason="Geen keuring uitgevoerd, maar wel opmerkingen aanwezig",
                section_excerpt=_normalize_text(target)[:400],
            )

    if not normalized:
        return DecisionExtraction(
            verdict="GEEN_KEURING",
            confidence="LOW",
            reason="Geen leesbare tekst in PDF (waarschijnlijk geen keuringsverslag)",
            section_excerpt="",
        )

    excerpt = _normalize_text(target)[:400]

    # Detect checkbox template *before* running RED patterns.
    is_checkbox_template = _is_checkbox_template(normalized)

    # Step 1 – RED: explicit non-compliance (first inside BESLUIT, then full document).
    for pat in _RED_PATTERNS:
        if is_checkbox_template and pat in _CHECKBOX_TEMPLATE_BOILERPLATE_RED:
            # "VOLDOET NIET AAN" is boilerplate in this layout; skip it as RED signal.
            continue
        if re.search(pat, normalized, flags=re.IGNORECASE):
            return DecisionExtraction(
                verdict="RED",
                confidence="HIGH",
                reason=f"Niet-conform indicator: /{pat}/",
                section_excerpt=excerpt,
            )
        if re.search(pat, normalized_full, flags=re.IGNORECASE):
            return DecisionExtraction(
                verdict="RED",
                confidence="MEDIUM",
                reason=f"Niet-conform indicator buiten BESLUIT: /{pat}/",
                section_excerpt=excerpt,
            )

    # Step 1b – Checkbox template: determine verdict via the INBREUKEN section.
    if is_checkbox_template:
        pre_besluit = _normalize_text(_extract_pre_besluit_section(full_text)).lower()
        # Heuristic: "Nihil" directly under the INBREUKEN/OPMERKINGEN heading means no infractions.
        inbreuken_nihil = bool(
            re.search(
                r"inbreuken.*?nihil|nota.*?nihil|\bnihil\b",
                pre_besluit,
                re.IGNORECASE | re.DOTALL,
            )
        )
        if inbreuken_nihil:
            # In this template "Nihil" appears under one combined heading that covers
            # INBREUKEN, OPMERKINGEN *and* NOTA'S.  The word "opmerkingen" in that
            # heading must not be treated as an actual remark, so we return GREEN
            # directly instead of running the full-document remarks scan.
            return DecisionExtraction(
                verdict="GREEN",
                confidence="MEDIUM",
                reason="Checkbox template: INBREUKEN/OPMERKINGEN/NOTA'S = Nihil",
                section_excerpt=excerpt,
            )
        # Infractions present in INBREUKEN section → non-compliant.
        return DecisionExtraction(
            verdict="RED",
            confidence="MEDIUM",
            reason="Checkbox template: INBREUKEN bevat inbreuken (geen Nihil)",
            section_excerpt=excerpt,
        )

    # Step 2 – Explicit ORANGE overrides found in remarks sections.
    for pat in _ORANGE_OVERRIDE_PATTERNS:
        if re.search(pat, normalized, flags=re.IGNORECASE) or re.search(pat, normalized_full, flags=re.IGNORECASE):
            return DecisionExtraction(
                verdict="ORANGE",
                confidence="MEDIUM",
                reason=f"Conform met opmerkingen op basis van /{pat}/",
                section_excerpt=excerpt,
            )

    # Step 2 – GREEN base present?
    green_pat_hit = None
    for pat in _GREEN_BASE_PATTERNS:
        if re.search(pat, normalized, flags=re.IGNORECASE):
            green_pat_hit = pat
            break

    if green_pat_hit:
        section_has_remarks = _has_remarks_marker(normalized) and not _has_no_remarks_marker(normalized)
        full_has_remarks = _has_remarks_marker(normalized_full) and not _has_no_remarks_marker(normalized_full)
        # Step 2a – upgrade to ORANGE if remarks are mentioned
        if section_has_remarks or full_has_remarks:
            confidence = "HIGH" if section_has_remarks else "MEDIUM"
            reason = "Conform maar met opmerkingen"
            if full_has_remarks and not section_has_remarks:
                reason += " (opmerkingen buiten BESLUIT gedetecteerd)"
            return DecisionExtraction(
                verdict="ORANGE",
                confidence=confidence,
                reason=reason,
                section_excerpt=excerpt,
            )
        # Step 2b – no remarks → GREEN
        return DecisionExtraction(
            verdict="GREEN",
            confidence="HIGH",
            reason=f"Conform zonder opmerkingen (/{green_pat_hit}/)",
            section_excerpt=excerpt,
        )

    # Step 3 – Alternate conformity signal in some layouts: "Vastgestelde inbreuken: Nihil".
    if re.search(_NIHIL_INBREUKEN_PATTERN, normalized, flags=re.IGNORECASE):
        section_has_remarks = _has_remarks_marker(normalized) and not _has_no_remarks_marker(normalized)
        full_has_remarks = _has_remarks_marker(normalized_full) and not _has_no_remarks_marker(normalized_full)
        if section_has_remarks or full_has_remarks:
            reason = "Conform op basis van 'Vastgestelde inbreuken: Nihil', met opmerkingen"
            if full_has_remarks and not section_has_remarks:
                reason += " (opmerkingen buiten BESLUIT gedetecteerd)"
            return DecisionExtraction(
                verdict="ORANGE",
                confidence="MEDIUM",
                reason=reason,
                section_excerpt=excerpt,
            )
        return DecisionExtraction(
            verdict="GREEN",
            confidence="MEDIUM",
            reason="Conform op basis van 'Vastgestelde inbreuken: Nihil'",
            section_excerpt=excerpt,
        )

    # Step 4 - fallback for garbled custom-font extraction.
    # Example markers seen in extracted text:
    #   VASTGESTELDE INBREUKEN -> "7BTUHFTUFMEFJOCSFVLFO"
    #   NIHIL -> "@WVWZ"
    if re.search(_ENCODED_INBREUKEN_PATTERN, normalized, flags=re.IGNORECASE) and re.search(
        _ENCODED_NIHIL_PATTERN, normalized, flags=re.IGNORECASE
    ):
        encoded_has_remarks = re.search(_ENCODED_OPMERKINGEN_PATTERN, normalized, flags=re.IGNORECASE)
        encoded_remarks_are_nil = re.search(_ENCODED_OPMERKINGEN_NIHIL_PATTERN, normalized, flags=re.IGNORECASE)
        if encoded_has_remarks and not encoded_remarks_are_nil:
            return DecisionExtraction(
                verdict="ORANGE",
                confidence="LOW",
                reason="Conform op basis van gecodeerde 'inbreuken: nihil'-markers, met opmerkingen",
                section_excerpt=excerpt,
            )
        return DecisionExtraction(
            verdict="GREEN",
            confidence="LOW",
            reason="Conform op basis van gecodeerde 'inbreuken: nihil'-markers",
            section_excerpt=excerpt,
        )

    # Encoded variant where inbreuken are present and not marked as nihil.
    if re.search(_ENCODED_INBREUKEN_PATTERN, normalized_full, flags=re.IGNORECASE) and not re.search(
        _ENCODED_NIHIL_PATTERN, normalized_full, flags=re.IGNORECASE
    ):
        return DecisionExtraction(
            verdict="RED",
            confidence="LOW",
            reason="Niet-conform op basis van gecodeerde 'vastgestelde inbreuken'-markers (geen nihil)",
            section_excerpt=excerpt,
        )

    return DecisionExtraction(
        verdict="UNKNOWN",
        confidence="LOW",
        reason="Geen herkenbaar besluitpatroon gevonden",
        section_excerpt=excerpt,
    )


def _find_pdf_path(
    base_dir: Path,
    site_id: str,
    filename: str,
    site_file_index: dict[str, dict[str, Path]] | None = None,
) -> Path | None:
    """Find PDF file in site directory, with case-insensitive and .pdf extension fallback."""
    if not site_id or not filename:
        return None

    site_dir = base_dir / site_id
    if not site_dir.exists() or not site_dir.is_dir():
        return None

    raw_name = filename.strip()
    names = [raw_name]
    if not raw_name.lower().endswith(".pdf"):
        names.append(f"{raw_name}.pdf")

    for name in names:
        candidate = site_dir / name
        if candidate.exists() and candidate.is_file():
            return candidate

    lower_map = site_file_index.get(site_id) if site_file_index is not None else None
    if lower_map is None:
        lower_map = {p.name.lower(): p for p in site_dir.iterdir() if p.is_file()}
    for name in names:
        found = lower_map.get(name.lower())
        if found is not None:
            return found

    return None


def _build_site_file_index(base_dir: Path, site_ids: set[str]) -> dict[str, dict[str, Path]]:
    """Pre-build lowercase filename index per site directory for efficient lookup."""
    index: dict[str, dict[str, Path]] = {}
    for site_id in site_ids:
        if not site_id:
            continue
        site_dir = base_dir / site_id
        if not site_dir.exists() or not site_dir.is_dir():
            continue
        index[site_id] = {p.name.lower(): p for p in site_dir.iterdir() if p.is_file()}
    return index


def _resolve_column(df: pd.DataFrame, candidates: list[str], fallback_index: int | None = None) -> str:
    """Resolve DataFrame column by name (case-insensitive), with optional fallback by index."""
    normalized = {str(col).strip().lower(): col for col in df.columns}
    for candidate in candidates:
        match = normalized.get(candidate.strip().lower())
        if match is not None:
            return str(match)

    if fallback_index is not None and 0 <= fallback_index < len(df.columns):
        return str(df.columns[fallback_index])

    raise ValueError(f"Kolom niet gevonden. Gezochte headers: {candidates}")


def verify_pdf_contents_and_date(
    *,
    input_excel: Path,
    output_excel: Path,
    since_date: str,
    skip_rows: int = 0,
    limit_rows: int | None = None,
) -> dict[str, int]:
    """Verify PDF inspection result (verdict) and date against enriched Excel.

    For each row in the input Excel:
    1. Extract decision/verdict from PDF BESLUIT section
    2. Extract inspection date from PDF document
    3. Compare verdict to enriched label
    4. Compare date to enriched visit date
    5. Report mismatches and content duplicates

    Returns statistics dictionary.
    """
    df = pd.read_excel(input_excel)

    site_col = _resolve_column(df, ["Site ID", "SiteID"], fallback_index=0)
    filename_col = _resolve_column(df, ["FILENAME", "Filename"], fallback_index=11)
    label_col = _resolve_column(df, ["Label"])
    visit_col = _resolve_column(df, ["Visit date", "Visit Date"])

    visit_dates = pd.to_datetime(df[visit_col], errors="coerce", dayfirst=True)
    since_ts = pd.Timestamp(since_date)

    result_rows: list[dict[str, Any]] = []
    stats = {
        "processed": 0,
        "date_filtered_out": 0,
        "pdf_missing": 0,
        "unreadable": 0,
        "no_inspection": 0,
        "content_duplicates": 0,
        "matched": 0,
        "mismatch": 0,
        "unknown": 0,
        "date_matched": 0,
        "date_mismatch": 0,
        "date_unknown": 0,
    }

    base_dir = input_excel.parent

    valid_visit_mask = visit_dates.notna() & (visit_dates > since_ts)
    eligible_positions = [int(pos) for pos in valid_visit_mask[valid_visit_mask].index]
    if skip_rows and skip_rows > 0:
        eligible_positions = eligible_positions[skip_rows:]
    if limit_rows is not None and limit_rows > 0:
        eligible_positions = eligible_positions[:limit_rows]
    stats["date_filtered_out"] = int((~valid_visit_mask).sum())

    pbar = tqdm(total=len(eligible_positions), desc="PDF-inhoud en datum verificatie", unit="rij")
    site_ids = {
        _normalize_site_id(df.iloc[pos].get(site_col))
        for pos in eligible_positions
    }
    site_file_index = _build_site_file_index(base_dir, site_ids)

    def _update_progress_postfix() -> None:
        pbar.set_postfix(
            {
                "processed": stats["processed"],
                "match": stats["matched"],
                "mismatch": stats["mismatch"],
                "pdf_missing": stats["pdf_missing"],
                "geen_keuring": stats["no_inspection"],
                "unknown": stats["unknown"],
                "date_mismatch": stats["date_mismatch"],
            },
            refresh=False,
        )

    for pos in eligible_positions:
        row = df.iloc[pos]
        excel_row_index = pos + 2
        visit_ts = visit_dates.iloc[pos]
        site_id = _normalize_site_id(row.get(site_col))
        filename = _normalize_text(row.get(filename_col))
        expected_label_raw = row.get(label_col)
        expected_label = _normalize_label(expected_label_raw)

        pdf_path = _find_pdf_path(base_dir, site_id, filename, site_file_index=site_file_index)
        if pdf_path is None:
            stats["processed"] += 1
            stats["pdf_missing"] += 1
            _update_progress_postfix()
            pbar.update(1)
            result_rows.append(
                {
                    "row_index": excel_row_index,
                    "site_id": site_id,
                    "filename": filename,
                    "visit_date": visit_ts.date().isoformat(),
                    "label_raw": expected_label_raw,
                    "label_expected": expected_label or "UNKNOWN",
                    "label_from_pdf": "UNKNOWN",
                    "verify_status": "PDF_MISSING",
                    "confidence": "LOW",
                    "reason": "PDF niet gevonden in site-map",
                    "pdf_path": "",
                    "besluit_excerpt": "",
                }
            )
            continue

        try:
            full_text = _extract_pdf_text(pdf_path)
            besluit_text = _extract_besluit_section(full_text)
            extraction = _classify_besluit_text(besluit_text, full_text)
            date_extraction = _extract_inspection_date(full_text, visit_date=visit_ts.date() if pd.notna(visit_ts) else None)
            pdf_content_hash = _hash_pdf_text(full_text)
        except Exception as exc:
            stats["processed"] += 1
            stats["unreadable"] += 1
            stats["no_inspection"] += 1
            _update_progress_postfix()
            pbar.update(1)
            label_expected_out = expected_label or "GEEN_KEURING"
            result_rows.append(
                {
                    "row_index": excel_row_index,
                    "site_id": site_id,
                    "filename": filename,
                    "visit_date": visit_ts.date().isoformat(),
                    "label_raw": expected_label_raw,
                    "label_expected": label_expected_out,
                    "label_from_pdf": "GEEN_KEURING",
                    "verify_status": "GEEN_KEURING",
                    "confidence": "LOW",
                    "reason": f"PDF lezen mislukt (behandeld als geen keuringsverslag): {exc}",
                    "pdf_path": str(pdf_path),
                    "pdf_content_hash": "",
                    "content_duplicate": False,
                    "content_duplicate_count": 0,
                    "duplicate_filenames": "",
                    "besluit_excerpt": "",
                    "inspection_date_pdf": "",
                    "inspection_date_pdf_fallback": "",
                    "date_verify_status": "UNKNOWN",
                    "date_confidence": "LOW",
                    "date_reason": "Datumextractie overgeslagen door leesfout",
                    "overall_verify_status": "UNKNOWN",
                }
            )
            continue

        visit_date_obj = visit_ts.date() if pd.notna(visit_ts) else None
        pdf_date = date_extraction.inspection_date
        pdf_date_fallback = date_extraction.fallback_date

        if visit_date_obj is None or pdf_date is None:
            date_verify_status = "UNKNOWN"
            stats["date_unknown"] += 1
        elif pdf_date == visit_date_obj or (pdf_date_fallback is not None and pdf_date_fallback == visit_date_obj):
            date_verify_status = "MATCH"
            stats["date_matched"] += 1
        else:
            date_verify_status = "MISMATCH"
            stats["date_mismatch"] += 1

        predicted = extraction.verdict
        label_expected_out = expected_label or ("GEEN_KEURING" if predicted == "GEEN_KEURING" else "UNKNOWN")
        if predicted == "GEEN_KEURING":
            verify_status = "GEEN_KEURING"
            stats["no_inspection"] += 1
        elif not expected_label or predicted == "UNKNOWN":
            verify_status = "UNKNOWN"
            stats["unknown"] += 1
        elif expected_label == predicted:
            verify_status = "MATCH"
            stats["matched"] += 1
        else:
            verify_status = "MISMATCH"
            stats["mismatch"] += 1

        stats["processed"] += 1
        _update_progress_postfix()
        pbar.update(1)

        if verify_status in {"PDF_MISSING", "UNKNOWN"}:
            overall_verify_status = "UNKNOWN"
        elif verify_status == "GEEN_KEURING":
            overall_verify_status = "GEEN_KEURING"
        elif verify_status == "MISMATCH" or date_verify_status == "MISMATCH":
            overall_verify_status = "MISMATCH"
        elif verify_status == "MATCH" and date_verify_status == "MATCH":
            overall_verify_status = "MATCH"
        else:
            overall_verify_status = "UNKNOWN"

        result_rows.append(
            {
                "row_index": excel_row_index,
                "site_id": site_id,
                "filename": filename,
                "visit_date": visit_ts.date().isoformat(),
                "label_raw": expected_label_raw,
                "label_expected": label_expected_out,
                "label_from_pdf": predicted,
                "verify_status": verify_status,
                "confidence": extraction.confidence,
                "reason": extraction.reason,
                "pdf_path": str(pdf_path),
                "pdf_content_hash": pdf_content_hash,
                "content_duplicate": False,
                "content_duplicate_count": 0,
                "duplicate_filenames": "",
                "besluit_excerpt": extraction.section_excerpt,
                "inspection_date_pdf": pdf_date.isoformat() if pdf_date else "",
                "inspection_date_pdf_fallback": pdf_date_fallback.isoformat() if pdf_date_fallback else "",
                "date_verify_status": date_verify_status,
                "date_confidence": date_extraction.confidence,
                "date_reason": date_extraction.reason,
                "overall_verify_status": overall_verify_status,
            }
        )

    pbar.close()

    content_groups: dict[str, list[int]] = {}
    for idx, row in enumerate(result_rows):
        content_hash = str(row.get("pdf_content_hash", "") or "")
        if content_hash:
            content_groups.setdefault(content_hash, []).append(idx)

    for indices in content_groups.values():
        if len(indices) <= 1:
            continue
        stats["content_duplicates"] += len(indices)
        full_names = []
        for idx in indices:
            site_id = str(result_rows[idx].get("site_id", "") or "")
            filename = str(result_rows[idx].get("filename", "") or "")
            full_names.append(f"{site_id}/{filename}" if site_id else filename)
        for idx in indices:
            row = result_rows[idx]
            current_site = str(row.get("site_id", "") or "")
            current_name = str(row.get("filename", "") or "")
            current_full = f"{current_site}/{current_name}" if current_site else current_name
            duplicates = [name for name in full_names if name != current_full]
            row["content_duplicate"] = True
            row["content_duplicate_count"] = len(indices)
            row["duplicate_filenames"] = "; ".join(sorted(dict.fromkeys(duplicates)))

    result_rows.sort(key=lambda x: x.get("row_index", 0))

    result_df = pd.DataFrame(result_rows)
    # Sanitize everything defensively before writing to xlsx.
    result_df = _sanitize_dataframe_for_excel(result_df)
    offending = _find_first_illegal_xml_value(result_df)
    if offending is not None:
        ridx, col, snippet = offending
        filename_hint = ""
        if "filename" in result_df.columns and 0 <= ridx < len(result_df):
            filename_hint = f", filename='{result_df.iloc[ridx].get('filename', '')}'"
        raise RuntimeError(
            f"Illegale XML-karakters gedetecteerd op result-row {ridx}, kolom '{col}'{filename_hint}."
            f" Waarde-start: {snippet!r}"
        )

    try:
        result_df.to_excel(output_excel, index=False)
    except Exception as exc:
        filename_hint = ""
        if "filename" in result_df.columns and not result_df.empty:
            filename_hint = f"; voorbeeld filename='{result_df.iloc[0].get('filename', '')}'"
        raise RuntimeError(f"Excel export mislukt ({output_excel}){filename_hint}: {exc}") from exc
    csv_path = output_excel.with_suffix(".csv")
    result_df.to_csv(csv_path, index=False)

    return stats


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser(
        description=(
            "Verifieer PDF-inhoud (inspectieresultaat + datum) tegen enriched Excel. "
            "Vergelijkt de verdict uit het BESLUIT-gedeelte en de inspectiedatum uit het PDF "
            "met de waarden in de enriched Excel."
        )
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help=f"Pad naar enriched Excel-bestand (default: {DEFAULT_INPUT_PATH})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Pad naar outputrapport (.xlsx). Default: <input>.pdf-contents-check.xlsx",
    )
    parser.add_argument(
        "--since-date",
        default=DEFAULT_SINCE_DATE,
        help="Verwerk enkel records met Visit date > deze datum (YYYY-MM-DD)",
    )
    parser.add_argument(
        "--skip",
        type=int,
        default=0,
        help="Optioneel: sla de eerste N in-scope records over (na datumfilter).",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Optioneel: verwerk enkel de eerste N in-scope records (na datumfilter)",
    )
    return parser.parse_args()


def main() -> int:
    """Main entry point."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    args = parse_args()
    input_excel = (args.input or DEFAULT_INPUT_PATH).resolve()
    if not input_excel.exists():
        # Fallback: look for any *.enriched.xlsx in the default dir
        candidates = sorted(DEFAULT_INPUT_DIR.glob("*.enriched.xlsx"))
        if candidates:
            input_excel = candidates[0].resolve()
            logger.info("Geen input opgegeven, gebruikt: %s", input_excel)
        else:
            logger.error("Inputbestand bestaat niet: %s", input_excel)
            return 1

    output_excel = args.output.resolve() if args.output else input_excel.with_name(f"{input_excel.stem}.pdf-contents-check.xlsx")

    try:
        stats = verify_pdf_contents_and_date(
            input_excel=input_excel,
            output_excel=output_excel,
            since_date=args.since_date,
            skip_rows=args.skip,
            limit_rows=args.limit,
        )
    except Exception as exc:
        logger.error("Verificatie mislukt: %s", exc)
        return 2

    logger.info("Rapport geschreven: %s", output_excel)
    logger.info("CSV geschreven: %s", output_excel.with_suffix('.csv'))
    logger.info("Stats: %s", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

