"""
Extract keuringresultaat (conform / conform-met-opmerkingen / inbreuken) and
keuringsdatum from PDF inspection documents referenced in the
20260917_ElektrischeKeuringen_datakwaliteit_verbeteringen.xlsx Excel file.

Approach:
1. Read the Excel sheet "nok_eigenschappen_ontbreken".
2. For each row, locate the referenced PDF in the bijlagen/ folder.
3. Detect the PDF structure type.
4. Extract verdict + date using structure-specific patterns.
5. Only report results with HIGH confidence (100% certain).
6. Output to CSV for review; do NOT overwrite existing Excel values.

Structure types detected:
  VIN_STD        - Vincotte legacy: "VERSLAG VAN PERIODIEKE CONTROLE" +
                   Interventiegegevens Datum + BESLUIT "voldoet aan/niet aan"
  VIN_STD_MET    - Vincotte MET: "DE INSTALLATIE IS CONFORM" +
                   "Besluit van de controle:" + "Datum(s) van controle:"
  VIN_PREFIX     - BTV/Vlaamse Overheid: "DE INSTALLATIE IS CONFORM" +
                   "Datum van het onderzoek" + "PROCES-VERBAAL"
  VIN_CONCLUSION - Bilingual: "Besluit / Conclusion:" + "Datum van bezoek"
  VIN_STD_NEW    - Vincotte HS-cabine: "Besluit" + "voldoet aan Boek 2" +
                   "Uitgevoerd op: DD/MM/YYYY"
  TRAFIROAD      - TRAFIROAD: "onderzochte elektrische installatie" +
                   "Datum van onderzoek : DD.MM.YYYY"
  OCB            - OCB: "BESLUIT BC" + "Datum van onderzoek:"
  TECH_BUREAU    - Tech Bureau Verbrugghen one-pagers
  VIN_STD_ENCODED - Vincotte with garbled custom font encoding
  UNKNOWN        - Any other structure
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import sys
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

import pandas as pd
from tqdm import tqdm

try:
    from pypdf import PdfReader
except ImportError:
    from PyPDF2 import PdfReader


EXCEL_PATH = Path("ElektrischeKeuring_rapport.xlsx")
SHEET_NAME = "assets"
BIJLAGEN_DIR = Path().home() / 'Downloads/Documenten_Keuring'

VERDICT_CONFORM = "conform"
VERDICT_CONFORM_MET_OPMERKINGEN = "conform-met-opmerkingen"
VERDICT_INBREUKEN = "inbreuken"


@dataclass
class ExtractionResult:
    excel_row: int
    site_id: str
    filename: str
    pdf_path: str
    structure: str
    verdict: str
    verdict_confidence: str
    verdict_reason: str
    inspection_date: str
    date_confidence: str
    date_reason: str
    text_length: int
    has_besluit: bool


# ---------------------------------------------------------------------------
# Filename matching
# ---------------------------------------------------------------------------

def _aggressive_norm(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9]", "", name).lower()


def build_file_index(directory: Path) -> dict[str, Path]:
    index: dict[str, Path] = {}
    for f in directory.iterdir():
        if not f.is_file():
            continue
        key = _aggressive_norm(f.name)
        if key not in index:
            index[key] = f
    return index


def find_pdf(
    directory: Path,
    name_a: Optional[str],
    name_b: Optional[str],
    index: dict[str, Path],
) -> Optional[Path]:
    for raw_name in (name_a, name_b):
        if raw_name is None or (isinstance(raw_name, float) and pd.isna(raw_name)):
            continue
        raw = str(raw_name).strip()
        if not raw:
            continue
        candidate = directory / raw
        if candidate.is_file():
            return candidate
        key = _aggressive_norm(raw)
        if key in index:
            return index[key]
        base = raw.rstrip(".PDF").rstrip(".pdf")
        for variant in (base + ".pdf", base + ".PDF"):
            key2 = _aggressive_norm(variant)
            if key2 in index:
                return index[key2]
        tokens = sorted(re.split(r"[^a-zA-Z0-9]+", key))
        if len(tokens) < 2:
            continue
        for f in directory.iterdir():
            ftokens = sorted(re.split(r"[^a-zA-Z0-9]+", _aggressive_norm(f.name)))
            if tokens == ftokens:
                return f
    return None


# ---------------------------------------------------------------------------
# PDF text extraction
# ---------------------------------------------------------------------------

def extract_pdf_text(pdf_path: Path) -> str:
    try:
        reader = PdfReader(str(pdf_path))
    except Exception:
        return ""
    parts: list[str] = []
    for page in reader.pages:
        try:
            parts.append(page.extract_text() or "")
        except Exception:
            parts.append("")
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Structure detection
# ---------------------------------------------------------------------------

@dataclass
class StructureInfo:
    name: str


def _looks_encoded(text: str) -> bool:
    lower = text.lower()
    if re.search(r"@wvwz|ojijm", lower):
        return True
    if re.search(r"[w7]btuhftu.*jo?csfvlfo", lower):
        return True
    if "HRabZMU" in text or "7&34" in text:
        return True
    if "voldoet" not in lower and "conform" not in lower and "besluit" not in lower:
        if re.search(r"[a-z][A-Z][a-z][A-Z][a-z][A-Z][a-z][A-Z]", text):
            return True
    return False


def detect_structure(text: str) -> StructureInfo:
    if not text or not text.strip():
        return StructureInfo("EMPTY")

    lower = text.lower()

    if "onderzochte elektrische installatie" in lower:
        return StructureInfo("TRAFIROAD")

    if re.search(r"verslag\s+van\s+(?:periodieke?\s*(?:controle|controlebezoek|controlebezoeken)|"
                 r"gelijkvormigheidsstudie|gelijkvormigheids\s*onderzoek|eerste\s+controle|"
                 r"vergelijkende\s+controle|cONTROLE)",
                 lower):
        return StructureInfo("VIN_STD")

    if "besluit" in lower and re.search(r"voldoet\s+(?:niet\s+)?aan\s+de\s+voorschriften\s+van\s+boek\s+2", lower):
        return StructureInfo("VIN_STD_NEW")

    if re.search(r"besluit\s*/\s*conclusion", lower) or "conclusion" in lower:
        return StructureInfo("VIN_CONCLUSION")

    if re.search(r"besluit\s+bc\d", lower) or (r"o\.c\.b\." in lower and "besluit" in lower):
        return StructureInfo("OCB")

    if "besluit van de controle" in lower:
        return StructureInfo("VIN_MET")

    if "proces-verbaal" in lower or "proces verbaal" in lower:
        return StructureInfo("VIN_PREFIX")

    if "technisch bureau" in lower or "verbruggen" in lower:
        return StructureInfo("TECH_BUREAU")

    if _looks_encoded(text):
        return StructureInfo("VIN_STD_ENCODED")

    if "bord" in lower and ("min. ipx" in lower or "ipx4" in lower or "min ipx" in lower):
        return StructureInfo("OCB_BOARD")

    if "interventiegegevens" in lower and "besluit" in lower and ("voldoet" in lower or "conform" in lower):
        return StructureInfo("VIN_STD")

    return StructureInfo("UNKNOWN")


# ---------------------------------------------------------------------------
# Verdict classification
# ---------------------------------------------------------------------------

_RED_PATTERNS_NL = [
    r"voldoet\s+niet\s+aan",
    r"niet\s+in\s+overeenstemming",
    r"niet\s+conform",
    r"niet\s+teruggevonden",
    r"niet\s+gevonden",
    r"niet\s+worden\s+gecontroleerd",
    r"niet\s+gecontroleerd",
    r"geen\s+keuring\s+uitgevoerd",
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
    r"installatie\s+is\s+niet\s+conform",
]

_GREEN_PATTERNS_NL = [
    r"voldoet\s+aan\s+de\s+(?:hierboven\s+vermelde\s+)?voorschriften",
    r"voldoet\s+aan\s+de\s+voorschriften\s+van\b",
    r"\bvolledig\s+conform\b",
    r"de\s+installatie\s+is\s+conform",
    r"installatie\s+is\s+conform\s+aan\s+de",
    r"is\s+conform\s+aan\s+de\s+hoger\s+vermelde",
]

_GREEN_NO_INBREUKEN = re.compile(r"vastgestelde\s+inbreuken\s*:\s*(?:-|\u2022)?\s*nihil", re.IGNORECASE)

_REMARK_PATTERNS = [
    r"opmerking(?:en)?",
    r"aanbeveling(?:en)?",
    r"dient\s+gevolg\s+gegeven",
    r"voorbehoud",
    r"behoudens\s+voormelde\s+inbreuk",
]

_NO_REMARKS_PATTERNS = [
    r"opmerking(?:en)?\s*[:\-]?\s*(?:-|\u2022)?\s*nihil\b",
    r"geen\s+opmerking(?:en)?\b",
    r"opmerking(?:en)?\s*\n\s*nihil\b",
]


def _has_remarks(text: str) -> bool:
    if not text:
        return False
    cleaned = re.sub(r"dient\s+gevolg\s+gegeven\s+te\s+worden\s+aan\s+de\s+opmerking\w*", "", text, flags=re.IGNORECASE)
    return any(re.search(pat, cleaned, re.IGNORECASE) for pat in _REMARK_PATTERNS)


def _has_no_remarks(text: str) -> bool:
    if not text:
        return False
    return any(re.search(pat, text, re.IGNORECASE) for pat in _NO_REMARKS_PATTERNS)


def _is_checkbox_template(text: str) -> bool:
    has_voldoet = bool(re.search(r"de\s+elektrische\s+installatie\s+voldoet\s+aan\s+de", text, re.IGNORECASE))
    has_niet = bool(re.search(r"de\s+elektrische\s+installatie\s+voldoet\s+niet\s+aan\s+de", text, re.IGNORECASE))
    return has_voldoet and has_niet


def _extract_besluit_section(text: str) -> str:
    match = re.search(r"(?im)^\s*BESLUIT\s*$", text)
    if not match:
        for m in re.finditer(r"(?i)\bBESLUIT\b", text):
            pre = text[:m.start()].lower()
            if "koninklijk besluit" in pre[-50:]:
                continue
            match = m
            break
    if not match:
        return ""
    tail = text[match.end():]
    stop = re.search(r"\n\s*(?:[A-Z][A-Z0-9 /().\-]{4,})\s*\n", tail)
    if not stop:
        stop = re.search(r"\n\s*(?:[A-Z][A-Z0-9 /().\-]{4,})\s*\n", tail, re.IGNORECASE)
    if stop and stop.start() > 20:
        return tail[:stop.start()].strip()
    return tail[:2500].strip()


def _classify_verdict(text: str, structure: str) -> tuple[str, str, str, bool]:
    if not text or not text.strip():
        return "", "LOW", "Geen leesbare tekst", False

    lower = text.lower()
    has_besluit = "besluit" in lower

    if structure == "VIN_STD_ENCODED":
        return _classify_encoded(text)
    if structure == "TRAFIROAD":
        return _classify_trafiroad(lower, has_besluit)
    if structure in ("OCB", "OCB_BOARD"):
        return _classify_ocb(lower, has_besluit)
    if structure == "VIN_PREFIX":
        return _classify_prefix(lower, has_besluit)
    if structure == "VIN_MET":
        return _classify_met(lower, has_besluit)
    if structure == "VIN_CONCLUSION":
        return _classify_conclusion(lower, has_besluit)
    if structure == "VIN_STD_NEW":
        return _classify_std_new(lower, has_besluit)
    if structure == "VIN_STD":
        return _classify_std(lower, has_besluit)
    if structure == "TECH_BUREAU":
        return _classify_tech_bureau(lower, has_besluit)
    return _classify_generic(lower, has_besluit)


def _classify_std(lower: str, has_besluit: bool) -> tuple[str, str, str, bool]:
    """VIN_STD: BESLUIT section is authoritative source for verdict."""
    beslout = _extract_besluit_section_for_class(lower)
    lower_beslout = beslout.lower() if beslout else lower

    is_checkbox = _is_checkbox_template(lower_beslout)

    for pat in _RED_PATTERNS_NL:
        hit_section = re.search(pat, lower_beslout, re.IGNORECASE)
        if hit_section:
            if is_checkbox and pat in ("voldoet niet aan", "niet in overeenstemming"):
                continue
            return VERDICT_INBREUKEN, "HIGH", f"Niet-conform: /{pat}/", True

    if is_checkbox:
        pre = _extract_pre_besluit(lower)
        if re.search(r"inbreuken.*?nihil|nota.*?nihil|\bnihil\b", pre, re.IGNORECASE | re.DOTALL):
            if _has_remarks(lower_beslout) and not _has_no_remarks(lower_beslout):
                return VERDICT_CONFORM_MET_OPMERKINGEN, "MEDIUM", "Checkbox: conform + opmerkingen", True
            return VERDICT_CONFORM, "MEDIUM", "Checkbox: INBREUKEN=Nihil", True
        return VERDICT_INBREUKEN, "MEDIUM", "Checkbox: INBREUKEN bevat inbreuken", True

    green_hit = any(re.search(pat, lower_beslout, re.IGNORECASE) for pat in _GREEN_PATTERNS_NL)
    if green_hit:
        section_has_remarks = _has_remarks(lower_beslout) and not _has_no_remarks(lower_beslout)
        if section_has_remarks:
            return VERDICT_CONFORM_MET_OPMERKINGEN, "HIGH", "Conform met opmerkingen", True
        return VERDICT_CONFORM, "HIGH", "Conform zonder opmerkingen", True

    # Fallback: green in full text + remarks
    green_in_full = any(re.search(pat, lower, re.IGNORECASE) for pat in _GREEN_PATTERNS_NL)
    full_has_remarks = _has_remarks(lower) and not _has_no_remarks(lower)
    if green_in_full and full_has_remarks:
        return VERDICT_CONFORM_MET_OPMERKINGEN, "MEDIUM", "Conform met opmerkingen (full text)", True

    # Fallback: red patterns in full text
    for pat in _RED_PATTERNS_NL:
        if re.search(pat, lower, re.IGNORECASE) and not is_checkbox:
            return VERDICT_INBREUKEN, "MEDIUM", f"Niet-conform (buiten BESLUIT): /{pat}/", True

    return "", "LOW", "Geen herkenbaar verdict", has_besluit


def _classify_std_new(lower: str, has_besluit: bool) -> tuple[str, str, str, bool]:
    """VIN_STD_NEW: 'Besluit' + 'voldoet aan Boek 2' + INBREUKEN section."""
    idx = lower.find("besluit")
    beslout_snippet = ""
    while idx >= 0:
        ctx = lower[idx:idx + 800]
        if "voldoet" in ctx:
            beslout_snippet = ctx
            break
        idx = lower.find("besluit", idx + 1)
    if not beslout_snippet:
        beslout_snippet = re.sub(r'\s+', ' ', lower[:2000])
    else:
        beslout_snippet = re.sub(r'\s+', ' ', beslout_snippet)

    is_checkbox = ("voldoet aan" in beslout_snippet and "voldoet niet aan" in beslout_snippet)

    if "voldoet niet aan" in beslout_snippet and not is_checkbox:
        return VERDICT_INBREUKEN, "HIGH", "Niet-conform: voldoet niet aan", True

    if "voldoet aan" in beslout_snippet and not is_checkbox:
        if _GREEN_NO_INBREUKEN.search(lower):
            if _has_remarks(beslout_snippet) and not _has_no_remarks(lower):
                return VERDICT_CONFORM_MET_OPMERKINGEN, "HIGH", "Conform met opmerkingen", True
            return VERDICT_CONFORM, "HIGH", "Conform zonder inbreuken", True
        inb_idx = lower.find("vastgestelde inbreuken")
        if inb_idx >= 0:
            inb_ctx = lower[inb_idx:inb_idx + 200]
            if "nihil" not in inb_ctx and "-" not in inb_ctx[:20]:
                return VERDICT_INBREUKEN, "MEDIUM", "Inbreuken gevonden", True
        if _has_remarks(beslout_snippet) and not _has_no_remarks(beslout_snippet):
            return VERDICT_CONFORM_MET_OPMERKINGEN, "HIGH", "Conform met opmerkingen", True
        return VERDICT_CONFORM, "MEDIUM", "Conform zonder duidelijke inbreuken", True

    if is_checkbox:
        inb_match = re.search(r"(?:inbreuken|afdeling\s+inbreuken)\s*\(i\).*?\n.*?\n", lower, re.IGNORECASE | re.DOTALL)
        if not inb_match:
            inb_match = re.search(r"(?:inbreuken\s*\(i\))(.*?)oplossingen|conform", lower, re.IGNORECASE | re.DOTALL)
            if not inb_match:
                inb_match = re.search(r"vastgestelde\s+inbreuken\s*[\s\S]{0,500}", lower, re.IGNORECASE | re.DOTALL)
        if inb_match:
            inb_text = inb_match.group(0)[:500] if inb_match.lastindex else inb_match.group(0)
            if "nihil" in inb_text.lower():
                if _has_remarks(lower) and not _has_no_remarks(lower):
                    return VERDICT_CONFORM_MET_OPMERKINGEN, "LOW", "Checkbox: inbreuken=nihil + opmerkingen", True
                return VERDICT_CONFORM, "LOW", "Checkbox: inbreuken=nihil", True
            has_listed_inbreuken = bool(re.search(r'\([io]\)\s*\w', inb_text)) or len(inb_text.strip()) > 30
            if has_listed_inbreuken:
                return VERDICT_INBREUKEN, "LOW", "Checkbox: inbreuken gelist", True
            return VERDICT_CONFORM, "LOW", "Checkbox: geen duidelijke inbreuken", True
        return "", "LOW", "Checkbox template: geen inbreuken sectie gevonden", True

    return "", "LOW", "Geen herkenbaar verdict in Besluit", True


def _classify_prefix(lower: str, has_besluit: bool) -> tuple[str, str, str, bool]:
    """VIN_PREFIX: 'DE INSTALLATIE IS CONFORM/NIET CONFORM'."""
    if re.search(r"is\s+niet\s+conform", lower):
        return VERDICT_INBREUKEN, "HIGH", "Niet-conform: installatie is niet conform", True
    if re.search(r"is\s+conform", lower):
        if _has_remarks(lower) and not _has_no_remarks(lower):
            return VERDICT_CONFORM_MET_OPMERKINGEN, "MEDIUM", "Conform met opmerkingen", True
        return VERDICT_CONFORM, "HIGH", "Conform: installatie is conform", True
    return "", "LOW", "Geen VIN_PREFIX verdict", has_besluit


def _classify_met(lower: str, has_besluit: bool) -> tuple[str, str, str, bool]:
    """VIN_MET: 'Besluit van de controle:' + 'voldoet aan voorschriften'."""
    idx = lower.find("besluit van de controle")
    if idx < 0:
        idx = lower.find("besluit")
    beslout_snippet = lower[idx:idx + 800] if idx >= 0 else lower

    if "voldoet niet aan" in beslout_snippet:
        return VERDICT_INBREUKEN, "HIGH", "Niet-conform: voldoet niet aan", True
    if re.search(r"voldoet\s+aan\s+de\s+voorschriften", beslout_snippet):
        if _has_remarks(beslout_snippet) and not _has_no_remarks(beslout_snippet):
            return VERDICT_CONFORM_MET_OPMERKINGEN, "HIGH", "Conform met opmerkingen", True
        return VERDICT_CONFORM, "HIGH", "Conform: voldoet aan de voorschriften", True
    return "", "LOW", "Geen VIN_MET verdict", True


def _classify_conclusion(lower: str, has_besluit: bool) -> tuple[str, str, str, bool]:
    """VIN_CONCLUSION: 'Besluit / Conclusion:' + bilingual verdict."""
    idx = lower.find("besluit")
    if idx < 0:
        idx = lower.find("conclusion")
    beslout_snippet = lower[idx:idx + 600] if idx >= 0 else lower

    if "voldoet niet aan" in beslout_snippet or "non-conforme" in beslout_snippet:
        return VERDICT_INBREUKEN, "HIGH", "Niet-conform: voldoet niet aan", True
    if re.search(r"voldoet\s+aan\s+de\s+voorschriften", beslout_snippet) or "est conforme" in beslout_snippet:
        if _has_remarks(beslout_snippet) and not _has_no_remarks(beslout_snippet):
            return VERDICT_CONFORM_MET_OPMERKINGEN, "HIGH", "Conform met opmerkingen", True
        return VERDICT_CONFORM, "HIGH", "Conform: voldoet aan de voorschriften", True
    return "", "LOW", "Geen VIN_CONCLUSION verdict", True


def _classify_trafiroad(lower: str, has_besluit: bool) -> tuple[str, str, str, bool]:
    """TRAFIROAD: 'De onderzochte elektrische installatie ... voldoet...'"""
    idx = lower.find("besluit")
    while idx >= 0 and idx < len(lower):
        ctx = lower[idx:idx + 800]
        if "onderzochte" in ctx or re.search(r"voldoet\s+(?:,?\s*behoudens|aan)", ctx):
            break
        idx = lower.find("besluit", idx + 1)
    beslout_snippet = lower[idx:idx + 800] if idx >= 0 else lower

    if re.search(r"voldoet\s+niet\s+aan", beslout_snippet):
        return VERDICT_INBREUKEN, "HIGH", "Niet-conform: voldoet niet aan", True

    if re.search(r"voldoet[^.]{0,30}behoudens\s+voormelde\s+inbreuk", beslout_snippet):
        return VERDICT_CONFORM_MET_OPMERKINGEN, "HIGH", "Conform met inbreuken", True

    if re.search(r"voldoet[^.]{0,30}(?:aan\s+voormelde\s+voorschriften|,?\s*rekening\s+houdend)", beslout_snippet):
        inb_section = re.search(r"0?7\.?\s*inbreuken\s*\n\s*\n(.+?)(?:\n\s*0?8\.|\Z)", lower, re.DOTALL)
        if not inb_section:
            inb_section = re.search(r"0?7\.?\s*inbreuken\s*(?:[\s\S]{1,100}?)\n\s*-?\s*(.+?)(?:\n\s*0?8\.|\Z)", lower, re.DOTALL)
        if inb_section:
            inb_text = inb_section.group(1).strip()
            if re.match(r"^(geen|n\.v\.t\.|nvt|-|–)$", inb_text.strip(), re.IGNORECASE):
                return VERDICT_CONFORM, "HIGH", "Conform: geen inbreuken", True
            if len(inb_text) > 10:
                return VERDICT_INBREUKEN, "MEDIUM", "Inbreuken in INBREUKEN sectie", True
            return VERDICT_CONFORM, "HIGH", "Conform: voldoet aan voorschriften (geen inbreuken in sectie)", True
        if not re.search(r"\binbreuk\b.*\binbreuken\b", lower):
            return VERDICT_CONFORM, "HIGH", "Conform: voldoet aan voorschriften, geen inbreuken genoemd", True
        return VERDICT_CONFORM, "MEDIUM", "Conform: voldoet aan voorschriften", True

    return "", "LOW", "Geen herkenbaar TRAFIROAD verdict", True


def _classify_ocb(lower: str, has_besluit: bool) -> tuple[str, str, str, bool]:
    idx = lower.find("besluit bc")
    if idx < 0:
        idx = lower.find("besluit")
    beslout_snippet = lower[idx:idx + 600] if idx >= 0 else lower

    if "niet conform" in beslout_snippet or "voldoet niet aan" in beslout_snippet:
        return VERDICT_INBREUKEN, "HIGH", "Niet-conform", True
    if re.search(r"is\s+conform\s+aan\s+de\s+hoger\s+vermelde", beslout_snippet) or "conform aan de hoger vermelde" in beslout_snippet:
        if _has_remarks(lower) and not _has_no_remarks(lower):
            return VERDICT_CONFORM_MET_OPMERKINGEN, "MEDIUM", "Conform met opmerkingen", True
        return VERDICT_CONFORM, "HIGH", "Conform: is conform aan hoger vermelde voorschriften", True
    return "", "LOW", "Geen OCB verdict", True


def _classify_tech_bureau(lower: str, has_besluit: bool) -> tuple[str, str, str, bool]:
    if re.search(r"vastgestelde\s+inbreuken\s*[:\-]?\s*(?:-|\u2022)?\s*nihil", lower, re.IGNORECASE):
        if _has_remarks(lower) and not _has_no_remarks(lower):
            return VERDICT_CONFORM_MET_OPMERKINGEN, "LOW", "Conform met opmerkingen (encoded)", True
        return VERDICT_CONFORM, "MEDIUM", "Conform: vastgestelde inbreuken = nihil", True
    if re.search(r"vastgeste[\w~]*\s*inbreuken|vastgestelde\s+inbreuken", lower):
        if not re.search(r"nihil|n\.v\.t\.|geen", lower):
            return VERDICT_INBREUKEN, "LOW", "Inbreuken gevonden (gedeeltelijk gelezen)", True
    if "is niet conform" in lower:
        return VERDICT_INBREUKEN, "HIGH", "Niet-conform: installatie is niet conform", True
    if "is conform" in lower:
        if _has_remarks(lower) and not _has_no_remarks(lower):
            return VERDICT_CONFORM_MET_OPMERKINGEN, "MEDIUM", "Conform met opmerkingen", True
        return VERDICT_CONFORM, "MEDIUM", "Conform: installatie is conform", True
    return "", "LOW", "Geen TECH_BUREAU verdict", has_besluit


def _classify_encoded(text: str) -> tuple[str, str, str, bool]:
    """Classify verdict from encoded Vincotte text using encoded patterns."""
    lower = text.lower()

    enc_inbreuken = bool(re.search(r"[w7]btuhftu.*jo?csfvlfo", lower))
    enc_nihil = bool(re.search(r"@wvwz|ojijm", lower))
    enc_remarks = bool(re.search(r"0qnfsljohfo|pqnfsljohfo", lower))
    enc_remarks_nihil = bool(re.search(
        r"(?:0qnfsljohfo|pqnfsljohfo)[\s\x00-\x1f:;,\-]*(?:@wvwz|ojijm)", lower))

    if enc_inbreuken and not enc_nihil:
        return VERDICT_INBREUKEN, "LOW", "Encoded: vastgestelde inbreuken zonder nihil", True
    if enc_inbreuken and enc_nihil:
        if enc_remarks and not enc_remarks_nihil:
            return VERDICT_CONFORM_MET_OPMERKINGEN, "LOW", "Encoded: inbreuken=nihil + opmerkingen", True
        return VERDICT_CONFORM, "LOW", "Encoded: vastgestelde inbreuken = nihil", True
    return "", "LOW", "Geen herkenbaar encoded verdict", True


def _classify_generic(lower: str, has_besluit: bool) -> tuple[str, str, str, bool]:
    for pat in _RED_PATTERNS_NL:
        if re.search(pat, lower, re.IGNORECASE):
            if _is_checkbox_template(lower) and pat in ("voldoet niet aan", "niet in overeenstemming"):
                continue
            return VERDICT_INBREUKEN, "HIGH" if has_besluit else "MEDIUM", f"Niet-conform: /{pat}/", has_besluit
    for pat in _GREEN_PATTERNS_NL:
        if re.search(pat, lower, re.IGNORECASE):
            if _has_remarks(lower) and not _has_no_remarks(lower):
                return VERDICT_CONFORM_MET_OPMERKINGEN, "HIGH", "Conform met opmerkingen", has_besluit
            return VERDICT_CONFORM, "HIGH", f"Conform: /{pat}/", has_besluit
    if re.search(r"is\s+conform", lower):
        return VERDICT_CONFORM, "MEDIUM", "Conform: is conform", has_besluit
    if re.search(r"is\s+niet\s+conform", lower):
        return VERDICT_INBREUKEN, "MEDIUM", "Niet-conform: is niet conform", has_besluit
    return "", "LOW", "Geen herkenbaar verdict", has_besluit


def _extract_besluit_section_for_class(lower: str) -> str:
    """Extract BESLUIT section for classification (re-derive from full text)."""
    return _extract_besluit_section(lower)


def _extract_pre_besluit(lower: str) -> str:
    """Return text before BESLUIT heading."""
    match = re.search(r"(?im)^\s*BESLUIT\s*$", lower)
    if not match:
        match = re.search(r"(?i)\bBESLUIT\b", lower)
    if not match:
        return lower
    return lower[:match.start()]


# ---------------------------------------------------------------------------
# Date extraction
# ---------------------------------------------------------------------------

_DATE_NL = r"(?P<day>\d{1,2})\s*/\s*(?P<month>\d{1,2})\s*/\s*(?P<year>(?:19|20)\d{2}|\d{2,3})"
_DATE_DOT = r"(?P<day>\d{1,2})\s*\.\s*(?P<month>\d{1,2})\s*\.\s*(?P<year>\d{2,4})"


def _parse_date_named(m: re.Match) -> Optional[str]:
    day = int(m.group('day'))
    month = int(m.group('month'))
    year_tok = m.group('year')
    if len(year_tok) == 4:
        year = int(year_tok)
    elif len(year_tok) == 2:
        yy = int(year_tok)
        year = (1900 + yy) if yy >= 30 else (2000 + yy)
    elif len(year_tok) == 3:
        return None
    else:
        return None
    try:
        return dt.date(year, month, day).isoformat()
    except ValueError:
        return None


def extract_inspection_date(text: str, structure: str, filename: str = "") -> tuple[str, str, str]:
    if not text or not text.strip():
        return "", "LOW", "Geen leesbare tekst voor datumextractie"

    # Structure-specific
    if structure == "VIN_STD":
        m = re.search(r"interventiegegevens[\s\S]{0,500}?\bdatum\s*[:\-]?\s*(?:van(?:af)?\s*)?" + _DATE_NL, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "HIGH", "Interventiegegevens Datum"
    elif structure == "VIN_STD_NEW":
        m = re.search(r"uitgevoerd\s+op\s*[:\-]?\s*(?:vanaf\s*)?" + _DATE_NL, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "HIGH", "Uitgevoerd op datum"
        m = re.search(r"interventiegegevens[\s\S]{0,500}?\bdatum\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "HIGH", "Interventiegegevens Datum (VIN_STD_NEW fallback)"
        m = re.search(r"(?im)^\s*datum\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "MEDIUM", "Generic 'Datum:' in VIN_STD_NEW"
    elif structure == "VIN_MET":
        m = re.search(r"datum\s*\(s\)\s*van\s+controle\s*[:\-]?\s*[\s\S]{0,800}?(" + _DATE_NL + ")", text, re.IGNORECASE)
        if not m:
            m = re.search(r"datum\s*\(s\)\s*van\s+controle\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "HIGH", "Datum(s) van controle"
    elif structure == "VIN_PREFIX":
        m = re.search(r"datum\s+van\s+(?:het\s+)?onderzoek\s*[:\-]?\s*[\s\S]{0,800}?(" + _DATE_NL + ")", text, re.IGNORECASE)
        if not m:
            m = re.search(r"datum\s+van\s+(?:het\s+)?onderzoek\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "HIGH", "Datum van het onderzoek"
    elif structure == "VIN_CONCLUSION":
        m = re.search(r"datum\s+van\s+bezoek\s*/\s*date\s+de\s+visite\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "HIGH", "Datum van bezoek"
    elif structure == "TRAFIROAD":
        m = re.search(r"datum\s+van\s+onderzoek\s*[:\-]?\s*" + _DATE_DOT, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "HIGH", "Datum van onderzoek (TRAFIROAD)"
    elif structure in ("OCB", "OCB_BOARD"):
        m = re.search(r"datum\s+(?:van\s+(?:het\s+)?)?onderzoek\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "HIGH", "Datum van onderzoek (OCB)"
    elif structure == "TECH_BUREAU":
        m = re.search(r"datum\s+(?:van\s+(?:het\s+)?)?onderzoek\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "MEDIUM", "Datum van onderzoek (TECH_BUREAU)"
        m = re.search(r"datum\s+van\s+(?:het\s+)?onderzoek\s*[:\-]?\s*" + _DATE_DOT, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "MEDIUM", "Datum van onderzoek (TECH_BUREAU, dot)"
        m = re.search(r"begin\s+werken\s*[:\-]?\s*" + _DATE_DOT, text, re.IGNORECASE)
        if m:
            iso = _parse_date_named(m)
            if iso:
                return iso, "MEDIUM", "Begin werken datum (TECH_BUREAU)"

    # Generic fallbacks
    m = re.search(r"interventiegegevens[\s\S]{0,500}?datum\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
    if m:
        iso = _parse_date_named(m)
        if iso:
            return iso, "MEDIUM", "Interventiegegevens Datum (generic)"

    m = re.search(r"interventiegegevens[\s\S]{0,500}?datum\s*[:\-]?\s*" + _DATE_DOT, text, re.IGNORECASE)
    if m:
        iso = _parse_date_named(m)
        if iso:
            return iso, "MEDIUM", "Interventiegegevens Datum (generic, dot)"

    m = re.search(r"(?im)^\s*datum\s*[:\-]?\s*" + _DATE_NL, text)
    if m:
        iso = _parse_date_named(m)
        if iso:
            return iso, "MEDIUM", "Generic 'Datum:' field"

    m = re.search(r"(?im)^\s*datum\s*[:\-]?\s*" + _DATE_DOT, text)
    if m:
        iso = _parse_date_named(m)
        if iso:
            return iso, "MEDIUM", "Generic 'Datum:' field (dot)"

    m = re.search(r"datum\s+van\s+(?:het\s+)?onderzoek\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
    if m:
        iso = _parse_date_named(m)
        if iso:
            return iso, "MEDIUM", "Datum van onderzoek (generic)"

    m = re.search(r"datum\s+van\s+bezoek\s*[:\-]?\s*" + _DATE_NL, text, re.IGNORECASE)
    if m:
        iso = _parse_date_named(m)
        if iso:
            return iso, "MEDIUM", "Datum van bezoek (generic)"

    # Last resort: unique date
    dates = []
    for dm in re.finditer(_DATE_NL, text):
        iso = _parse_date_named(dm)
        if iso:
            dates.append(iso)
    # Also dotted dates
    for dm in re.finditer(_DATE_DOT, text):
        iso = _parse_date_named(dm)
        if iso:
            dates.append(iso)

    # Filter known noise
    noise = {"2019-09-08", "2019-10-28", "1993-08-12", "1993-12-04", "1987-01-01"}
    dates = [d for d in dates if d not in noise]
    unique = sorted(set(dates))
    if len(unique) == 1:
        return unique[0], "MEDIUM", "Unieke datum in document"
    if len(unique) > 1:
        return "", "LOW", f"Meerdere datums ({len(unique)} uniek), geen eenduidige inspectiedatum"

    if structure == "VIN_STD_ENCODED" and filename:
        fdate = _extract_date_from_filename(filename)
        if fdate:
            return fdate, "LOW", "Datum uit bestandsnaam"
    return "", "LOW", "Geen datum gevonden in PDF"


def _extract_date_from_filename(filename: str) -> Optional[str]:
    """Try to extract a date from filename patterns like YY-MM-DD."""
    m = re.search(r"(?P<yy>\d{2})[-_.](?P<mm>\d{2})[-_.](?P<dd>\d{2})", filename)
    if m:
        yy = int(m.group('yy'))
        year = (1900 + yy) if yy >= 30 else (2000 + yy)
        try:
            return dt.date(year, int(m.group('mm')), int(m.group('dd'))).isoformat()
        except ValueError:
            pass
    return None


# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

def process_excel(excel_path: Path, output_path: Path, bijlagen_dir: Path,
                  limit: int | None = None, skip: int = 0) -> dict[str, int]:
    df = pd.read_excel(excel_path, sheet_name=SHEET_NAME, dtype=str)
    df = df.fillna("")

    file_index = build_file_index(bijlagen_dir)

    results: list[ExtractionResult] = []
    stats: dict[str, Any] = {
        "total": len(df),
        "processed": 0,
        "pdf_found": 0,
        "pdf_not_found": 0,
        "skipped_no_name": 0,
        "skipped_corrupt": 0,
        "structure_counts": {},
        "verdict_high": 0,
        "verdict_medium": 0,
        "verdict_low": 0,
        "verdict_empty": 0,
        "date_high": 0,
        "date_medium": 0,
        "date_low": 0,
        "date_empty": 0,
    }

    positions = list(df.index)
    if skip > 0:
        positions = positions[skip:]
    if limit is not None:
        positions = positions[:limit]

    pbar = tqdm(positions, desc="PDF-extractie", unit="record")
    for pos in pbar:
        row = df.iloc[pos]
        excel_row = pos + 2
        filename = row.get("bijlage.bestandsnaam", "") or row.get("naam", "")
        site_id = row.get("assetId.identificator", "")
        if not filename or filename == "":
            stats["skipped_no_name"] += 1
            continue

        pdf_path = find_pdf(bijlagen_dir, row.get("bijlage.bestandsnaam", ""), row.get("naam", ""), file_index)
        if pdf_path is None:
            stats["pdf_not_found"] += 1
            continue

        stats["pdf_found"] += 1

        text = extract_pdf_text(pdf_path)
        if not text or not text.strip():
            stats["skipped_corrupt"] += 1
            continue

        struct_info = detect_structure(text)
        structure = struct_info.name
        stats["structure_counts"][structure] = stats["structure_counts"].get(structure, 0) + 1

        verdict, vconf, vreason, has_besluit = _classify_verdict(text, structure)
        date_iso, dconf, dreason = extract_inspection_date(text, structure, filename)

        if vconf == "HIGH":
            stats["verdict_high"] += 1
        elif vconf == "MEDIUM":
            stats["verdict_medium"] += 1
        elif verdict:
            stats["verdict_low"] += 1
        else:
            stats["verdict_empty"] += 1

        if dconf == "HIGH":
            stats["date_high"] += 1
        elif dconf == "MEDIUM":
            stats["date_medium"] += 1
        elif date_iso:
            stats["date_low"] += 1
        else:
            stats["date_empty"] += 1

        stats["processed"] += 1

        if vconf in ("HIGH",) or dconf in ("HIGH",):
            tqdm.write(
                f"  Row {excel_row}: {structure} | verdict={verdict}({vconf}) | "
                f"date={date_iso}({dconf}) | file={pdf_path.name}"
            )

        results.append(ExtractionResult(
            excel_row=excel_row,
            site_id=site_id,
            filename=filename,
            pdf_path=str(pdf_path),
            structure=structure,
            verdict=verdict,
            verdict_confidence=vconf,
            verdict_reason=vreason,
            inspection_date=date_iso,
            date_confidence=dconf,
            date_reason=dreason,
            text_length=len(text),
            has_besluit=has_besluit,
        ))

    pbar.close()

    result_df = pd.DataFrame([asdict(r) for r in results])
    result_df = result_df.sort_values("excel_row")
    result_df.to_csv(output_path, index=False)
    result_df.to_excel(output_path.with_suffix(".xlsx"), index=False)

    return stats


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract keuringresultaat + keuringsdatum from PDFs for the Excel"
    )
    parser.add_argument("--excel", type=Path, default=EXCEL_PATH)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument("--limit", type=int, default=None)
    parser.add_argument("--skip", type=int, default=0)
    parser.add_argument("--bijlagen", type=Path, default=BIJLAGEN_DIR)
    args = parser.parse_args()

    output_path = args.output or args.excel.parent / f"{args.excel.stem}.extraction-results.csv"
    stats = process_excel(args.excel, output_path, args.bijlagen, args.limit, args.skip)

    print("\n=== Statistics ===")
    for k, v in stats.items():
        if k == "structure_counts":
            print(f"  {k}:")
            for sk, sv in sorted(v.items(), key=lambda x: -x[1]):
                print(f"    {sk}: {sv}")
        else:
            print(f"  {k}: {v}")
    print(f"\nOutput: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
