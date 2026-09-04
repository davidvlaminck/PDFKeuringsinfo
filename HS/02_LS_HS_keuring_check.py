"""
Check whether each PDF inspection report is a high-voltage (hoogspanning) or low-voltage (laagspanning) inspection.

Based on verify_pdf_contents_result_and_date.py, but simplified to only extract the inspection type.
"""
from __future__ import annotations

import argparse
import importlib
import logging
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Any

from openpyxl import load_workbook, Workbook
from tqdm import tqdm

logger = logging.getLogger(__name__)

HS_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = HS_DIR / "LS_pdfs_oplijsting.xlsx"
DEFAULT_PREVIOUS_OUTPUT = HS_DIR / "LS_pdfs_HS_keuring_check_full.xlsx"
DEFAULT_OUTPUT = HS_DIR / "LS_pdfs_HS_keuring_check_full_recheck.xlsx"
DEFAULT_PDF_BASE_DIR = Path("/home/davidlinux/Documenten/AWV/Keuringen")


@dataclass(frozen=True)
class InspectionTypeResult:
    is_high_voltage: bool | None
    found_text: str
    confidence: str


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


@lru_cache(maxsize=1)
def _get_pdf_reader():
    try:
        pdf_module = importlib.import_module("pypdf")
        return getattr(pdf_module, "PdfReader")
    except Exception as exc:
        raise RuntimeError("Dependency 'pypdf' ontbreekt. Run eerst: uv sync") from exc


def _extract_pdf_text(pdf_path: Path) -> str:
    PdfReader = _get_pdf_reader()
    reader = PdfReader(str(pdf_path))
    parts: list[str] = []
    for page in reader.pages:
        try:
            parts.append(page.extract_text() or "")
        except Exception:
            parts.append("")
    text = "\n".join(parts)

    if text.strip() and not _is_likely_garbage(text):
        if not parts[0].strip():
            ocr_text = _extract_pdf_text_ocr(pdf_path)
            if ocr_text.strip():
                return ocr_text
        return text

    return _extract_pdf_text_ocr(pdf_path)


def _is_likely_garbage(text: str, threshold: float = 0.15) -> bool:
    if not text or not text.strip():
        return True

    sample = text[:2000]
    non_printable = sum(1 for ch in sample if not ch.isprintable() and ch not in "\n\r\t")
    if len(sample) == 0:
        return True
    return (non_printable / len(sample)) >= threshold


def _extract_pdf_text_ocr(pdf_path: Path) -> str:
    try:
        import pymupdf
        from PIL import Image
        import pytesseract
        import io
    except ImportError:
        logger.debug("OCR dependencies ontbreken; sla OCR over voor %s", pdf_path)
        return ""

    try:
        doc = pymupdf.open(pdf_path)
        if len(doc) == 0:
            return ""
        page = doc[0]
        mat = pymupdf.Matrix(300 / 72, 300 / 72)
        pix = page.get_pixmap(matrix=mat)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        doc.close()
    except Exception as exc:
        logger.warning("Kon pagina 1 niet renderen voor OCR (%s): %s", pdf_path, exc)
        return ""

    try:
        ocr_text = pytesseract.image_to_string(img, lang="eng")
        return ocr_text or ""
    except Exception as exc:
        logger.warning("OCR mislukt voor %s: %s", pdf_path, exc)
        return ""


def _extract_xlsx_text(xlsx_path: Path) -> str:
    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    parts: list[str] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for cell_value in row:
                if cell_value is not None:
                    parts.append(str(cell_value))
    wb.close()
    return "\n".join(parts)


def classify_inspection_type(pdf_text: str) -> InspectionTypeResult:
    if not pdf_text or not pdf_text.strip():
        return InspectionTypeResult(None, "", "LOW")

    lines = pdf_text.splitlines()
    combined = " ".join(line.strip() for line in lines if line.strip())

    combined = re.sub(r"verslaag", "verslag", combined, flags=re.IGNORECASE)
    combined = re.sub(r"([A-Z]{2,})\s+([A-Z])\b", r"\1\2", combined)

    hs_high_conf_patterns = [
        r"verslag\s+van\s+controlebezoek\s*h\.?\s*(?:l\.?\s*)?s\.?\s*installatie",
        r"verslag\s+van\s+elektrische\s+hoogspanningsinstallatie",
        r"verslag\s*(?:[\(\[]?\d[\)\]]?\s*)?:?\s*elektrische\s+installaties?\s+op\s+hoogspanning",
        r"verslag\s+van\s+controle\s+van\s+(?:een\s+)?(?:elektrische\s+)?hoogspanningsinstallatie",
        r"verslag\s+van\s+periodieke\s+controle\s+van\s+een\s+(?:elektrische\s+)?hoogspanningsinstallatie",
        r"verslag\s+van\s+eerste\s+controle\s*\([^)]*\)\s+van\s+een\s+elektrische\s+hoogspanningsinstallatie",
        r"verslag\s+van\s+eerste\s+controle\s+van\s+een\s+elektrische\s+hoogspanningsinstallatie",
        r"controlebezoek\s+van\s+een\s+elektrische\s+installaties?\s+op\s+hoogspanning",
        r"verslag\s+van\s+controle\s+van\s+elektrische\s+hoogspanningsinstallatie",
        r"proces-verbaal\s+\d*\s*van\s+periodiek\s+controlebezoek\s+van\s+een\s+hoogspanningsinstallatie",
        r"verslag\s+van\s+gelijkvormigheidsonderzoek\s+van\s+(?:een\s+)?(?:elektrische\s+)?hoogspanningsinstallatie",
        r"verslag\s+\d*\s*van\s+(?:het\s+)?gelijkvormigheidsonderzoek\s+van\s+(?:een\s+)?(?:elektrische\s+)?hoogspanningsinstallatie",
        r"verslag\s+van\s+periodiek\s+controlebezoek\s+van\s+een\s+(?:elektrische\s+)?hoogspanningsinstallatie",
        r"verslag\s+van\s+controlebezoek\s+hoogspanningsinstallatie",
    ]

    ls_high_conf_patterns = [
        r"verslag\s+van\s+controle\s+van\s+een\s+(?:elektrische\s+)?laagspanningsinstallatie",
        r"verslag\s+van\s+periodieke\s+controle\s+van\s+een\s+(?:elektrische\s+)?laagspanningsinstallatie",
        r"verslag\s+van\s+eerste\s+controle\s*\([^)]*\)\s+van\s+een\s+elektrische\s+laagspanningsinstallatie",
        r"verslag\s+van\s+eerste\s+controle\s+van\s+een\s+elektrische\s+laagspanningsinstallatie",
        r"controlebezoek\s+van\s+een\s+elektrische\s+installatie\s+op\s+laagspanning",
        r"verslag\s+van\s+controle\s+van\s+elektrische\s+laagspanningsinstallatie",
        r"verslag\s+van\s+gelijkvormigheidsonderzoek\s+van\s+(?:een\s+)?(?:elektrische\s+)?laagspanningsinstallatie",
        r"verslag\s+\d*\s*van\s+(?:het\s+)?gelijkvormigheidsonderzoek\s+van\s+(?:een\s+)?(?:elektrische\s+)?laagspanningsinstallatie",
        r"ONDERZOEKSVERSLAG.*?Betreft\s*:?\s*elektrische\s+installatie\s+L\.S\.",
    ]

    hs_medium_conf_patterns = [
        r"elektrische\s+installaties?\s+op\s+hoogspanning",
        r"elektrische\s+hoogspanningsinstallatie",
        r"hoogspanningsinstallaties?",
        r"hs[- ]?cabine",
    ]

    ls_medium_conf_patterns = [
        r"elektrische\s+installatie\s+op\s+laagspanning\s+en\s+op\s+zeer\s+lage\s+spanning",
        r"elektrische\s+laagspanningsinstallatie",
        r"elektrische\s+installatie\s+op\s+laagspanning",
        r"elektrische\s+installatie\s+L\.S\.",
        r"(?=(?:.*voeding\s+afkomstig\s+van\s+vast\s+laagspanningsnet))elektrische\s+neon[\s-]?installatie",
        r"voeding\s+afkomstig\s+van\s+vast\s+laagspanningsnet",
        r"onderzoeksverslag.*L\.S\.",
        r"Betreft\s*:\s*elektrische\s+installatie\s+L\.S\.",
       r"laagspanningsinstallaties?",
    ]

    def find_matching_text(patterns: list[str], text: str) -> str | None:
        for pat in patterns:
            m = re.search(pat, text, flags=re.IGNORECASE)
            if m:
                return m.group(0)
        return None

    hs_text = find_matching_text(hs_high_conf_patterns, combined)
    ls_text = find_matching_text(ls_high_conf_patterns, combined)

    if hs_text and not ls_text:
        return InspectionTypeResult(True, hs_text, "HIGH")
    if ls_text and not hs_text:
        return InspectionTypeResult(False, ls_text, "HIGH")
    if hs_text and ls_text:
        return InspectionTypeResult(True, hs_text, "MEDIUM")

    hs_text = find_matching_text(hs_medium_conf_patterns, combined)
    ls_text = find_matching_text(ls_medium_conf_patterns, combined)

    if hs_text and not ls_text:
        return InspectionTypeResult(True, hs_text, "MEDIUM")
    if ls_text and not hs_text:
        has_neon = re.search(r"elektrische\s+neon[\s-]?installatie", combined, flags=re.IGNORECASE)
        has_voeding = re.search(r"voeding\s+afkomstig\s+van\s+vast\s+laagspanningsnet", combined, flags=re.IGNORECASE)
        if has_neon and has_voeding:
            return InspectionTypeResult(False, ls_text, "HIGH")
        return InspectionTypeResult(False, ls_text, "MEDIUM")
    if hs_text and ls_text:
        return InspectionTypeResult(True, hs_text, "MEDIUM")

    hs_lines = [line.strip() for line in lines if "hoogspanning" in line.lower()]
    ls_lines = [line.strip() for line in lines if "laagspanning" in line.lower()]

    hs_text = hs_lines[0] if hs_lines else ""
    ls_text = ls_lines[0] if ls_lines else ""

    if hs_text and not ls_text:
        return InspectionTypeResult(True, hs_text, "MEDIUM")
    if ls_text and not hs_text:
        return InspectionTypeResult(False, ls_text, "MEDIUM")
    if hs_text and ls_text:
        return InspectionTypeResult(True, hs_text, "LOW")

    return InspectionTypeResult(None, "", "LOW")


def classify_attachment_type(pdf_text: str) -> str:
    if not pdf_text:
        return "keuring"

    lowered = pdf_text.lower()
    risicoanalyse_patterns = [
        r"bijkomende opnames risicoanalyse distributiecabines",
        r"sythesedocument.*bijstand.*risicoanalyse",
        r"bijstand in het kader van.*risicoanalyse",
        r"risico\s*[- ]?index",
        r"bijlage\s*a\s*deze\s+bijlage\s+maakt\s+deel\s+uit\s+v\d+.*?document\s+met\s+ref",
    ]
    for pat in risicoanalyse_patterns:
        if re.search(pat, lowered, flags=re.DOTALL):
            return "risicoanalyse"
    return "keuring"


def load_input_rows(input_path: Path) -> list[dict[str, Any]]:
    wb = load_workbook(input_path, read_only=True, data_only=True)
    ws = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    rows: list[dict[str, Any]] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        record = {str(h).strip() if h is not None else f"col_{i}": v for i, (h, v) in enumerate(zip(headers, row))}
        rows.append(record)
    wb.close()
    return rows


def check_inspection_types(
    rows: list[dict[str, Any]],
    pdf_base_dir: Path,
    limit: int | None = None,
    skip: int = 0,
) -> tuple[list[dict[str, Any]], dict[str, int]]:
    results: list[dict[str, Any]] = []
    stats = {
        "processed": 0,
        "hs_keuring": 0,
        "ls_keuring": 0,
        "unknown": 0,
        "pdf_read_error": 0,
    }

    eligible = rows[skip:]
    if limit is not None:
        eligible = eligible[:limit]

    pbar = tqdm(total=len(eligible), desc="PDF type controle", unit="pdf")

    for record in eligible:
        pdf_path_str = _normalize_text(record.get("PDF pad") or record.get("pdf_path"))
        if not pdf_path_str:
            stats["unknown"] += 1
            results.append({
                "LS UUID": record.get("ls_uuid", record.get("LS UUID", "")),
                "LS Naam": record.get("ls_naam", record.get("LS Naam", "")),
                "Site ID": record.get("site_id", record.get("Site ID", "")),
                "Bestandsnaam": record.get("filename", record.get("Bestandsnaam", "")),
                "PDF pad": pdf_path_str,
                "HS_Keuring": "",
                "Gevonden_tekst": "",
                "Keuring_type": "",
                "Vertrouwen": "",
                "Opmerking": "Geen PDF-pad",
            })
            pbar.update(1)
            continue

        pdf_path = Path(pdf_path_str)
        if not pdf_path.exists():
            stats["unknown"] += 1
            results.append({
                "LS UUID": record.get("ls_uuid", record.get("LS UUID", "")),
                "LS Naam": record.get("ls_naam", record.get("LS Naam", "")),
                "Site ID": record.get("site_id", record.get("Site ID", "")),
                "Bestandsnaam": record.get("filename", record.get("Bestandsnaam", "")),
                "PDF pad": pdf_path_str,
                "HS_Keuring": "",
                "Gevonden_tekst": "",
                "Keuring_type": "",
                "Vertrouwen": "",
                "Opmerking": "PDF niet gevonden",
            })
            pbar.update(1)
            continue

        try:
            file_ext = pdf_path.suffix.lower()
            if file_ext == ".xlsx":
                text = _extract_xlsx_text(pdf_path)
            else:
                text = _extract_pdf_text(pdf_path)
        except Exception as exc:
            stats["pdf_read_error"] += 1
            stats["unknown"] += 1
            results.append({
                "LS UUID": record.get("ls_uuid", record.get("LS UUID", "")),
                "LS Naam": record.get("ls_naam", record.get("LS Naam", "")),
                "Site ID": record.get("site_id", record.get("Site ID", "")),
                "Bestandsnaam": record.get("filename", record.get("Bestandsnaam", "")),
                "PDF pad": pdf_path_str,
                "HS_Keuring": "",
                "Gevonden_tekst": "",
                "Keuring_type": "",
                "Vertrouwen": "",
                "Opmerking": f"PDF lezen mislukt: {exc}",
            })
            pbar.update(1)
            continue

        classification = classify_inspection_type(text)
        stats["processed"] += 1

        if classification.is_high_voltage is True:
            stats["hs_keuring"] += 1
            keuring_type = "HOOGSPANNING"
            hs_keuring = "TRUE"
        elif classification.is_high_voltage is False:
            stats["ls_keuring"] += 1
            keuring_type = "LAAGSPANNING"
            hs_keuring = ""
        else:
            stats["unknown"] += 1
            keuring_type = ""
            hs_keuring = ""

        results.append({
            "LS UUID": record.get("ls_uuid", record.get("LS UUID", "")),
            "LS Naam": record.get("ls_naam", record.get("LS Naam", "")),
            "Site ID": record.get("site_id", record.get("Site ID", "")),
            "Bestandsnaam": record.get("filename", record.get("Bestandsnaam", "")),
            "PDF pad": pdf_path_str,
            "HS_Keuring": hs_keuring,
            "Gevonden_tekst": classification.found_text,
            "Keuring_type": keuring_type,
            "Vertrouwen": classification.confidence,
            "Type bijlage": classify_attachment_type(text),
            "Opmerking": "",
        })

        pbar.set_postfix(
            {
                "processed": stats["processed"],
                "hs": stats["hs_keuring"],
                "ls": stats["ls_keuring"],
                "unknown": stats["unknown"],
                "errors": stats["pdf_read_error"],
            },
            refresh=False,
        )
        pbar.update(1)

    pbar.close()
    logger.info("Check stats: %s", stats)
    return results, stats


def write_output(results: list[dict[str, Any]], output_path: Path) -> None:
    if not results:
        logger.warning("Geen resultaten om te schrijven.")
        return

    results_sorted = sorted(
        results,
        key=lambda r: (
            r.get("LS Naam", ""),
            r.get("Site ID", ""),
            r.get("Bestandsnaam", ""),
        ),
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "HS_LS_Keuring_check"

    headers = [
        "LS UUID",
        "LS Naam",
        "Site ID",
        "Bestandsnaam",
        "PDF pad",
        "HS_Keuring",
        "Gevonden_tekst",
        "Keuring_type",
        "Vertrouwen",
        "Type bijlage",
        "Opmerking",
    ]

    ws.append(headers)

    for record in results_sorted:
        ws.append(
            [
                record.get("LS UUID", ""),
                record.get("LS Naam", ""),
                record.get("Site ID", ""),
                record.get("Bestandsnaam", ""),
                record.get("PDF pad", ""),
                record.get("HS_Keuring", ""),
                record.get("Gevonden_tekst", ""),
                record.get("Keuring_type", ""),
                record.get("Vertrouwen", ""),
                record.get("Type bijlage", ""),
                record.get("Opmerking", ""),
            ]
        )

    # Autofit kolommen, behalve kolom E (PDF pad, index 4)
    for col_idx, header in enumerate(headers, 1):
        if col_idx == 5:
            continue
        max_length = len(str(header))
        for row_idx in range(2, len(results_sorted) + 2):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is not None:
                max_length = max(max_length, len(str(cell_value)))
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = adjusted_width

    try:
        wb.save(output_path)
    except Exception as exc:
        raise RuntimeError(f"Excel export mislukt ({output_path}): {exc}") from exc


def check_ls_hs_keuring(
    *,
    input_path: Path,
    output_path: Path,
    pdf_base_dir: Path,
    previous_output: Path | None = None,
    limit: int | None = None,
    skip: int = 0,
) -> dict[str, int]:
    logger.info("Lees input uit: %s", input_path)
    rows = load_input_rows(input_path)
    logger.info("Aantal rijen in input: %d", len(rows))

    if previous_output and previous_output.exists():
        logger.info("Lees vorige output uit: %s", previous_output)
        prev_wb = load_workbook(previous_output, read_only=True, data_only=True)
        prev_ws = prev_wb.active
        prev_headers = [cell.value for cell in next(prev_ws.iter_rows(min_row=1, max_row=1))]
        pdf_path_col = None
        for i, h in enumerate(prev_headers):
            if str(h).strip() == "PDF pad":
                pdf_path_col = i
                break
        confidence_col = None
        for i, h in enumerate(prev_headers):
            if str(h).strip() == "Vertrouwen":
                confidence_col = i
                break

        if pdf_path_col is not None and confidence_col is not None:
            high_confidence_paths = set()
            type_bijlage_col = None
            for i, h in enumerate(prev_headers):
                if str(h).strip() == "Type bijlage":
                    type_bijlage_col = i
                    break
            keuring_paths = set()
            for row in prev_ws.iter_rows(min_row=2, values_only=True):
                if len(row) > max(pdf_path_col, confidence_col):
                    pdf_path = row[pdf_path_col]
                    confidence = row[confidence_col]
                    if confidence == "HIGH" and pdf_path:
                        high_confidence_paths.add(str(pdf_path).strip())
                    if type_bijlage_col is not None and len(row) > type_bijlage_col:
                        if row[type_bijlage_col] == "keuring" and pdf_path:
                            keuring_paths.add(str(pdf_path).strip())
            prev_wb.close()
            logger.info("Aantal reeds HIGH confidence gemarkeerd: %d", len(high_confidence_paths))
            logger.info("Aantal type 'keuring' gemarkeerd: %d", len(keuring_paths))
            rows = [
                r for r in rows
                if str(r.get("PDF pad", "")).strip() not in high_confidence_paths
                and str(r.get("PDF pad", "")).strip() in keuring_paths
            ]
            logger.info("Aantal te controleren na filter: %d", len(rows))

    results, stats = check_inspection_types(rows, pdf_base_dir, limit=limit, skip=skip)

    logger.info("Schrijf output naar: %s", output_path)
    write_output(results, output_path)

    stats["output_rows"] = len(results)
    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Controleer per PDF of het een hoogspannings- of laagspanningskeuring is."
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_INPUT,
        help=f"Pad naar input Excel (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Pad naar output Excel (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--pdf-base-dir",
        type=Path,
        default=DEFAULT_PDF_BASE_DIR,
        help=f"Basismap met site directories en PDFs (default: {DEFAULT_PDF_BASE_DIR})",
    )
    parser.add_argument(
        "--previous-output",
        type=Path,
        default=DEFAULT_PREVIOUS_OUTPUT,
        help=f"Pad naar vorige output Excel (default: {DEFAULT_PREVIOUS_OUTPUT})",
    )
    parser.add_argument(
        "--skip",
        type=int,
        default=0,
        help="Optioneel: sla de eerste N in-scope records over.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Optioneel: verwerk enkel de eerste N in-scope records.",
    )
    return parser.parse_args()


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    args = parse_args()

    if not args.input.exists():
        logger.error("Inputbestand bestaat niet: %s", args.input)
        return 1

    try:
        stats = check_ls_hs_keuring(
            input_path=args.input,
            output_path=args.output,
            pdf_base_dir=args.pdf_base_dir,
            previous_output=args.previous_output,
            limit=args.limit,
            skip=args.skip,
        )
    except Exception as exc:
        logger.error("HS/LS check mislukt: %s", exc)
        return 3

    logger.info("Klaar. Output: %s", args.output)
    logger.info("Stats: %s", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
