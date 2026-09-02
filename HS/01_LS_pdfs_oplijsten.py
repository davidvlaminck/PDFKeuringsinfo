"""
List all PDF inspection reports for Laagspanningsborden (LS borden) assets.

This script reads the HS Excel with LS bord UUIDs and the enriched Vincotte Excel,
then produces an output Excel listing every PDF associated with each LS bord.
"""
from __future__ import annotations

import argparse
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook, Workbook

logger = logging.getLogger(__name__)

HS_DIR = Path(__file__).resolve().parent
DEFAULT_HS_INPUT = HS_DIR / "LSBord_bevestiging_HSCabine.xlsx"
DEFAULT_ENRICHED_INPUT = Path("/home/davidlinux/Documenten/AWV/Keuringen/MyVinçotte Min.Vl.Gemeenschap Reports list.enriched.xlsx")
DEFAULT_OUTPUT = HS_DIR / "LS_pdfs_oplijsting.xlsx"
DEFAULT_PDF_BASE_DIR = Path("/home/davidlinux/Documenten/AWV/Keuringen")


@dataclass(frozen=True)
class LsAsset:
    uuid: str
    naam: str
    naampad: str
    commentaar: str


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def load_ls_assets(hs_path: Path) -> list[LsAsset]:
    wb = load_workbook(hs_path, read_only=True, data_only=True)
    ws = wb["query-result"]
    assets: list[LsAsset] = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        uuid = _normalize_text(row[3] if len(row) > 3 else None)
        naam = _normalize_text(row[4] if len(row) > 4 else None)
        naampad = _normalize_text(row[5] if len(row) > 5 else None)
        commentaar = _normalize_text(row[6] if len(row) > 6 else None)
        if uuid:
            assets.append(LsAsset(uuid=uuid, naam=naam, naampad=naampad, commentaar=commentaar))
    wb.close()
    return assets


def load_enriched_rows(enriched_path: Path) -> list[dict[str, Any]]:
    wb = load_workbook(enriched_path, read_only=True, data_only=True)
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


def resolve_pdf_path(base_dir: Path, site_id: str, filename: str) -> Path | None:
    if not site_id or not filename:
        return None
    site_dir = base_dir / str(site_id)
    if not site_dir.exists() or not site_dir.is_dir():
        return None
    candidate = site_dir / str(filename)
    if candidate.exists() and candidate.is_file():
        return candidate
    if not str(filename).lower().endswith(".pdf"):
        candidate = site_dir / f"{filename}.pdf"
        if candidate.exists() and candidate.is_file():
            return candidate
    lower_map = {p.name.lower(): p for p in site_dir.iterdir() if p.is_file()}
    found = lower_map.get(str(filename).lower())
    if found is not None:
        return found
    if not str(filename).lower().endswith(".pdf"):
        found = lower_map.get(f"{filename.lower()}.pdf")
        if found is not None:
            return found
    return None


def build_pdf_list(
    ls_assets: list[LsAsset],
    enriched_rows: list[dict[str, Any]],
    pdf_base_dir: Path,
) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    stats = {
        "ls_assets": len(ls_assets),
        "enriched_rows": len(enriched_rows),
        "matched_assets": 0,
        "matched_rows": 0,
        "pdf_found": 0,
        "pdf_missing": 0,
    }

    uuid_to_asset = {a.uuid: a for a in ls_assets}

    for row in enriched_rows:
        uuid = _normalize_text(row.get("Arango UUID"))
        asset = uuid_to_asset.get(uuid)
        if asset is None:
            continue

        site_id = _normalize_text(row.get("Site ID"))
        filename = _normalize_text(row.get("FILENAME"))
        pdf_path = resolve_pdf_path(pdf_base_dir, site_id, filename)

        results.append(
            {
                "ls_uuid": asset.uuid,
                "ls_naam": asset.naam,
                "ls_naampad": asset.naampad,
                "ls_commentaar": asset.commentaar,
                "site_id": site_id,
                "filename": filename,
                "pdf_path": str(pdf_path) if pdf_path else "",
                "pdf_exists": "JA" if pdf_path else "NEEN",
                "visit_date": row.get("Visit Date"),
                "label": row.get("Label"),
                "report_nr": row.get("Report nr"),
                "visiting_place": row.get("Visiting place"),
                "arango_naam": row.get("Arango Naam"),
                "arango_naampad": row.get("Arango Naampad"),
                "uitkomst_enrich": row.get("Uitkomst enrich script"),
            }
        )
        stats["matched_rows"] += 1
        if pdf_path:
            stats["pdf_found"] += 1
        else:
            stats["pdf_missing"] += 1

    stats["matched_assets"] = len({r["ls_uuid"] for r in results})
    logger.info("Build PDF list stats: %s", stats)
    return results


def write_output(results: list[dict[str, Any]], output_path: Path) -> None:
    if not results:
        logger.warning("Geen resultaten om te schrijven.")
        return

    results_sorted = sorted(
        results,
        key=lambda r: (
            r.get("ls_naam", ""),
            r.get("site_id", ""),
            r.get("visit_date", ""),
            r.get("filename", ""),
        ),
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "LS PDFs"

    headers = [
        "LS UUID",
        "LS Naam",
        "LS Naampad",
        "LS Commentaar",
        "Site ID",
        "Bestandsnaam",
        "PDF pad",
        "PDF bestaat",
        "Bezoekdatum",
        "Label",
        "Rapport nr",
        "Bezoekplaats",
        "Arango Naam",
        "Arango Naampad",
        "Uitkomst enrich",
    ]

    ws.append(headers)

    for record in results_sorted:
        ws.append(
            [
                record.get("ls_uuid", ""),
                record.get("ls_naam", ""),
                record.get("ls_naampad", ""),
                record.get("ls_commentaar", ""),
                record.get("site_id", ""),
                record.get("filename", ""),
                record.get("pdf_path", ""),
                record.get("pdf_exists", ""),
                record.get("visit_date", ""),
                record.get("label", ""),
                record.get("report_nr", ""),
                record.get("visiting_place", ""),
                record.get("arango_naam", ""),
                record.get("arango_naampad", ""),
                record.get("uitkomst_enrich", ""),
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


def list_ls_pdfs(
    *,
    hs_input: Path,
    enriched_input: Path,
    output: Path,
    pdf_base_dir: Path,
) -> dict[str, int]:
    logger.info("Lees LS assets uit: %s", hs_input)
    ls_assets = load_ls_assets(hs_input)
    logger.info("Aantal LS assets gevonden: %d", len(ls_assets))

    logger.info("Lee enriched Excel uit: %s", enriched_input)
    enriched_rows = load_enriched_rows(enriched_input)
    logger.info("Aantal rijen in enriched Excel: %d", len(enriched_rows))

    results = build_pdf_list(ls_assets, enriched_rows, pdf_base_dir)

    logger.info("Schrijf output naar: %s", output)
    write_output(results, output)

    return {
        "ls_assets": len(ls_assets),
        "enriched_rows": len(enriched_rows),
        "output_rows": len(results),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Lis PDF-keuringrapporten voor Laagspanningsborden (LS borden)."
    )
    parser.add_argument(
        "--hs-input",
        type=Path,
        default=DEFAULT_HS_INPUT,
        help=f"Pad naar HS Excel met LS bord UUIDs (default: {DEFAULT_HS_INPUT})",
    )
    parser.add_argument(
        "--enriched-input",
        type=Path,
        default=DEFAULT_ENRICHED_INPUT,
        help=f"Pad naar enriched Vincotte Excel (default: {DEFAULT_ENRICHED_INPUT})",
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
    return parser.parse_args()


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    args = parse_args()

    if not args.hs_input.exists():
        logger.error("HS inputbestand bestaat niet: %s", args.hs_input)
        return 1
    if not args.enriched_input.exists():
        logger.error("Enriched inputbestand bestaat niet: %s", args.enriched_input)
        return 2

    try:
        stats = list_ls_pdfs(
            hs_input=args.hs_input,
            enriched_input=args.enriched_input,
            output=args.output,
            pdf_base_dir=args.pdf_base_dir,
        )
    except Exception as exc:
        logger.error("Oplijsten mislukt: %s", exc)
        return 3

    logger.info("Klaar. Output: %s", args.output)
    logger.info("Stats: %s", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
