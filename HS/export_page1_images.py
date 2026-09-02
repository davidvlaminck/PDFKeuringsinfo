"""
Export page 1 of PDFs without extractable text to images for manual inspection.
"""
from __future__ import annotations

import argparse
import logging
from pathlib import Path

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from tqdm import tqdm

logger = logging.getLogger(__name__)

HS_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = HS_DIR / "LS_pdfs_oplijsting.xlsx"
DEFAULT_OUTPUT = HS_DIR / "LS_pdfs_no_text_page1_images.xlsx"
DEFAULT_PDF_BASE_DIR = Path("/home/davidlinux/Documenten/AWV/Keuringen")
DEFAULT_IMAGE_DIR = HS_DIR / "pdf_page1_images"


def _normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


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


def render_page1_to_image(pdf_path: Path, output_path: Path, dpi: int = 200) -> bool:
    try:
        import fitz
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            return False
        page = doc[0]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        pix.save(str(output_path))
        doc.close()
        return output_path.exists() and output_path.stat().st_size > 0
    except Exception as exc:
        logger.warning("Kon pagina 1 niet renderen voor %s: %s", pdf_path, exc)
        return False


def export_no_text_pdfs(
    rows: list[dict[str, Any]],
    pdf_base_dir: Path,
    image_dir: Path,
) -> tuple[list[dict[str, Any]], dict[str, int]]:
    results: list[dict[str, Any]] = []
    stats = {
        "processed": 0,
        "has_text": 0,
        "no_text": 0,
        "rendered": 0,
        "render_failed": 0,
        "pdf_missing": 0,
    }

    image_dir.mkdir(parents=True, exist_ok=True)

    pbar = tqdm(total=len(rows), desc="PDF pagina 1 export", unit="pdf")

    for record in rows:
        pdf_path_str = _normalize_text(record.get("PDF pad") or record.get("pdf_path"))
        if not pdf_path_str:
            stats["pdf_missing"] += 1
            pbar.update(1)
            continue

        pdf_path = Path(pdf_path_str)
        if not pdf_path.exists():
            stats["pdf_missing"] += 1
            pbar.update(1)
            continue

        try:
            import fitz
            doc = fitz.open(pdf_path)
            page1_text = ""
            if len(doc) > 0:
                page1_text = doc[0].get_text() or ""
            doc.close()
        except Exception:
            page1_text = ""

        stats["processed"] += 1
        has_text = bool(page1_text.strip())
        stats["has_text" if has_text else "no_text"] += 1

        image_filename = ""
        image_full_path = ""
        if not has_text:
            safe_name = pdf_path.stem.replace("/", "_").replace("\\", "_")
            image_filename = f"{safe_name}_page1.png"
            image_full_path = image_dir / image_filename
            if render_page1_to_image(pdf_path, image_full_path):
                stats["rendered"] += 1
            else:
                stats["render_failed"] += 1

        results.append({
            "LS UUID": record.get("ls_uuid", record.get("LS UUID", "")),
            "LS Naam": record.get("ls_naam", record.get("LS Naam", "")),
            "Site ID": record.get("site_id", record.get("Site ID", "")),
            "Bestandsnaam": record.get("filename", record.get("Bestandsnaam", "")),
            "PDF pad": pdf_path_str,
            "Pagina1_tekst": page1_text.strip()[:500] if page1_text else "",
            "Heeft_tekst": "JA" if has_text else "NEE",
            "Afbeelding_bestand": str(image_full_path) if image_full_path else "",
            "Afbeelding_bestandsnaam": image_filename,
        })

        pbar.set_postfix(
            {
                "processed": stats["processed"],
                "tekst": stats["has_text"],
                "geen_tekst": stats["no_text"],
                "gerend": stats["rendered"],
            },
            refresh=False,
        )
        pbar.update(1)

    pbar.close()
    logger.info("Export stats: %s", stats)
    return results, stats


def write_output(results: list[dict[str, Any]], output_path: Path) -> None:
    if not results:
        logger.warning("Geen resultaten om te schrijven.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "PDFs zonder tekst"

    headers = [
        "LS UUID",
        "LS Naam",
        "Site ID",
        "Bestandsnaam",
        "PDF pad",
        "Pagina1_tekst",
        "Heeft_tekst",
        "Afbeelding_bestand",
        "Afbeelding_bestandsnaam",
    ]

    ws.append(headers)

    for record in results:
        ws.append(
            [
                record.get("LS UUID", ""),
                record.get("LS Naam", ""),
                record.get("Site ID", ""),
                record.get("Bestandsnaam", ""),
                record.get("PDF pad", ""),
                record.get("Pagina1_tekst", ""),
                record.get("Heeft_tekst", ""),
                record.get("Afbeelding_bestand", ""),
                record.get("Afbeelding_bestandsnaam", ""),
            ]
        )

    try:
        wb.save(output_path)
    except Exception as exc:
        raise RuntimeError(f"Excel export mislukt ({output_path}): {exc}") from exc


def export_no_text(
    *,
    input_path: Path,
    output_path: Path,
    pdf_base_dir: Path,
    image_dir: Path,
    limit: int | None = None,
    skip: int = 0,
) -> dict[str, int]:
    logger.info("Lees input uit: %s", input_path)
    rows = load_input_rows(input_path)
    logger.info("Aantal rijen in input: %d", len(rows))

    eligible = rows[skip:]
    if limit is not None:
        eligible = eligible[:limit]

    results, stats = export_no_text_pdfs(eligible, pdf_base_dir, image_dir)

    logger.info("Schrijf output naar: %s", output_path)
    write_output(results, output_path)

    stats["output_rows"] = len(results)
    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Exporteer pagina 1 van PDFs zonder tekst naar afbeeldingen."
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
        "--image-dir",
        type=Path,
        default=DEFAULT_IMAGE_DIR,
        help=f"Map voor opgeslagen afbeeldingen (default: {DEFAULT_IMAGE_DIR})",
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
        stats = export_no_text(
            input_path=args.input,
            output_path=args.output,
            pdf_base_dir=args.pdf_base_dir,
            image_dir=args.image_dir,
            limit=args.limit,
            skip=args.skip,
        )
    except Exception as exc:
        logger.error("Export mislukt: %s", exc)
        return 3

    logger.info("Klaar. Output: %s", args.output)
    logger.info("Stats: %s", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
