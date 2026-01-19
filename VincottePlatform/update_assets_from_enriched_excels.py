"""Update assets based on enriched Vincotte Excel files.

For each row in an *enriched* Excel:
- We consider using the API if:
  - visit_date (Excel) is more recent than Arango keuringsdatum, OR
  - Arango keuringsdatum is empty but Arango UUID is filled.

For now, this script only implements the *read-only* step:
- Find candidate rows
- Fetch and print a small sample of 'eigenschappen' for the asset referenced by 'Arango UUID'

This script is intended to be runnable from PyCharm.
"""

from __future__ import annotations

import argparse
import logging
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from API.eminfra.EMInfraClient import EMInfraClient
from API.Enums import AuthType, Environment

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class CandidateRow:
    excel_path: Path
    row_index: int
    asset_uuid: str
    arango_keuringsdatum: date | None
    excel_visit_date: date | None


def _parse_date(value: Any) -> date | None:
    """Parse dates from excel cells.

    Supports:
    - datetime/date objects
    - strings like '2017-08-23' or '23/08/2017'
    """

    if value is None:
        return None

    if isinstance(value, date) and not isinstance(value, datetime):
        return value

    if isinstance(value, datetime):
        return value.date()

    s = str(value).strip()
    if not s:
        return None

    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    for fmt in ("%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    return None


def _find_col_by_header(ws, header: str) -> int | None:
    header_row = 1
    wanted = header.strip().lower()
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        if str(v).strip().lower() == wanted:
            return c
    return None


def _find_visit_date_col(ws) -> int:
    """Find the visit date column.

    Priority:
    1) Exact header 'visit_date' (case-insensitive)
    2) Header containing both 'visit' and 'date'
    3) Fallback to column G (7)

    Returns 1-based column index.
    """

    header_row = 1
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        if str(v).strip().lower() == "visit_date":
            return c

    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        s = str(v).strip().lower()
        if "visit" in s and "date" in s:
            return c

    return 7


def iter_candidates_from_excel(excel_path: Path, *, debug: bool = False) -> list[CandidateRow]:
    wb = load_workbook(excel_path)
    ws = wb.active

    uuid_col = _find_col_by_header(ws, "Arango UUID")
    arango_date_col = _find_col_by_header(ws, "Arango keuringsdatum")
    visit_date_col = _find_visit_date_col(ws)

    if uuid_col is None:
        raise ValueError(f"Missing required column header 'Arango UUID' in {excel_path.name}")

    if debug:
        logger.info(
            "Using columns | Arango UUID=%s, Arango keuringsdatum=%s, visit_date=%s (header=%r)",
            uuid_col,
            arango_date_col,
            visit_date_col,
            ws.cell(row=1, column=visit_date_col).value,
        )

    candidates: list[CandidateRow] = []

    for r in range(2, ws.max_row + 1):
        asset_uuid = ws.cell(row=r, column=uuid_col).value
        if asset_uuid is None or str(asset_uuid).strip() == "":
            continue
        asset_uuid = str(asset_uuid).strip()

        arango_dt = _parse_date(ws.cell(row=r, column=arango_date_col).value) if arango_date_col else None
        excel_dt = _parse_date(ws.cell(row=r, column=visit_date_col).value)

        # Candidate rules:
        # - Excel date newer than Arango date
        # - OR Arango date empty but UUID filled
        if arango_dt is None:
            candidates.append(
                CandidateRow(
                    excel_path=excel_path,
                    row_index=r,
                    asset_uuid=asset_uuid,
                    arango_keuringsdatum=None,
                    excel_visit_date=excel_dt,
                )
            )
            continue

        if excel_dt is None:
            continue

        if excel_dt > arango_dt:
            candidates.append(
                CandidateRow(
                    excel_path=excel_path,
                    row_index=r,
                    asset_uuid=asset_uuid,
                    arango_keuringsdatum=arango_dt,
                    excel_visit_date=excel_dt,
                )
            )

    return candidates


def fetch_all_eigenschappen(eminfra: EMInfraClient, asset_uuid: str):
    """Fetch all eigenschappen for an asset (read-only)."""

    return eminfra.eigenschap_service.get_eigenschappen(asset_uuid)


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=(
            "Scan enriched Vincotte Excel files and for candidate rows fetch all eigenschappen of the related asset via EMInfra API."
        )
    )

    p.add_argument(
        "--settings",
        type=Path,
        default=Path("/home/davidlinux/Documents/AWV/resources/settings_SyncOTLDataToLegacy.json"),
        help="Path to settings JSON used by the API requester (default: %(default)s)",
    )
    p.add_argument(
        "--env",
        default="prd",
        choices=["prd", "dev", "tei", "aim"],
        help="Environment (default: %(default)s)",
    )
    p.add_argument(
        "--excel",
        type=Path,
        default=Path(__file__).resolve().parent / "reports.enriched.xlsx",
        help="Path to an enriched Excel file to process (default: %(default)s)",
    )
    p.add_argument(
        "--limit",
        type=int,
        default=1,
        help="Max number of candidate rows to fetch eigenschappen for (default: %(default)s)",
    )
    p.add_argument(
        "--debug-cols",
        action="store_true",
        help="Log which columns are used for Arango UUID / Arango keuringsdatum / visit_date.",
    )

    return p


def main() -> int:
    logging.basicConfig(
        format="%(asctime)s %(levelname)-8s %(message)s",
        level=logging.INFO,
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    args = _build_parser().parse_args()

    env_map = {
        "prd": Environment.PRD,
        "dev": Environment.DEV,
        "tei": Environment.TEI,
        "aim": Environment.AIM,
    }

    excel_path: Path = args.excel
    if not excel_path.exists():
        logger.error("Excel not found: %s", excel_path)
        return 2

    candidates = iter_candidates_from_excel(excel_path, debug=bool(args.debug_cols))
    logger.info("Found %d candidate row(s) in %s", len(candidates), excel_path.name)

    if not candidates:
        return 0

    eminfra = EMInfraClient(env=env_map[args.env], auth_type=AuthType.JWT, settings_path=args.settings)

    for c in candidates[: max(0, int(args.limit))]:
        logger.info(
            "Row %d | uuid=%s | Arango=%s | Excel=%s => fetching eigenschappen",
            c.row_index,
            c.asset_uuid,
            c.arango_keuringsdatum.isoformat() if c.arango_keuringsdatum else None,
            c.excel_visit_date.isoformat() if c.excel_visit_date else None,
        )

        try:
            eigs = fetch_all_eigenschappen(eminfra, c.asset_uuid)
            logger.info("Fetched %d eigenschappen", len(eigs))
            for e in eigs[:5]:
                try:
                    logger.info("  - %s = %s", e.eigenschap.naam, e.typedValue)
                except Exception:
                    logger.info("  - %r", e)
        except Exception:
            logger.exception("Failed fetching eigenschappen for %s", c.asset_uuid)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
