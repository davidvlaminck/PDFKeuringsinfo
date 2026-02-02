"""Update assets based on enriched Vincotte Excel files.

For each row in an *enriched* Excel:
- We consider using the API if:
  - visit_date (Excel) is more recent than Arango keuringsdatum, OR
  - Arango keuringsdatum is empty but Arango UUID is filled.

⚠️ Safety: this script defaults to DRY-RUN (no updates). It can fetch data from PRD,
but will not PATCH anything unless you explicitly pass --apply.

For now:
- Find candidate rows
- Fetch all eigenschapwaarden for the asset referenced by 'Arango UUID'
- Discover the Eigenschap UUIDs for the fields we want to update
- Prepare update payloads for those two fields:
  - "datum laatste keuring"  -> Excel visit date (column G)
  - "resultaat keuring"     -> based on Excel code in column B

This script is intended to be runnable from PyCharm.
"""

from __future__ import annotations

import argparse
import logging
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

# Make sure in-repo packages (API/, utils/) are importable when running this file directly.
import sys
REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

from API.eminfra.EMInfraClient import EMInfraClient
from API.Enums import AuthType, Environment
from API.eminfra.EMInfraDomain import Eigenschap, EigenschapValueUpdateDTO

logger = logging.getLogger(__name__)

# Toggle default dry-run behavior. Set to True to default to dry-run (no PATCH).
# You can change this in the `if __name__ == "__main__"` block below.
DEFAULT_DRY_RUN: bool = True

RESULT_CODE_TO_TEXT: dict[int, str] = {
    1: "conform",
    3: "niet van toepassing (buiten gebruik)",
    4: "conform met opmerkingen",
    5: "niet-conform met inbreuken",
}


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
    - Excel serial day numbers (e.g. 45291) via openpyxl conversion
    """

    if value is None:
        return None

    # Excel serial date numbers sometimes come through as int/float.
    if isinstance(value, (int, float)):
        try:
            # from_excel returns datetime for dates
            dt = from_excel(value)
            if isinstance(dt, datetime):
                return dt.date()
            if isinstance(dt, date):
                return dt
        except Exception:
            # If it's just a small number (like 16) that's not a date serial,
            # treat it as unparsable.
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
    4) Fallback to column F (6)

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

    # historical fallbacks
    return 7 if ws.max_column >= 7 else 6


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
    """Fetch all eigenschapwaarden for an asset (read-only).

    We prefer the LSDeel 'Eigenschappen' kenmerkType uuid to avoid picking the wrong kenmerk.
    """

    # From user API sample: Eigenschappen - lgc:installatie#LSDeel
    LSDEEL_KENMERK_UUID = "106add0d-c86d-4c4e-83c7-7dc8c2baab0f"
    try:
        eigs = eminfra.eigenschap_service.get_eigenschappen(asset_uuid, kenmerk_uuid=LSDEEL_KENMERK_UUID)
        if eigs:
            return eigs
    except Exception:
        # fall back to autodetect
        pass

    return eminfra.eigenschap_service.get_eigenschappen(asset_uuid)


def _excel_result_code_to_text(value: Any) -> str | None:
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    try:
        code = int(float(s))
    except Exception:
        return None
    return RESULT_CODE_TO_TEXT.get(code)


def _should_skip_due_to_result(result_text: str | None) -> bool:
    # When the excel indicates 'buiten gebruik', we should ignore the record.
    return (result_text or "").strip().lower() == "niet van toepassing (buiten gebruik)"


def _find_eigenschap_by_name(eigenschappen, name: str) -> Any | None:
    """Find an EigenschapValueDTO by eigenschap.naam (case-insensitive)."""
    wanted = name.strip().lower()
    for ev in eigenschappen:
        try:
            if ev.eigenschap and str(ev.eigenschap.naam).strip().lower() == wanted:
                return ev
        except Exception:
            continue
    return None


def _build_update_dto(existing_ev, typed_value: dict) -> EigenschapValueUpdateDTO:
    """Build an update DTO that reuses the same Eigenschap (uuid stays fixed)."""
    eig: Eigenschap = existing_ev.eigenschap
    return EigenschapValueUpdateDTO(typedValue=typed_value, eigenschap=eig)


LSDEEL_EIGENSCHAP_DATUM_UUID = "7c34fa88-4740-409a-8316-88e250f49548"
LSDEEL_EIGENSCHAP_RESULTAAT_UUID = "fef12239-764e-4068-9f76-73f0d6bb6446"


def _build_update_dto_by_eigenschap_uuid(eigenschap_uuid: str, typed_value: dict) -> EigenschapValueUpdateDTO:
    """Build an update DTO when the asset has no existing eigenschapwaarden yet.

    For LSDeel we assume the Eigenschap UUIDs are fixed.
    The API expects Eigenschap serialized; our client uses the Eigenschap dataclass,
    which requires some mandatory fields. We provide minimal placeholders.
    """

    # Minimal Link list to satisfy dataclass requirements.
    links = []

    # Provide minimal schema-ish fields. The API typically ignores most of these
    # besides uuid, but they are required for our dataclass construction.
    if eigenschap_uuid == LSDEEL_EIGENSCHAP_DATUM_UUID:
        uri = "https://ins.data.wegenenverkeer.be/ns/attribuut#EMObject.datumLaatsteKeuring"
        naam = "datum laatste keuring"
        label = naam
        categorie = "INSPECTIE"
        definitie = ""
        type_dict = {"_type": "date"}
    elif eigenschap_uuid == LSDEEL_EIGENSCHAP_RESULTAAT_UUID:
        uri = "https://ins.data.wegenenverkeer.be/ns/attribuut#EMObject.resultaatKeuring"
        naam = "resultaat keuring"
        label = naam
        categorie = "INSPECTIE"
        definitie = ""
        # In the schema this is a selection; the API accepts text typedValue with one of the allowed values.
        type_dict = {"_type": "selection"}
    else:
        uri = ""
        naam = ""
        label = ""
        categorie = ""
        definitie = ""
        type_dict = {}

    eig = Eigenschap(
        uuid=eigenschap_uuid,
        createdOn="",
        modifiedOn="",
        uri=uri,
        label=label,
        naam=naam,
        alleenLezen=False,
        actief=True,
        definitie=definitie,
        categorie=categorie,
        type=type_dict,
        links=links,
    )
    return EigenschapValueUpdateDTO(typedValue=typed_value, eigenschap=eig)


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description=(
            "From an enriched Vincotte Excel, process candidate rows and prepare/apply updates for 'datum laatste keuring' and "
            "'resultaat keuring' via EMInfra API (default DRY-RUN)."
        )
    )

    p.add_argument(
        "--settings",
        type=Path,
        default=Path("/home/davidlinux/Documenten/AWV/resources/settings_SyncOTLDataToLegacy.json"),
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
        default=None,
        help="Path to an enriched Excel file to process (default: process all *.enriched.xlsx in VincottePlatform)",
    )
    p.add_argument(
        "--all",
        action="store_true",
        help="Process all *.xlsx files that contain '.enriched' in the filename in VincottePlatform/",
    )
    p.add_argument(
        "--debug-cols",
        action="store_true",
        help="Log which columns are used for Arango UUID / Arango keuringsdatum / visit_date.",
    )
    p.add_argument(
        "--row",
        type=int,
        default=None,
        help="Excel row number to update (1-based). If provided, overrides candidate selection.",
    )
    p.add_argument(
        "--asset-uuid",
        type=str,
        default=None,
        help="Arango UUID to update. If provided, finds the first row matching this UUID.",
    )
    p.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Max number of candidate rows to process (default: no limit).",
    )

    p.add_argument(
        "--apply",
        action="store_true",
        help="Actually execute PATCH requests (default: dry-run).",
    )
    p.add_argument(
        "--yes-really-apply-on-prd",
        action="store_true",
        help="Extra safety confirmation required when --apply is used with --env prd.",
    )

    return p


def _diagnose_eigenschappen(eminfra: EMInfraClient, asset_uuid: str) -> None:
    """Print helpful diagnostics when no eigenschapwaarden are returned."""

    try:
        kenmerken = eminfra.kenmerk_service.get_kenmerken_by_uuid(asset_uuid=asset_uuid)
        logger.warning("Diagnostics: asset has %d kenmerk(en)", len(kenmerken))
        for k in kenmerken:
            t = getattr(k, "type", None) or {}
            logger.warning("  kenmerk: %s | %s", t.get("uuid"), t.get("naam"))

        for k in kenmerken:
            t = getattr(k, "type", None) or {}
            naam = (t.get("naam") or "")
            uuid = t.get("uuid")
            if uuid and naam.startswith("Eigenschappen"):
                try:
                    eigs = eminfra.eigenschap_service.get_eigenschappen(asset_uuid=asset_uuid, kenmerk_uuid=uuid)
                    logger.warning("  eigenschapwaarden for %s (%s): %d", uuid, naam, len(eigs))
                    if eigs:
                        for ev in eigs[:3]:
                            logger.warning("    - %s = %s", ev.eigenschap.naam, ev.typedValue)
                except Exception:
                    logger.exception("  failed fetching eigenschapwaarden for kenmerk %s (%s)", uuid, naam)
    except Exception:
        logger.exception("Diagnostics failed while listing kenmerken for asset %s", asset_uuid)


def _extract_date_from_typedvalue(typedValue: Any) -> date | None:
    if not typedValue:
        return None
    # Expecting dict like {'_type':'date','value':'YYYY-MM-DD'} or a plain string
    try:
        if isinstance(typedValue, dict):
            val = typedValue.get("value")
            return _parse_date(val)
        return _parse_date(typedValue)
    except Exception:
        return None


def _apply_updates_for_asset(
    *,
    eminfra: EMInfraClient,
    asset_uuid: str,
    new_visit_date: date,
    new_result_text: str,
    apply: bool,
    arango_date: date | None = None,
) -> bool:
    """Prepare (and optionally apply) the two updates for one asset.

    Returns True if we attempted to apply (or would apply in dry-run), False if skipped.
    """

    eigs = fetch_all_eigenschappen(eminfra, asset_uuid)
    logger.info("Fetched %d eigenschapwaarden", len(eigs))

    # Determine current stored date, if any, from fetched eigenschappen
    current_db_date: date | None = None
    if eigs:
        ev_date = _find_eigenschap_by_name(eigs, "datum laatste keuring")
        if ev_date is not None:
            current_db_date = _extract_date_from_typedvalue(getattr(ev_date, "typedValue", None))

    if not eigs:
        _diagnose_eigenschappen(eminfra, asset_uuid)
        logger.warning(
            "No existing eigenschapwaarden found; will build update DTOs using fixed LSDeel Eigenschap UUIDs."
        )

        date_typed_value = {"_type": "date", "value": new_visit_date.isoformat()}
        result_typed_value = {"_type": "text", "value": new_result_text}

        # When no existing values: consult the arango_date provided (from the Excel 'Arango keuringsdatum' column)
        if arango_date is not None and new_visit_date <= arango_date:
            logger.info(
                "Skipping update for %s: Excel visit_date %s is not more recent than Arango date %s",
                asset_uuid,
                new_visit_date,
                arango_date,
            )
            return False

        date_update = _build_update_dto_by_eigenschap_uuid(LSDEEL_EIGENSCHAP_DATUM_UUID, date_typed_value)
        result_update = _build_update_dto_by_eigenschap_uuid(LSDEEL_EIGENSCHAP_RESULTAAT_UUID, result_typed_value)

        logger.info(
            "Prepared update (no existing values): datum laatste keuring uuid=%s -> %s",
            LSDEEL_EIGENSCHAP_DATUM_UUID,
            date_typed_value,
        )
        logger.info(
            "Prepared update (no existing values): resultaat keuring uuid=%s -> %s",
            LSDEEL_EIGENSCHAP_RESULTAAT_UUID,
            result_typed_value,
        )

        if not apply:
            return True

        eminfra.eigenschap_service.update_eigenschap_by_uuid(asset_uuid, date_update)
        eminfra.eigenschap_service.update_eigenschap_by_uuid(asset_uuid, result_update)
        return True

    # If we have existing eigenschapwaarden, compare current_db_date (if available)
    if eigs:
        ev_date = _find_eigenschap_by_name(eigs, "datum laatste keuring")
        ev_result = _find_eigenschap_by_name(eigs, "resultaat keuring")

        if ev_date is None or ev_result is None:
            logger.error(
                "Missing required eigenschap on asset %s | datum laatste keuring found=%s | resultaat keuring found=%s",
                asset_uuid,
                ev_date is not None,
                ev_result is not None,
            )
            return False

        logger.info(
            "Target eigenschap UUIDs | datum laatste keuring=%s | resultaat keuring=%s",
            getattr(ev_date.eigenschap, "uuid", None),
            getattr(ev_result.eigenschap, "uuid", None),
        )

        try:
            logger.info("Current typedValue (datum laatste keuring): %s", getattr(ev_date, "typedValue", None))
            logger.info("Current typedValue (resultaat keuring): %s", getattr(ev_result, "typedValue", None))
        except Exception:
            pass

        # If current DB date exists, only update when the Excel date is strictly more recent.
        if current_db_date is not None and new_visit_date <= current_db_date:
            logger.info(
                "Skipping update for %s: Excel visit_date %s is not more recent than current DB date %s",
                asset_uuid,
                new_visit_date,
                current_db_date,
            )
            return False

        date_typed_value = {"_type": "date", "value": new_visit_date.isoformat()}
        result_typed_value = {"_type": "text", "value": new_result_text}

        date_update = _build_update_dto(ev_date, date_typed_value)
        result_update = _build_update_dto(ev_result, result_typed_value)

        logger.info("Prepared update: datum laatste keuring -> %s", date_typed_value)
        logger.info("Prepared update: resultaat keuring -> %s", result_typed_value)

        if not apply:
            return True

        eminfra.eigenschap_service.update_eigenschap_by_uuid(asset_uuid, date_update)
        eminfra.eigenschap_service.update_eigenschap_by_uuid(asset_uuid, result_update)
        return True


def main() -> int:
    logging.basicConfig(
        format="%(asctime)s %(levelname)-8s %(message)s",
        level=logging.INFO,
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    args = _build_parser().parse_args()

    # Respect module-level DEFAULT_DRY_RUN unless user explicitly passed --apply on CLI
    user_provided_apply = "--apply" in sys.argv
    if not user_provided_apply:
        # DEFAULT_DRY_RUN == True -> args.apply = False (dry-run)
        # DEFAULT_DRY_RUN == False -> args.apply = True (apply by default)
        args.apply = not DEFAULT_DRY_RUN

    env_map = {
        "prd": Environment.PRD,
        "dev": Environment.DEV,
        "tei": Environment.TEI,
        "aim": Environment.AIM,
    }

    base_dir = Path(__file__).resolve().parent

    # Build list of excel files to process: either the single file provided via --excel,
    # or all files in the VincottePlatform dir that contain '.enriched' when --all is used
    # or when --excel is not provided.
    excel_paths: list[Path] = []
    if args.all or args.excel is None:
        candidates = [
            p for p in sorted(base_dir.iterdir()) if p.is_file() and p.suffix.lower() == ".xlsx" and ".enriched" in p.name.lower()
        ]
        if not candidates:
            logger.error("No enriched Excel files found in %s", base_dir)
            return 2
        excel_paths = candidates
    else:
        if not args.excel.exists():
            logger.error("Excel not found: %s", args.excel)
            return 2
        excel_paths = [args.excel]

    # Determine whether the --apply flag was provided on the command line.
    apply_passed_on_cli = "--apply" in sys.argv

    # PRD safety logic:
    # - If the user explicitly passed --apply on the CLI, require --yes-really-apply-on-prd.
    # - If args.apply is True because DEFAULT_DRY_RUN was changed in-file (no --apply in argv),
    #   treat that in-file override as an implicit confirmation: set args.yes_really_apply_on_prd
    #   and log a conspicuous warning. This lets you flip DEFAULT_DRY_RUN in the file without
    #   needing to pass an extra CLI flag.
    if args.env == "prd" and args.apply:
        if apply_passed_on_cli:
            if not args.yes_really_apply_on_prd:
                logger.error(
                    "Refusing to PATCH on PRD without --yes-really-apply-on-prd. "
                    "Run again with --apply --yes-really-apply-on-prd if you are sure."
                )
                return 2
        else:
            # In-file override: respect it and mark as confirmed.
            logger.warning(
                "DEFAULT_DRY_RUN has been disabled in-file; treating this as confirmation to apply on PRD."
            )
            args.yes_really_apply_on_prd = True

    # Create EMInfraClient once and reuse for all files
    eminfra = EMInfraClient(env=env_map[args.env], auth_type=AuthType.JWT, settings_path=args.settings)

    total_processed = total_attempted = total_applied = total_skipped = 0

    for excel_path in excel_paths:
        logger.info("Processing Excel: %s", excel_path.name)

        wb = load_workbook(excel_path, data_only=True)
        ws = wb.active

        visit_date_col = _find_visit_date_col(ws)

        uuid_col = _find_col_by_header(ws, "Arango UUID")
        if uuid_col is None:
            logger.error("Could not find 'Arango UUID' header in %s", excel_path.name)
            # skip this file but continue with others
            continue

        rows_to_process: list[int] = []

        if args.row:
            rows_to_process = [int(args.row)]
        elif args.asset_uuid:
            wanted = str(args.asset_uuid).strip()
            for r in range(2, ws.max_row + 1):
                v = ws.cell(row=r, column=uuid_col).value
                if v is not None and str(v).strip() == wanted:
                    rows_to_process = [r]
                    break
            if not rows_to_process:
                logger.error("No row found with Arango UUID=%s in %s", wanted, excel_path.name)
                # continue to next file, don't treat as fatal
                continue
        else:
            candidates = iter_candidates_from_excel(excel_path, debug=bool(args.debug_cols))
            logger.info("Found %d candidate row(s) in %s", len(candidates), excel_path.name)
            rows_to_process = [c.row_index for c in candidates]

        if args.limit is not None:
            rows_to_process = rows_to_process[: max(0, int(args.limit))]

        if not rows_to_process:
            logger.info("Nothing to do for %s.", excel_path.name)
            continue

        attempted = applied = skipped = 0

        for row in rows_to_process:
            asset_uuid_raw = ws.cell(row=row, column=uuid_col).value
            if asset_uuid_raw is None or str(asset_uuid_raw).strip() == "":
                skipped += 1
                logger.warning("Skipping row %d: missing Arango UUID", row)
                continue
            asset_uuid = str(asset_uuid_raw).strip()

            result_code_raw = ws.cell(row=row, column=2).value  # column B
            visit_date_raw = ws.cell(row=row, column=visit_date_col).value

            new_visit_date = _parse_date(visit_date_raw)
            new_result_text = _excel_result_code_to_text(result_code_raw)

            logger.info(
                "Row %d | Arango UUID=%s | Excel visit_date(col %d raw)=%r -> %s | Excel result(raw)=%r -> %r",
                row,
                asset_uuid,
                visit_date_col,
                visit_date_raw,
                new_visit_date,
                result_code_raw,
                new_result_text,
            )

            if new_visit_date is None:
                skipped += 1
                logger.error("Skipping row %d: could not parse visit_date", row)
                continue
            if new_result_text is None:
                skipped += 1
                logger.error("Skipping row %d: could not map result code", row)
                continue

            if _should_skip_due_to_result(new_result_text):
                skipped += 1
                logger.info("Skipping row %d: result indicates buiten gebruik (%r)", row, new_result_text)
                continue

            attempted += 1
            ok = _apply_updates_for_asset(
                eminfra=eminfra,
                asset_uuid=asset_uuid,
                new_visit_date=new_visit_date,
                new_result_text=new_result_text,
                apply=bool(args.apply),
                arango_date=_parse_date(ws.cell(row=row, column=_find_col_by_header(ws, "Arango keuringsdatum")).value),
            )

            if ok:
                if args.apply:
                    applied += 1
            else:
                skipped += 1

        if not args.apply:
            logger.warning("DRY-RUN for %s: not applying updates. Re-run with --apply to PATCH.", excel_path.name)

        logger.info(
            "Done for %s. rows=%d attempted=%d applied=%d skipped=%d",
            excel_path.name,
            len(rows_to_process),
            attempted,
            applied,
            skipped,
        )

        total_processed += len(rows_to_process)
        total_attempted += attempted
        total_applied += applied
        total_skipped += skipped

    logger.info(
        "All done. files=%d total_rows=%d attempted=%d applied=%d skipped=%d",
        len(excel_paths),
        total_processed,
        total_attempted,
        total_applied,
        total_skipped,
    )

    return 0


if __name__ == "__main__":
    # Configure default dry-run behavior here. Set to False if you want the script
    # to APPLY updates by default (unless overridden on the command line with --apply).
    # Be cautious when setting this to False when targeting PRD.
    DEFAULT_DRY_RUN = False  # keep dry-run as default for safety
    raise SystemExit(main())
