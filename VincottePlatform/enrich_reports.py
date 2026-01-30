from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

# python-arango import:
# - Some environments expose ArangoClient at `arango.ArangoClient`
# - Others expose it at `arango.client.ArangoClient`
# - If the *wrong* `arango` package is installed, imports will fail or the module will be incomplete.
try:
    from arango import ArangoClient  # type: ignore
except Exception:  # pragma: no cover
    try:
        from arango.client import ArangoClient  # type: ignore
    except Exception as e:  # pragma: no cover
        raise ImportError(
            "Failed to import ArangoClient. Make sure you have 'python-arango' installed, and NOT the legacy 'arango' package. "
            "If needed: pip uninstall -y arango && pip install -U python-arango"
        ) from e

from openpyxl import load_workbook

logger = logging.getLogger(__name__)

DEFAULT_SETTINGS_PATH = Path("/home/davidlinux/Documenten/AWV/resources/settings_SyncToArangoDB.json")


@dataclass(frozen=True)
class ArangoMatch:
    uuid: str
    naampad: str
    keuringsdatum: str | None = None


def normalize_installation_number(value: Any) -> str:
    """Normalize an installation number.

    Rules:
    - None -> ""
    - Convert to str
    - strip() leading/trailing whitespace
    - remove any whitespace inside the string (spaces, NBSP, tabs, ...)

    This matches the user's requirement: "kolom E is de installatie nummer; als er spaties zijn, negeer ze dan".
    """

    if value is None:
        return ""

    s = str(value).strip()
    if not s:
        return ""

    # Remove all whitespace characters including NBSP.
    return re.sub(r"\s+", "", s)


def load_settings(settings_path: Path) -> dict[str, Any]:
    with open(settings_path, "r", encoding="utf-8") as f:
        return json.load(f)


def create_arango_db(settings: dict[str, Any], env_key: str = "prd"):
    """Create a python-arango database connection from the AWV settings json."""

    db_settings = settings["databases"][env_key]

    # ArangoDBConnectionFactory in SampleArangoDBCode uses ArangoClient() without hosts,
    # which connects to localhost:8529 by default.
    client = ArangoClient()

    return client.db(
        name=db_settings["database"],
        username=db_settings["user"],
        password=db_settings["password"],
    )


AQL_ALL_ACTIVE_LSDEEL_ASSETS = """
LET lsdeel_key = FIRST(
  FOR at IN assettypes
    FILTER at.short_uri == \"lgc:installatie#LSDeel\"
    LIMIT 1
    RETURN at._key
)

FOR a IN assets
  FILTER a.NaampadObject_naampad != null
  FILTER a.AIMDBStatus_isActief == true
  FILTER a.assettype_key == lsdeel_key
  RETURN {
    uuid: a._key,
    naampad: a.NaampadObject_naampad,
    keuringsdatum: a.ins.EMObject_datumLaatsteKeuring
  }
"""


def fetch_all_active_lsdeel_assets(db) -> list[ArangoMatch]:
    """Fetch all active LSDeel assets upfront (uuid + naampad + keuringsdatum)."""

    cursor = db.aql.execute(AQL_ALL_ACTIVE_LSDEEL_ASSETS, batch_size=2000)
    results: list[ArangoMatch] = []
    for r in cursor:
        uuid = (r or {}).get("uuid")
        naampad = (r or {}).get("naampad")
        keuringsdatum = (r or {}).get("keuringsdatum")
        if uuid and naampad:
            results.append(ArangoMatch(uuid=uuid, naampad=naampad, keuringsdatum=keuringsdatum))
    return results


def _find_last_used_header_col(ws, header_row: int = 1) -> int:
    """Return the last column index that contains a non-empty header cell.

    Some Vincotte exports have formatting up to column Z/AA even though headers stop at M/N.
    Using ws.max_column would append our columns far to the right (e.g. starting at AA).
    """

    last = 0
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is None:
            continue
        if isinstance(v, str) and v.strip() == "":
            continue
        last = c
    return last


def _infer_installation_column_letter(ws) -> str:
    """Infer the installation number column from header row.

    Heuristic: the installation number is in the column whose header equals 'Visiting place'.
    If that doesn't exist, fallback to default E.

    Notes:
    - reports.xlsx has installation numbers under header 'Visiting place' (col E)
    - reports (2).xlsx has installation numbers under header 'Visiting place' (col F)
    """

    header_row = 1
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip().lower() == "visiting place":
            return chr(ord("A") + c - 1)
    return "E"


def _has_any_whitespace(value: Any) -> bool:
    if value is None:
        return False
    return bool(re.search(r"\s", str(value)))


def _normalize_for_matching(value: Any) -> str:
    return normalize_installation_number(value)


def extract_valid_installation_number(raw_installation: Any) -> Optional[str]:
    """
    Extract and validate an installation number.
    Accepts two valid formats:
      - (WO|WW|A|C|G) followed by exactly 4 digits, e.g. A1234, WO1234
      - 3 digits, then letter A/C/G, then 1 digit, e.g. 123A1
    The input may contain spaces around the letter (e.g. '123 A 1'), so we try multiple
    strategies:
      1. Check each whitespace-separated token after stripping
      2. Check the entire string once all whitespace is removed
      3. Fallback: strip non-alphanumerics from the first token and re-check

    Returns the cleaned installation number (uppercase, no spaces) or None if invalid.
    """
    pattern1 = re.compile(r'^(?:WO|WW|A|C|G)\d{4}$', re.IGNORECASE)
    pattern2 = re.compile(r'^\d{3}[ACG]\d$', re.IGNORECASE)

    if raw_installation is None:
        return None

    s = str(raw_installation).strip()
    if not s:
        return None

    # 1) Try each whitespace-separated token (covers many simple cases)
    for token in s.split():
        cand = re.sub(r'\s+', '', token).upper()
        if pattern1.match(cand) or pattern2.match(cand):
            return cand

    # 2) Try the whole string with all whitespace removed (covers '123 A 1', 'A 1234', etc.)
    s_no_ws = re.sub(r'\s+', '', s).upper()
    if pattern1.match(s_no_ws) or pattern2.match(s_no_ws):
        return s_no_ws

    # 3) Fallback: keep only alphanumeric from the first token and re-test
    first = s.split()[0]
    cand_alnum = re.sub(r'[^A-Z0-9]', '', first.upper())
    if pattern1.match(cand_alnum) or pattern2.match(cand_alnum):
        return cand_alnum

    return None


def _match_lsdeel_by_install_number(lsdeels: list[ArangoMatch], installation_number: str) -> tuple[str, Optional[ArangoMatch]]:
    """Return (status, match).

    status:
      - 'missing'   : 0 matches
      - 'ambiguous' : >1 matches
      - 'unique'    : exactly 1 match (match returned)
    """

    matches: list[ArangoMatch] = []
    for a in lsdeels:
        if installation_number in a.naampad:
            matches.append(a)
            if len(matches) > 1:
                return "ambiguous", None

    if len(matches) == 1:
        return "unique", matches[0]
    return "missing", None


def enrich_workbook(
    *,
    input_path: Path,
    output_path: Path,
    installation_column_letter: str,
    settings_path: Path = DEFAULT_SETTINGS_PATH,
    env_key: str = "prd",
    uuid_column_header: str = "Arango UUID",
    naampad_column_header: str = "Arango Naampad",
    keuringsdatum_column_header: str = "Arango keuringsdatum",
    outcome_column_header: str = "Uitkomst enrich script",
    sheet_name: str | None = None,
    row_start: int = 2,
    limit_rows: int | None = None,
) -> dict[str, int]:
    """Enrich a Vincotte Excel: add Arango uuid/naampad/keuringsdatum + outcome."""

    settings = load_settings(settings_path)
    db = create_arango_db(settings=settings, env_key=env_key)

    logger.info("Fetching active LSDeel assets from ArangoDB...")
    lsdeels = fetch_all_active_lsdeel_assets(db)
    logger.info("Fetched %d LSDeel assets.", len(lsdeels))

    wb = load_workbook(input_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    installation_col = ord(installation_column_letter.upper()) - ord("A") + 1

    header_row = 1

    header_values = {str(ws.cell(row=header_row, column=c).value).strip(): c for c in range(1, ws.max_column + 1)}

    need_create = any(
        h not in header_values
        for h in [uuid_column_header, naampad_column_header, keuringsdatum_column_header, outcome_column_header]
    )
    if not need_create:
        uuid_col = header_values[uuid_column_header]
        naampad_col = header_values[naampad_column_header]
        keuringsdatum_col = header_values[keuringsdatum_column_header]
        outcome_col = header_values[outcome_column_header]
    else:
        last_used = _find_last_used_header_col(ws, header_row=header_row)
        uuid_col = header_values.get(uuid_column_header, last_used + 1)
        used = max(last_used, uuid_col)
        naampad_col = header_values.get(naampad_column_header, used + 1)
        used = max(used, naampad_col)
        keuringsdatum_col = header_values.get(keuringsdatum_column_header, used + 1)
        used = max(used, keuringsdatum_col)
        outcome_col = header_values.get(outcome_column_header, used + 1)

        ws.cell(row=header_row, column=uuid_col, value=uuid_column_header)
        ws.cell(row=header_row, column=naampad_col, value=naampad_column_header)
        ws.cell(row=header_row, column=keuringsdatum_col, value=keuringsdatum_column_header)
        ws.cell(row=header_row, column=outcome_col, value=outcome_column_header)

    processed = filled = 0
    skipped_invalid_install = skipped_whitespace_install = 0
    missing = ambiguous = 0

    max_row = ws.max_row
    if limit_rows is not None:
        max_row = min(max_row, row_start + limit_rows - 1)

    cache: dict[str, tuple[str, Optional[ArangoMatch]]] = {}

    for r in range(row_start, max_row + 1):
        processed += 1
        raw_install = ws.cell(row=r, column=installation_col).value

        valid_install = extract_valid_installation_number(raw_install)
        if not valid_install:
            skipped_invalid_install += 1
            ws.cell(row=r, column=outcome_col, value="SKIP_INVALID")
            continue

        install = valid_install
        if install not in cache:
            try:
                cache[install] = _match_lsdeel_by_install_number(lsdeels, install)
            except Exception:
                logger.exception("Matching failed for installation_number=%r (raw=%r)", install, raw_install)
                cache[install] = ("missing", None)

        status, match = cache[install]

        if status == "unique" and match is not None:
            ws.cell(row=r, column=uuid_col, value=match.uuid)
            ws.cell(row=r, column=naampad_col, value=match.naampad)
            ws.cell(row=r, column=keuringsdatum_col, value=match.keuringsdatum)
            ws.cell(row=r, column=outcome_col, value="FILLED")
            filled += 1
        elif status == "ambiguous":
            ws.cell(row=r, column=outcome_col, value="AMBIGUOUS")
            ambiguous += 1
        else:
            ws.cell(row=r, column=outcome_col, value="MISSING")
            missing += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    return {
        "processed": processed,
        "filled": filled,
        "skipped_invalid_install": skipped_invalid_install,
        "skipped_whitespace_install": skipped_whitespace_install,
        "missing": missing,
        "ambiguous": ambiguous,
    }


def _default_output_for_input(input_path: Path) -> Path:
    if input_path.suffix.lower() == ".xlsx":
        return input_path.with_name(input_path.stem + ".enriched.xlsx")
    return input_path.with_name(input_path.name + ".enriched.xlsx")


def _merge_stats(total: dict[str, int], stats: dict[str, int]) -> dict[str, int]:
    for k, v in stats.items():
        total[k] = total.get(k, 0) + int(v)
    return total


def _enrich_workbook_with_preloaded_lsdeels(
    *,
    wb,
    ws,
    input_path: Path,
    output_path: Path,
    installation_column_letter: str,
    lsdeels: list[ArangoMatch],
    uuid_column_header: str = "Arango UUID",
    naampad_column_header: str = "Arango Naampad",
    keuringsdatum_column_header: str = "Arango keuringsdatum",
    outcome_column_header: str = "Uitkomst enrich script",
    row_start: int = 2,
    limit_rows: int | None = None,
) -> dict[str, int]:
    installation_col = ord(installation_column_letter.upper()) - ord("A") + 1

    header_row = 1
    header_values = {str(ws.cell(row=header_row, column=c).value).strip(): c for c in range(1, ws.max_column + 1)}

    need_create = any(
        h not in header_values
        for h in [uuid_column_header, naampad_column_header, keuringsdatum_column_header, outcome_column_header]
    )
    if not need_create:
        uuid_col = header_values[uuid_column_header]
        naampad_col = header_values[naampad_column_header]
        keuringsdatum_col = header_values[keuringsdatum_column_header]
        outcome_col = header_values[outcome_column_header]
    else:
        last_used = _find_last_used_header_col(ws, header_row=header_row)
        uuid_col = header_values.get(uuid_column_header, last_used + 1)
        used = max(last_used, uuid_col)
        naampad_col = header_values.get(naampad_column_header, used + 1)
        used = max(used, naampad_col)
        keuringsdatum_col = header_values.get(keuringsdatum_column_header, used + 1)
        used = max(used, keuringsdatum_col)
        outcome_col = header_values.get(outcome_column_header, used + 1)

        ws.cell(row=header_row, column=uuid_col, value=uuid_column_header)
        ws.cell(row=header_row, column=naampad_col, value=naampad_column_header)
        ws.cell(row=header_row, column=keuringsdatum_col, value=keuringsdatum_column_header)
        ws.cell(row=header_row, column=outcome_col, value=outcome_column_header)

    processed = filled = 0
    skipped_invalid_install = skipped_whitespace_install = 0
    missing = ambiguous = 0

    max_row = ws.max_row
    if limit_rows is not None:
        max_row = min(max_row, row_start + limit_rows - 1)

    cache: dict[str, tuple[str, Optional[ArangoMatch]]] = {}

    for r in range(row_start, max_row + 1):
        processed += 1
        raw_install = ws.cell(row=r, column=installation_col).value

        valid_install = extract_valid_installation_number(raw_install)
        if not valid_install:
            skipped_invalid_install += 1
            ws.cell(row=r, column=outcome_col, value="SKIP_INVALID")
            continue

        install = valid_install
        if install not in cache:
            cache[install] = _match_lsdeel_by_install_number(lsdeels, install)

        status, match = cache[install]

        if status == "unique" and match is not None:
            ws.cell(row=r, column=uuid_col, value=match.uuid)
            ws.cell(row=r, column=naampad_col, value=match.naampad)
            ws.cell(row=r, column=keuringsdatum_col, value=match.keuringsdatum)
            ws.cell(row=r, column=outcome_col, value="FILLED")
            filled += 1
        elif status == "ambiguous":
            ws.cell(row=r, column=outcome_col, value="AMBIGUOUS")
            ambiguous += 1
        else:
            ws.cell(row=r, column=outcome_col, value="MISSING")
            missing += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    return {
        "processed": processed,
        "filled": filled,
        "skipped_invalid_install": skipped_invalid_install,
        "skipped_whitespace_install": skipped_whitespace_install,
        "missing": missing,
        "ambiguous": ambiguous,
    }


def enrich_all_excels_in_dir(
    *,
    directory: Path,
    settings_path: Path = DEFAULT_SETTINGS_PATH,
    env_key: str = "prd",
    limit_rows: int | None = None,
) -> dict[str, dict[str, int]]:
    """Enrich all *.xlsx files in directory that do NOT have 'enriched' in their name."""

    directory = directory.resolve()
    results: dict[str, dict[str, int]] = {}

    settings = load_settings(settings_path)
    db = create_arango_db(settings=settings, env_key=env_key)

    logger.info("Fetching active LSDeel assets from ArangoDB (shared for all files)...")
    lsdeels = fetch_all_active_lsdeel_assets(db)
    logger.info("Fetched %d LSDeel assets.", len(lsdeels))

    candidates = [
        p
        for p in sorted(directory.iterdir())
        if p.is_file() and p.suffix.lower() == ".xlsx" and "enriched" not in p.name.lower()
    ]

    for input_path in candidates:
        try:
            wb = load_workbook(input_path)
            ws = wb.active
            inferred_col = _infer_installation_column_letter(ws)
            output_path = _default_output_for_input(input_path)

            stats = _enrich_workbook_with_preloaded_lsdeels(
                wb=wb,
                ws=ws,
                input_path=input_path,
                output_path=output_path,
                installation_column_letter=inferred_col,
                lsdeels=lsdeels,
                limit_rows=limit_rows,
            )

            results[str(input_path)] = stats
            logger.info("Enriched %s -> %s | %s", input_path.name, output_path.name, stats)
        except Exception:
            logger.exception("Failed to enrich %s", input_path)

    totals: dict[str, int] = {}
    for stats in results.values():
        _merge_stats(totals, stats)

    logger.info(
        "Totals | processed=%d filled=%d skipped_invalid=%d skipped_whitespace=%d missing=%d ambiguous=%d",
        totals.get("processed", 0),
        totals.get("filled", 0),
        totals.get("skipped_invalid_install", 0),
        totals.get("skipped_whitespace_install", 0),
        totals.get("missing", 0),
        totals.get("ambiguous", 0),
    )

    return results


def _build_parser():
    import argparse

    default_input = Path(__file__).resolve().parent / "reports.xlsx"

    p = argparse.ArgumentParser(
        description=(
            "Enrich Vincotte Excel(s) by looking up a unique LSDeel asset in ArangoDB based on installation number, "
            "and adding Arango UUID + Arango Naampad + Arango keuringsdatum columns.\n\n"
            "Modes:\n"
            "- Single file: processes VincottePlatform/reports.xlsx\n"
            "- Batch: use --all to process all *.xlsx in VincottePlatform without 'enriched' in the filename"
        )
    )

    p.add_argument("--all", action="store_true", help="Process all non-enriched .xlsx files in VincottePlatform/")

    p.add_argument("--input", type=Path, default=default_input, help="Input Excel path (single-file mode)")
    p.add_argument("--output", type=Path, default=None, help="Output Excel path (default: derived from input)")
    p.add_argument(
        "--installation-col",
        default=None,
        help="Excel column letter containing the installation number (default: inferred from header, else E)",
    )
    p.add_argument(
        "--settings",
        type=Path,
        default=DEFAULT_SETTINGS_PATH,
        help="Path to settings_SyncToArangoDB.json (default: %(default)s)",
    )
    p.add_argument(
        "--env",
        default="prd",
        help="Environment key inside settings['databases'] (default: %(default)s)",
    )
    p.add_argument(
        "--limit",
        type=int,
        default=None,
        help="Limit number of rows processed (starting at row 2); useful for smoke tests.",
    )
    p.add_argument(
        "--sheet",
        default=None,
        help="Optional sheet name; defaults to active sheet.",
    )

    return p


def main() -> int:
    logging.basicConfig(
        format="%(asctime)s %(levelname)-8s %(message)s",
        level=logging.INFO,
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    args = _build_parser().parse_args()

    if args.all:
        directory = Path(__file__).resolve().parent
        results = enrich_all_excels_in_dir(
            directory=directory,
            settings_path=args.settings,
            env_key=args.env,
            limit_rows=args.limit,
        )

        logger.info("Batch done. Files processed: %d", len(results))
        return 0

    input_path: Path = args.input
    output_path: Path = args.output or _default_output_for_input(input_path)

    if args.installation_col:
        installation_col = str(args.installation_col).strip()
    else:
        try:
            ws = load_workbook(input_path).active
            installation_col = _infer_installation_column_letter(ws)
        except Exception:
            installation_col = "E"

    logger.info("Input: %s", input_path)
    logger.info("Output: %s", output_path)
    logger.info("Installation column: %s", installation_col)

    stats = enrich_workbook(
        input_path=input_path,
        output_path=output_path,
        installation_column_letter=installation_col,
        settings_path=args.settings,
        env_key=args.env,
        sheet_name=args.sheet,
        limit_rows=args.limit,
    )

    logger.info("Done: %s", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

