from __future__ import annotations

import argparse
import json
import logging
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, cast

from openpyxl import load_workbook

try:
    from arango import ArangoClient  # type: ignore
except Exception:  # pragma: no cover
    from arango.client import ArangoClient  # type: ignore

logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_FILENAME = "MyVinotte Min.Vl.Gemeenschap Reports list.xlsx"
DEFAULT_INPUT_DIR = Path("/home/davidlinux/Documenten/AWV/Keuringen")
DEFAULT_INPUT_PATH = DEFAULT_INPUT_DIR / DEFAULT_INPUT_FILENAME
DEFAULT_SETTINGS_PATH = Path("/home/davidlinux/Documenten/AWV/resources/settings_SyncToArangoDB.json")
DEFAULT_ASSETTYPE_SHORT_URI = "onderdeel#Laagspanningsbord"
KNOWN_PROVINCE_CODES = {"AN", "OV", "VB", "WV", "LI", "VL", "BR"}
MANUAL_INSTALLATION_CORRECTIONS = {
    "0134A4": "013A4",
    "WO0206": "111A8.WO0206",
}

MANUAL_INPUT_NAAMPAD_MATCHES = {
    "KW0060TU": "WEV.Tunnel.KW0060/KW0060.HSCab/LSGEDEELTE",
    "010A8": "K511.010A8/K511.010A8.K/K511.010A8.LsDeel",
    "239A8": "K307.239A8/K307.239A8.K/K307.239A8.LSDeel",
    "C0518A": "C0518/KAST_A/KAST_A.LSDeel",
    "C0518B": "C0518/KAST_B/KAST_B.LSDeel",
}


@dataclass(frozen=True)
class AssetMatch:
    uuid: str
    name: str
    naampad: str
    naampad_first: str
    naampad_second: str


@dataclass(frozen=True)
class MivModuleMatch:
    uuid: str
    name: str
    naampad: str
    parent_path: str


def normalize_installation_number(value: Any) -> str:
    """Trim and remove all whitespace to get a stable match key."""
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    return re.sub(r"\s+", "", s).upper()


def _normalize_kw_variants(candidate: str) -> str:
    """Normalize KW/KO suffix variants to base code (4 digits)."""
    c = candidate.upper()
    # Allow slash variants like KW0557/TU or KW005/SL.
    c = re.sub(r"/(?:[A-Z0-9]+)$", "", c, flags=re.IGNORECASE)
    # Examples: KW0017BR1, KW0017BR, KW0557TU, KW0515PS, KW005SL -> base code.
    m = re.fullmatch(r"(K[OW]\d{4})(?:[A-Z]+\d*)$", c, flags=re.IGNORECASE)
    if m:
        return m.group(1).upper()
    return c


def _normalize_wo_ww_suffix(candidate: str) -> str:
    """Normalize WO/WW patterns with trailing letters (e.g. WO0972VL -> WO0972)."""
    c = candidate.upper()
    m = re.fullmatch(r"((?:WO|WW)\d{4})[A-Z]+", c, flags=re.IGNORECASE)
    if m:
        return m.group(1).upper()
    return c


def extract_valid_installation_number(value: Any) -> str:
    """Return normalized installation number when it matches expected patterns, else ''."""
    normalized = normalize_installation_number(value)
    if not normalized:
        return ""

    pattern_wo_ww = re.compile(r"^(?:WO|WW)\d{4}$", re.IGNORECASE)
    pattern_ko_kw = re.compile(r"^K[OW]\d{4}$", re.IGNORECASE)
    pattern_acg_4 = re.compile(r"^[ACG]\d{4}$", re.IGNORECASE)
    pattern_wad_3 = re.compile(r"^WAD\d{3}$", re.IGNORECASE)
    pattern_wod_3 = re.compile(r"^WOD\d{3}$", re.IGNORECASE)
    pattern_xt_3 = re.compile(r"^XT[A-Z]{3}$", re.IGNORECASE)
    pattern_oz_4 = re.compile(r"^OZ\d{4}$", re.IGNORECASE)
    pattern_met_5 = re.compile(r"^MET\d{5}$", re.IGNORECASE)
    pattern_3_l_1 = re.compile(r"^\d{3}[ACG]\d$", re.IGNORECASE)

    def is_valid_code(candidate: str) -> bool:
        return bool(
            pattern_wo_ww.fullmatch(candidate)
            or pattern_ko_kw.fullmatch(candidate)
            or pattern_acg_4.fullmatch(candidate)
            or pattern_wad_3.fullmatch(candidate)
            or pattern_wod_3.fullmatch(candidate)
            or pattern_xt_3.fullmatch(candidate)
            or pattern_oz_4.fullmatch(candidate)
            or pattern_met_5.fullmatch(candidate)
            or pattern_3_l_1.fullmatch(candidate)
        )

    if normalized in MANUAL_INSTALLATION_CORRECTIONS:
        return MANUAL_INSTALLATION_CORRECTIONS[normalized]

    normalized = _normalize_kw_variants(normalized)
    normalized = _normalize_wo_ww_suffix(normalized)

    # Normalize textual KW suffixes to base code.
    kw_with_suffix = re.fullmatch(r"(K[OW]\d{4})[A-Z]+\d*", normalized, flags=re.IGNORECASE)
    if kw_with_suffix:
        normalized = kw_with_suffix.group(1).upper()

    if is_valid_code(normalized):
        return normalized.upper()

    raw_text = str(value).strip()
    # Accept common labels like 'Installatie nr: A5348' by stripping the prefix first.
    without_label = re.sub(r"^\s*installatie\s*nr\s*[:\-\.]?\s*", "", raw_text, flags=re.IGNORECASE)

    token_patterns = [
        re.compile(r"(?<![A-Z0-9])K[OW]\s*\d{4}(?:\s*/\s*[A-Z0-9]+|[A-Z0-9]+)(?![A-Z0-9])", re.IGNORECASE),
        re.compile(r"(?<![A-Z0-9])K[OW]\s*\d{4}(?![A-Z0-9])", re.IGNORECASE),
        re.compile(r"(?<![A-Z0-9])(?:WO|WW)\s*\d{4}(?![A-Z0-9])", re.IGNORECASE),
        re.compile(r"(?<![A-Z0-9])[ACG]\s*\d{4}(?![A-Z0-9])", re.IGNORECASE),
        re.compile(r"(?<![A-Z0-9])WA[DO]\s*\d{3}(?![A-Z0-9])", re.IGNORECASE),
        re.compile(r"(?<![A-Z0-9])XT\s*[A-Z]{3}(?![A-Z0-9])", re.IGNORECASE),
        re.compile(r"(?<![A-Z0-9])OZ\s*\d{4}(?![A-Z0-9])", re.IGNORECASE),
        re.compile(r"(?<![A-Z0-9])MET\s*\d{5}(?![A-Z0-9])", re.IGNORECASE),
        re.compile(r"(?<![A-Z0-9])\d{3}\s*[ACG]\s*\d(?![A-Z0-9])", re.IGNORECASE),
    ]

    for search_text in (without_label, raw_text):
        for pattern in token_patterns:
            match = pattern.search(search_text)
            if not match:
                continue
            candidate = normalize_installation_number(match.group(0))
            candidate = _normalize_kw_variants(candidate)
            candidate = _normalize_wo_ww_suffix(candidate)
            kw_suffix = re.fullmatch(r"(K[OW]\d{4})[A-Z]+\d*", candidate, flags=re.IGNORECASE)
            if kw_suffix:
                candidate = kw_suffix.group(1).upper()
            if candidate in MANUAL_INSTALLATION_CORRECTIONS:
                return MANUAL_INSTALLATION_CORRECTIONS[candidate]
            if is_valid_code(candidate):
                return candidate.upper()

    return ""


def _parent_path(naampad: str) -> str:
    parts = naampad.split("/")
    if len(parts) <= 1:
        return ""
    return "/".join(parts[:-1])


def _extract_miv_module_name(value: Any) -> tuple[str, str | None]:
    """Extract MIV module name from free text and optional province code."""
    if value is None:
        return "", None

    text = str(value).upper()
    match = re.search(r"\bMIV\s*(\d{3})\b", text)
    if not match:
        return "", None

    digits = match.group(1)
    module_name = f"MIV{digits}"
    tail = text[match.end():]

    sat_match = re.search(r"\bSS\s*([1-2])\b", tail)
    if sat_match:
        module_name = f"MIV{digits}.{sat_match.group(1)}"

    province_code: str | None = None
    prov_match = re.search(r"\b([A-Z]{2})\s+LVE\b", tail)
    if prov_match:
        province_code = prov_match.group(1)

    return module_name, province_code


def load_settings(settings_path: Path) -> dict[str, Any]:
    with settings_path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def create_arango_db(settings: dict[str, Any], env_key: str):
    db_settings = settings["databases"][env_key]
    hosts = db_settings.get("hosts")

    if hosts:
        client = ArangoClient(hosts=hosts)
    else:
        # Keep compatibility with setups that rely on localhost defaults.
        client = ArangoClient()

    return client.db(
        name=db_settings["database"],
        username=db_settings["user"],
        password=db_settings["password"],
    )


AQL_ASSETS_BY_TYPE = """
LET assettype_key = FIRST(
  FOR at IN assettypes
    FILTER at.short_uri == @assettype_short_uri
    LIMIT 1
    RETURN at._key
)

FOR a IN assets
  FILTER assettype_key != null
  FILTER a.assettype_key == assettype_key
  FILTER a.AIMDBStatus_isActief == true
  FILTER a.NaampadObject_naampad != null
  RETURN {
    uuid: a._key,
    name: a.AIMNaamObject_naam,
    naampad: a.NaampadObject_naampad,
    naampad_parts: a.naampad_parts
  }
"""

AQL_MIV_MODULES = """
FOR a IN assets
  LET at = DOCUMENT('assettypes', a.assettype_key)
  FILTER at != null
  FILTER CONTAINS(at.short_uri, 'installatie#MIVModule')
  FILTER a.AIMDBStatus_isActief == true
  FILTER a.NaampadObject_naampad != null
  RETURN {
    uuid: a._key,
    name: a.AIMNaamObject_naam,
    naampad: a.NaampadObject_naampad
  }
"""


def fetch_active_assets_by_type(db, assettype_short_uri: str) -> list[AssetMatch]:
    cursor = db.aql.execute(
        AQL_ASSETS_BY_TYPE,
        bind_vars={"assettype_short_uri": assettype_short_uri},
        batch_size=2000,
    )
    rows: list[AssetMatch] = []
    for row in cursor:
        uuid = (row or {}).get("uuid")
        name = (row or {}).get("name")
        naampad = (row or {}).get("naampad")
        naampad_parts = (row or {}).get("naampad_parts")

        if not isinstance(uuid, str) or not isinstance(naampad, str):
            continue

        resolved_name = name if isinstance(name, str) else ""

        first_part = ""
        second_part = ""
        if isinstance(naampad_parts, list) and naampad_parts:
            candidate = naampad_parts[0]
            if isinstance(candidate, str):
                first_part = normalize_installation_number(candidate)
            if len(naampad_parts) > 1 and isinstance(naampad_parts[1], str):
                second_part = normalize_installation_number(naampad_parts[1])
        if not first_part:
            first_part = normalize_installation_number(naampad.split("/")[0])
        if not second_part:
            parts = naampad.split("/")
            if len(parts) > 1:
                second_part = normalize_installation_number(parts[1])

        rows.append(
            AssetMatch(
                uuid=uuid,
                name=resolved_name,
                naampad=naampad,
                naampad_first=first_part,
                naampad_second=second_part,
            )
        )
    return rows


def fetch_active_miv_modules(db) -> list[MivModuleMatch]:
    cursor = db.aql.execute(AQL_MIV_MODULES, batch_size=2000)
    rows: list[MivModuleMatch] = []
    for row in cursor:
        uuid = (row or {}).get("uuid")
        name = (row or {}).get("name")
        naampad = (row or {}).get("naampad")
        if not isinstance(uuid, str) or not isinstance(naampad, str):
            continue
        resolved_name = name if isinstance(name, str) else ""
        rows.append(
            MivModuleMatch(
                uuid=uuid,
                name=resolved_name,
                naampad=naampad,
                parent_path=_parent_path(naampad),
            )
        )
    return rows


def _match_asset_via_miv_tree(
    *,
    assets: list[AssetMatch],
    miv_modules: list[MivModuleMatch],
    miv_module_name: str,
) -> tuple[str, AssetMatch | None]:
    if not miv_module_name:
        return "missing", None

    module_matches = [m for m in miv_modules if m.name.upper() == miv_module_name.upper()]
    if not module_matches:
        return "missing", None
    if len(module_matches) > 1:
        return "ambiguous_name", None

    module = module_matches[0]
    siblings = [
        a for a in assets
        if _parent_path(a.naampad) == module.parent_path and a.name.upper().endswith(".LSDEEL")
    ]

    if len(siblings) == 1:
        return "unique_miv_tree", siblings[0]
    if len(siblings) > 1:
        return "ambiguous_name", None
    return "missing", None


def _find_last_used_header_col(ws, header_row: int = 1) -> int:
    last = 0
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=header_row, column=col).value
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        last = col
    return max(last, 1)


def _find_col_by_header(ws, header_name: str, header_row: int = 1) -> int | None:
    target = header_name.strip().lower()
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=header_row, column=col).value
        if isinstance(value, str) and value.strip().lower() == target:
            return col
    return None


def _match_asset_strict(assets: list[AssetMatch], installation_number: str) -> tuple[str, AssetMatch | None]:
    if not installation_number:
        return "invalid_installation", None

    expected_name = f"{installation_number}.LSDeel".upper()
    name_matches = [asset for asset in assets if asset.name.upper() == expected_name]

    if not name_matches:
        return "missing", None

    strict_matches = [asset for asset in name_matches if asset.naampad_first == installation_number]
    if len(strict_matches) == 1:
        return "unique", strict_matches[0]
    if len(strict_matches) > 1:
        return "ambiguous_name", None

    # Fallback: one unique name match is still considered deterministic.
    if len(name_matches) == 1:
        return "unique", name_matches[0]

    return "ambiguous_name", None


def _match_asset_by_naampad_first(assets: list[AssetMatch], installation_number: str) -> tuple[str, AssetMatch | None]:
    """Lower-confidence fallback: uniquely match by naampad first segment only."""
    if not installation_number:
        return "invalid_installation", None

    matches = [asset for asset in assets if asset.naampad_first == installation_number]
    if len(matches) == 1:
        return "unique_naampad_only", matches[0]
    if len(matches) > 1:
        return "ambiguous", None
    return "missing", None


def _extract_asw_target_name(value: Any) -> str:
    """Extract ASW-style name token and transform to expected LSDeel name."""
    if value is None:
        return ""
    raw = str(value)
    match = re.search(r"([A-Z0-9]+(?:\.[A-Z0-9]+)*)\.ASW\b", raw, flags=re.IGNORECASE)
    if not match:
        return ""
    prefix = match.group(1).upper()
    return f"{prefix}.GEB.LSDEEL"


def _match_asset_by_exact_name(assets: list[AssetMatch], expected_name_upper: str) -> tuple[str, AssetMatch | None]:
    if not expected_name_upper:
        return "missing", None
    matches = [asset for asset in assets if asset.name.upper() == expected_name_upper]
    if not matches:
        return "missing", None
    if len(matches) == 1:
        return "unique", matches[0]
    return "ambiguous_name", None


def _match_asset_by_exact_naampad(assets: list[AssetMatch], expected_naampad: str) -> tuple[str, AssetMatch | None]:
    if not expected_naampad:
        return "missing", None
    target = expected_naampad.upper()
    matches = [asset for asset in assets if asset.naampad.upper() == target]
    if not matches:
        return "missing", None
    if len(matches) == 1:
        return "unique", matches[0]
    return "ambiguous", None


def _default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}.enriched{input_path.suffix}")


def _resolve_input_path(cli_input: Path | None) -> Path:
    if cli_input is not None:
        return cli_input.resolve()

    if DEFAULT_INPUT_PATH.exists():
        return DEFAULT_INPUT_PATH.resolve()

    candidates = sorted(
        p for p in DEFAULT_INPUT_DIR.glob("*Reports list.xlsx")
        if ".enriched" not in p.name and not p.name.startswith("~$")
    )
    if candidates:
        return candidates[0].resolve()

    return DEFAULT_INPUT_PATH.resolve()


def _resolve_settings_path(cli_settings: Path | None) -> Path:
    if cli_settings is not None:
        return cli_settings.resolve()

    env_path = os.getenv("ARANGO_SETTINGS_PATH")
    if env_path:
        return Path(env_path).resolve()

    return DEFAULT_SETTINGS_PATH.resolve()


def _extract_xt_site_code(value: Any) -> str:
    """Extract XT site token (XT + 3 letters) from free text."""
    if value is None:
        return ""

    normalized = normalize_installation_number(value)
    if re.fullmatch(r"XT[A-Z]{3}", normalized):
        return normalized

    raw = str(value).upper()
    match = re.search(r"(?<![A-Z0-9])(XT\s*[A-Z]{3})(?![A-Z0-9])", raw)
    if not match:
        return ""

    candidate = normalize_installation_number(match.group(1))
    if re.fullmatch(r"XT[A-Z]{3}", candidate):
        return candidate
    return ""


def _match_asset_via_xt_naampad(
    *,
    assets: list[AssetMatch],
    xt_site_code: str,
) -> tuple[str, AssetMatch | None]:
    if not xt_site_code:
        return "missing", None

    matches = [
        asset
        for asset in assets
        if xt_site_code in asset.naampad.upper() and asset.name.upper().endswith(".LSDEEL")
    ]

    if len(matches) == 1:
        return "unique_xt", matches[0]
    if len(matches) > 1:
        return "ambiguous_name", None
    return "missing", None


def _match_asset_via_oz_naampad_second(
    *,
    assets: list[AssetMatch],
    oz_code: str,
) -> tuple[str, AssetMatch | None]:
    if not oz_code:
        return "missing", None

    matches = [
        asset
        for asset in assets
        if asset.naampad_second == oz_code and asset.name.upper().endswith(".LSDEEL")
    ]

    if len(matches) == 1:
        return "unique_oz", matches[0]
    if len(matches) > 1:
        return "ambiguous_name", None
    return "missing", None


def _format_outcome_text(status: str, source_label: str, manual_correction: bool = False) -> str:
    """Convert internal status to clear user-facing output text."""
    if status == "unique_oz":
        return f"Match via OZ-zone30 naampad via {source_label}"
    if status == "unique_xt":
        return f"Match via XT-site naampad via {source_label}"
    if status == "unique_miv_tree":
        return f"Match via MIV in boomstructuur via {source_label}"
    if status == "unique":
        suffix = " na manuele correctie" if manual_correction else ""
        return f"Match gevonden via {source_label}{suffix}"
    if status == "unique_naampad_only":
        return f"Match gevonden via {source_label}, enkel via naampad"
    if status == "invalid_installation":
        return "Geen geldig installatienummer in Visiting place of Description"
    if status == "missing":
        return f"Geen match gevonden via {source_label}"
    if status == "ambiguous_name":
        return f"Meerdere matches op naam via {source_label}; geen unieke asset"
    if status == "ambiguous":
        return f"Meerdere matches gevonden via {source_label}; geen unieke asset"
    return f"Onbekende uitkomst ({status}) via {source_label}"


def enrich_workbook(
    *,
    input_path: Path,
    output_path: Path,
    settings_path: Path,
    env_key: str,
    assettype_short_uri: str,
    installation_column_letter: str = "E",
    sheet_name: str | None = None,
    row_start: int = 2,
    limit_rows: int | None = None,
    uuid_column_header: str = "Arango UUID",
    name_column_header: str = "Arango Naam",
    naampad_column_header: str = "Arango Naampad",
    outcome_column_header: str = "Uitkomst enrich script",
) -> dict[str, int]:
    settings = load_settings(settings_path)
    db = create_arango_db(settings=settings, env_key=env_key)

    logger.info("Using settings: %s (env=%s)", settings_path, env_key)
    logger.info("Fetching active assets for assettype '%s'...", assettype_short_uri)
    try:
        assets = fetch_active_assets_by_type(db, assettype_short_uri=assettype_short_uri)
    except Exception as exc:
        raise RuntimeError(
            f"Arango query failed for env '{env_key}' with settings '{settings_path}'."
        ) from exc
    logger.info("Fetched %d assets.", len(assets))
    logger.info("Fetching active MIV modules...")
    miv_modules = fetch_active_miv_modules(db)
    logger.info("Fetched %d MIV modules.", len(miv_modules))

    wb = load_workbook(input_path)
    selected_ws = wb[sheet_name] if sheet_name else wb.active
    ws = cast(Any, selected_ws)

    header_row = 1
    visiting_place_col = _find_col_by_header(ws, "Visiting place", header_row=header_row)
    description_col = _find_col_by_header(ws, "Description", header_row=header_row)
    installation_col = ord(installation_column_letter.upper()) - ord("A") + 1
    if visiting_place_col is None:
        visiting_place_col = installation_col

    last_header_col = _find_last_used_header_col(ws, header_row=header_row)

    uuid_col = last_header_col + 1
    name_col = last_header_col + 2
    naampad_col = last_header_col + 3
    outcome_col = last_header_col + 4

    ws.cell(row=header_row, column=uuid_col, value=uuid_column_header)
    ws.cell(row=header_row, column=name_col, value=name_column_header)
    ws.cell(row=header_row, column=naampad_col, value=naampad_column_header)
    ws.cell(row=header_row, column=outcome_col, value=outcome_column_header)

    stats = {
        "processed": 0,
        "filled": 0,
        "filled_oz": 0,
        "filled_xt": 0,
        "filled_miv_tree": 0,
        "filled_naampad_only": 0,
        "invalid_installation": 0,
        "missing": 0,
        "ambiguous_name": 0,
        "ambiguous": 0,
    }
    unknown_province_codes: set[str] = set()

    max_row = ws.max_row if limit_rows is None else min(ws.max_row, row_start + limit_rows - 1)

    for row_idx in range(row_start, max_row + 1):
        stats["processed"] += 1

        visiting_raw = ws.cell(row=row_idx, column=visiting_place_col).value if visiting_place_col else None
        description_raw = ws.cell(row=row_idx, column=description_col).value if description_col else None

        install_from_visiting = extract_valid_installation_number(visiting_raw)
        install_from_description = extract_valid_installation_number(description_raw)
        xt_from_visiting = _extract_xt_site_code(visiting_raw)
        xt_from_description = _extract_xt_site_code(description_raw)
        oz_from_visiting = extract_valid_installation_number(visiting_raw) if visiting_raw is not None else ""
        oz_from_description = extract_valid_installation_number(description_raw) if description_raw is not None else ""
        miv_from_visiting, province_from_visiting = _extract_miv_module_name(visiting_raw)
        miv_from_description, province_from_description = _extract_miv_module_name(description_raw)

        manual_correction_applied = False
        manual_naampad_target = ""
        manual_naampad_source = ""

        if visiting_raw is not None:
            manual_naampad_target = MANUAL_INPUT_NAAMPAD_MATCHES.get(normalize_installation_number(visiting_raw), "")
            if manual_naampad_target:
                manual_naampad_source = "Visiting place"
        if not manual_naampad_target and description_raw is not None:
            manual_naampad_target = MANUAL_INPUT_NAAMPAD_MATCHES.get(normalize_installation_number(description_raw), "")
            if manual_naampad_target:
                manual_naampad_source = "Description"

        if visiting_raw is not None and install_from_visiting:
            normalized_visiting = normalize_installation_number(visiting_raw)
            for raw_value, corrected_value in MANUAL_INSTALLATION_CORRECTIONS.items():
                if install_from_visiting == corrected_value and raw_value in normalized_visiting:
                    manual_correction_applied = True
                    break
        if not install_from_visiting and description_raw is not None and install_from_description:
            normalized_description = normalize_installation_number(description_raw)
            for raw_value, corrected_value in MANUAL_INSTALLATION_CORRECTIONS.items():
                if install_from_description == corrected_value and raw_value in normalized_description:
                    manual_correction_applied = True
                    break

        candidate_install = install_from_visiting or install_from_description
        source_label = "Visiting place" if install_from_visiting else ("Description" if install_from_description else "none")

        if manual_naampad_target:
            status, match = _match_asset_by_exact_naampad(assets, manual_naampad_target)
            if status in {"unique", "ambiguous"}:
                source_label = manual_naampad_source or source_label
                manual_correction_applied = True
            else:
                status, match = _match_asset_strict(assets, candidate_install)
        else:
            status, match = _match_asset_strict(assets, candidate_install)

        if status == "missing" and candidate_install:
            naampad_only_status, naampad_only_match = _match_asset_by_naampad_first(assets, candidate_install)
            if naampad_only_status in {"unique_naampad_only", "ambiguous"}:
                status, match = naampad_only_status, naampad_only_match

        # ASW fallback: VASA01.ASW -> VASA01.GEB.LSDeel
        if status in {"missing", "invalid_installation"}:
            asw_target = _extract_asw_target_name(visiting_raw) or _extract_asw_target_name(description_raw)
            if asw_target:
                asw_status, asw_match = _match_asset_by_exact_name(assets, asw_target)
                if asw_status in {"unique", "ambiguous_name"}:
                    status, match = asw_status, asw_match
                    source_label = f"{source_label} (ASW->GEB.LSDeel)"

        if status in {"missing", "invalid_installation"}:
            miv_candidate = miv_from_visiting or miv_from_description
            if miv_candidate:
                miv_status, miv_match = _match_asset_via_miv_tree(
                    assets=assets,
                    miv_modules=miv_modules,
                    miv_module_name=miv_candidate,
                )
                if miv_status in {"unique_miv_tree", "ambiguous_name"}:
                    status, match = miv_status, miv_match
                    source_label = "Visiting place" if miv_from_visiting else "Description"

                province_code = province_from_visiting or province_from_description
                if province_code and province_code not in KNOWN_PROVINCE_CODES:
                    unknown_province_codes.add(province_code)

        if status in {"missing", "invalid_installation"}:
            xt_candidate = xt_from_visiting or xt_from_description
            if xt_candidate:
                xt_status, xt_match = _match_asset_via_xt_naampad(
                    assets=assets,
                    xt_site_code=xt_candidate,
                )
                if xt_status in {"unique_xt", "ambiguous_name"}:
                    status, match = xt_status, xt_match
                    source_label = "Visiting place" if xt_from_visiting else "Description"

        if status in {"missing", "invalid_installation"}:
            oz_candidate = oz_from_visiting or oz_from_description
            if oz_candidate:
                oz_status, oz_match = _match_asset_via_oz_naampad_second(
                    assets=assets,
                    oz_code=oz_candidate,
                )
                if oz_status in {"unique_oz", "ambiguous_name"}:
                    status, match = oz_status, oz_match
                    source_label = "Visiting place" if oz_from_visiting else "Description"

        ws.cell(
            row=row_idx,
            column=outcome_col,
            value=_format_outcome_text(status, source_label, manual_correction=manual_correction_applied),
        )

        if status in {"unique", "unique_miv_tree", "unique_xt", "unique_oz", "unique_naampad_only"} and match is not None:
            ws.cell(row=row_idx, column=uuid_col, value=match.uuid)
            ws.cell(row=row_idx, column=name_col, value=match.name)
            ws.cell(row=row_idx, column=naampad_col, value=match.naampad)
            stats["filled"] += 1
            if status == "unique_miv_tree":
                stats["filled_miv_tree"] += 1
            if status == "unique_xt":
                stats["filled_xt"] += 1
            if status == "unique_oz":
                stats["filled_oz"] += 1
            if status == "unique_naampad_only":
                stats["filled_naampad_only"] += 1
        else:
            if status in stats:
                stats[status] += 1
            else:
                stats["missing"] += 1

    wb.save(output_path)
    if unknown_province_codes:
        logger.warning(
            "Unknown province code(s) in MIV input encountered: %s",
            ", ".join(sorted(unknown_province_codes)),
        )
    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Enrich Vincotte Excel with Arango asset UUIDs.")
    parser.add_argument(
        "--input",
        type=Path,
        default=None,
        help=(
            "Input Excel file path (default: %s; fallback: first '*Reports list.xlsx' in %s)"
            % (DEFAULT_INPUT_PATH, DEFAULT_INPUT_DIR)
        ),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output Excel path (default: <input>.enriched.xlsx). Input is never overwritten.",
    )
    parser.add_argument(
        "--settings",
        type=Path,
        default=None,
        help="Path to Arango settings JSON (default: ARANGO_SETTINGS_PATH or AWV settings path)",
    )
    parser.add_argument("--env", default="prd", help="Settings env key (default: prd)")
    parser.add_argument(
        "--assettype-short-uri",
        default=DEFAULT_ASSETTYPE_SHORT_URI,
        help="Assettype short_uri to fetch from Arango",
    )
    parser.add_argument("--installation-col", default="E", help="Excel installation column letter")
    parser.add_argument("--sheet", default=None, help="Worksheet name (default: active sheet)")
    parser.add_argument("--row-start", type=int, default=2, help="Data row start (default: 2)")
    parser.add_argument("--limit", type=int, default=None, help="Optional max number of rows to process")
    return parser.parse_args()


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    args = parse_args()

    input_path = _resolve_input_path(args.input)
    settings_path = _resolve_settings_path(args.settings)

    if args.output:
        output_path = args.output.resolve()
        if output_path == input_path:
            logger.warning("Output equals input; forcing non-destructive output path.")
            output_path = _default_output_path(input_path)
    else:
        output_path = _default_output_path(input_path)

    if not input_path.exists():
        logger.error("Input file does not exist: %s", input_path)
        return 1
    if not settings_path.exists():
        logger.error("Settings file does not exist: %s", settings_path)
        return 2

    try:
        stats = enrich_workbook(
            input_path=input_path,
            output_path=output_path,
            settings_path=settings_path,
            env_key=args.env,
            assettype_short_uri=args.assettype_short_uri,
            installation_column_letter=args.installation_col,
            sheet_name=args.sheet,
            row_start=args.row_start,
            limit_rows=args.limit,
        )
    except Exception as exc:
        logger.error("Enrichment failed: %s", exc)
        return 3

    logger.info("Done: %s", output_path)
    logger.info("Stats: %s", stats)
    logger.info("Applied strict mapping on AIMNaamObject_naam and naampad first part.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

