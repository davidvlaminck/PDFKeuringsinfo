from __future__ import annotations

import argparse
import datetime as dt
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
from pypdf import PdfReader
from tqdm import tqdm

logger = logging.getLogger(__name__)

DEFAULT_INPUT_DIR = Path("/home/davidlinux/Documenten/AWV/Keuringen")
DEFAULT_INPUT_FILENAME = "MyVinotte Min.Vl.Gemeenschap Reports list.enriched.xlsx"
DEFAULT_INPUT_PATH = DEFAULT_INPUT_DIR / DEFAULT_INPUT_FILENAME
DEFAULT_SINCE_DATE = "2021-01-01"


@dataclass(frozen=True)
class InfractionEntry:
    code: str
    description: str


@dataclass(frozen=True)
class CodeMapping:
    mapping_code: str
    mapping_context: str
    mapping_source: str


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def _normalize_site_id(value: Any) -> str:
    text = _normalize_text(value)
    if text.endswith(".0") and text.replace(".", "", 1).isdigit():
        return text[:-2]
    return text


def _normalize_label(value: Any) -> str:
    token = _normalize_text(value).upper().replace(" ", "")
    red_tokens = {"RED", "ROOD", "R", "NIETCONFORM", "NONCONFORM", "NC"}
    if token in red_tokens:
        return "RED"
    return token


def _resolve_column(df: pd.DataFrame, candidates: list[str], fallback_index: int | None = None) -> str:
    normalized = {str(col).strip().lower(): str(col) for col in df.columns}
    for candidate in candidates:
        found = normalized.get(candidate.strip().lower())
        if found is not None:
            return found
    if fallback_index is not None and 0 <= fallback_index < len(df.columns):
        return str(df.columns[fallback_index])
    raise ValueError(f"Kolom niet gevonden. Gezochte headers: {candidates}")


def _read_pdf_text(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    parts: list[str] = []
    for page in reader.pages:
        try:
            parts.append(page.extract_text() or "")
        except Exception:
            parts.append("")
    return "\n".join(parts)


def _find_pdf_path(base_dir: Path, site_id: str, filename: str) -> Path | None:
    if not site_id or not filename:
        return None
    site_dir = base_dir / site_id
    if not site_dir.exists() or not site_dir.is_dir():
        return None

    raw_name = filename.strip()
    candidates = [raw_name]
    if not raw_name.lower().endswith(".pdf"):
        candidates.append(f"{raw_name}.pdf")

    for name in candidates:
        path = site_dir / name
        if path.exists() and path.is_file():
            return path

    lower_map = {p.name.lower(): p for p in site_dir.iterdir() if p.is_file()}
    for name in candidates:
        hit = lower_map.get(name.lower())
        if hit is not None:
            return hit
    return None


def _is_section_stop(line: str) -> bool:
    normalized = _normalize_text(line).upper()
    if not normalized:
        return False
    return bool(
        re.match(
            r"^(?:[IVX]{1,4}[.)]?\s+)?(?:OPMERKINGEN|TABELLEN|BESLUIT|CONCLUSIE|ALGEMENE\s+TECHNISCHE|METINGEN)\b",
            normalized,
        )
    )


def _extract_infractions_from_text(pdf_text: str) -> list[InfractionEntry]:
    if not pdf_text.strip():
        return []

    lines = [_normalize_text(line) for line in pdf_text.splitlines()]
    section_starts = [
        i
        for i, line in enumerate(lines)
        if re.match(r"^(?:III\.?\s+)?INBREUKEN\b", line, flags=re.IGNORECASE)
    ]
    if not section_starts:
        section_starts = [
            i
            for i, line in enumerate(lines)
            if re.search(r"\bVastgestelde\s+inbreuken\b", line, flags=re.IGNORECASE)
        ]

    candidate_sections: list[list[str]] = []
    for start in section_starts:
        section: list[str] = []
        for line in lines[start + 1 :]:
            if _is_section_stop(line):
                break
            section.append(line)
        if section:
            candidate_sections.append(section)

    if not candidate_sections:
        candidate_sections = [lines]

    best_entries: list[InfractionEntry] = []
    # Some templates prefix infractions with "I" (inbreuk) or "O" (opmerking), e.g. "I 1522 ...".
    entry_start_re = re.compile(r"^(?:[IO]\s+)?(?P<code>F\d{1,3}|[1-9]\d{3})\b\s*(?P<rest>.*)$", re.IGNORECASE)

    for section in candidate_sections:
        entries: list[InfractionEntry] = []
        current_code = ""
        current_parts: list[str] = []

        def flush_current() -> None:
            nonlocal current_code, current_parts
            if not current_code:
                return
            description = _normalize_text(" ".join(current_parts))
            entries.append(InfractionEntry(code=current_code.upper(), description=description))
            current_code = ""
            current_parts = []

        for line in section:
            if not line:
                continue
            if _is_section_stop(line):
                break
            if re.match(r"^(Blz\s*:|Verslagnummer\s*:|Contractnummer\s*:)", line, flags=re.IGNORECASE):
                continue
            if re.match(r"^[0-9\s]{10,}$", line):
                continue

            m = entry_start_re.match(line)
            if m:
                code = m.group("code").upper()
                rest = _normalize_text(m.group("rest"))
                # Exclude postal code/city lines like "3000 Leuven" but keep real code+description lines.
                if re.match(r"^(?:[IO]\s+)?[1-9]\d{3}\s+[A-Za-zÀ-ÿ\- ]{2,40}$", line):
                    continue
                flush_current()
                current_code = code
                current_parts = [rest] if rest else []
                continue

            if current_code:
                # Continue multiline description until next code/header.
                current_parts.append(line)

        flush_current()

        if len(entries) > len(best_entries):
            best_entries = entries

    if not best_entries and candidate_sections:
        # Some PDFs put a short "INBREUKEN / OPMERKINGEN / ..." menu first.
        # If section slicing misses the real list, retry against full document lines.
        candidate_sections = [lines]
        for section in candidate_sections:
            entries: list[InfractionEntry] = []
            current_code = ""
            current_parts: list[str] = []

            def flush_current_retry() -> None:
                nonlocal current_code, current_parts
                if not current_code:
                    return
                description = _normalize_text(" ".join(current_parts))
                entries.append(InfractionEntry(code=current_code.upper(), description=description))
                current_code = ""
                current_parts = []

            for line in section:
                if not line:
                    continue
                if re.match(r"^(Blz\s*:|Verslagnummer\s*:|Contractnummer\s*:)", line, flags=re.IGNORECASE):
                    continue
                if re.match(r"^[0-9\s]{10,}$", line):
                    continue

                m = entry_start_re.match(line)
                if m:
                    code = m.group("code").upper()
                    rest = _normalize_text(m.group("rest"))
                    if re.match(r"^(?:[IO]\s+)?[1-9]\d{3}\s+[A-Za-zÀ-ÿ\- ]{2,40}$", line):
                        continue
                    flush_current_retry()
                    current_code = code
                    current_parts = [rest] if rest else []
                    continue

                if current_code:
                    current_parts.append(line)

            flush_current_retry()
            if len(entries) > len(best_entries):
                best_entries = entries

    # Drop empty descriptions duplicates while keeping first hit order.
    seen: set[tuple[str, str]] = set()
    cleaned: list[InfractionEntry] = []
    for entry in best_entries:
        key = (entry.code, entry.description)
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(entry)
    return cleaned


def _parse_arei_mapping(arei_path: Path) -> tuple[dict[str, tuple[str, str]], list[tuple[int, int, str, str]], list[tuple[int, int, str, str]]]:
    df = pd.read_excel(arei_path, sheet_name=0, header=None)
    exact_fcode: dict[str, tuple[str, str]] = {}
    fcode_ranges: list[tuple[int, int, str, str]] = []
    numeric_ranges: list[tuple[int, int, str, str]] = []

    for _, row in df.iterrows():
        c0 = _normalize_text(row.iloc[0])
        c1 = _normalize_text(row.iloc[1])
        c2 = _normalize_text(row.iloc[2])
        if not c0 and not c1:
            continue

        # Mapping for F-code family labels.
        if re.fullmatch(r"F\d+", c0, flags=re.IGNORECASE):
            exact_fcode[c0.upper()] = (c0.upper(), c2)
        f_range = re.fullmatch(r"F(\d+)\s*-\s*(\d+)", c0, flags=re.IGNORECASE)
        if f_range:
            start, end = int(f_range.group(1)), int(f_range.group(2))
            fcode_ranges.append((start, end, c0.upper(), c2))

        # Mapping for numeric AREI ranges like 1101-1105.
        for match in re.finditer(r"(\d{4})\s*[-–]\s*(\d{4})", c1):
            s, e = int(match.group(1)), int(match.group(2))
            numeric_ranges.append((s, e, c0.upper(), c2))
        if re.fullmatch(r"\d{4}", c1):
            v = int(c1)
            numeric_ranges.append((v, v, c0.upper(), c2))

    return exact_fcode, fcode_ranges, numeric_ranges


def _map_infraction_code(
    code: str,
    exact_fcode: dict[str, tuple[str, str]],
    fcode_ranges: list[tuple[int, int, str, str]],
    numeric_ranges: list[tuple[int, int, str, str]],
) -> CodeMapping:
    up = code.upper()
    if re.fullmatch(r"F\d+", up):
        if up in exact_fcode:
            mapping_code, context = exact_fcode[up]
            return CodeMapping(mapping_code, context, "AREI_FCODE")
        n = int(up[1:])
        for start, end, family, context in fcode_ranges:
            if start <= n <= end:
                return CodeMapping(family, context, "AREI_FCODE_RANGE")
        return CodeMapping("UNMAPPED", "", "UNMAPPED")

    if re.fullmatch(r"[1-9]\d{3}", up):
        v = int(up)
        for start, end, family, context in numeric_ranges:
            if start <= v <= end:
                return CodeMapping(family, context, "AREI_NUMERIC_RANGE")
        return CodeMapping("UNMAPPED", "", "UNMAPPED")

    return CodeMapping("UNMAPPED", "", "UNMAPPED")


def _extract_postal_code(city_value: Any) -> int | None:
    m = re.match(r"\s*(\d{4})\b", str(city_value or ""))
    if not m:
        return None
    return int(m.group(1))


def _province_from_postal(postal_code: int | None) -> str:
    if postal_code is None:
        return ""
    # Belgian postal-range approximation focused on Flemish data + Brussels.
    if 1000 <= postal_code <= 1299:
        return "Brussels"
    if 1500 <= postal_code <= 1999 or 3000 <= postal_code <= 3499:
        return "Vlaams-Brabant"
    if 2000 <= postal_code <= 2999:
        return "Antwerpen"
    if 3500 <= postal_code <= 3999:
        return "Limburg"
    if 8000 <= postal_code <= 8999:
        return "West-Vlaanderen"
    if 9000 <= postal_code <= 9999:
        return "Oost-Vlaanderen"
    return ""


def _province_from_text(description: str) -> str:
    text = description.upper()
    if re.search(r"ANTWERP", text):
        return "Antwerpen"
    if re.search(r"LIMBURG", text):
        return "Limburg"
    if re.search(r"WEST[\s\-]?VLAAND", text) or re.search(r"\bWVL\b", text):
        return "West-Vlaanderen"
    if re.search(r"OOST[\s\-]?VLAAND", text) or re.search(r"\bOVL\b", text):
        return "Oost-Vlaanderen"
    if re.search(r"VL[\s\-]?BRABANT|VLAAMS[\s\-]?BRABANT", text):
        return "Vlaams-Brabant"

    miv_prov = re.search(r"\b(AN|OV|WV|LI|VB|BR)\s+LVE\b", text)
    if miv_prov:
        code = miv_prov.group(1)
        return {
            "AN": "Antwerpen",
            "OV": "Oost-Vlaanderen",
            "WV": "West-Vlaanderen",
            "LI": "Limburg",
            "VB": "Vlaams-Brabant",
            "BR": "Brussels",
        }.get(code, "")

    return ""


def _infer_province(city_value: Any, description: Any) -> str:
    postal_code = _extract_postal_code(city_value)
    from_postal = _province_from_postal(postal_code)
    if from_postal:
        return from_postal
    from_text = _province_from_text(_normalize_text(description))
    if from_text:
        return from_text
    return "UNKNOWN"


def _build_asset_key(row: pd.Series, cols: dict[str, str]) -> str:
    uuid = _normalize_text(row.get(cols["arango_uuid"]))
    if uuid:
        return uuid

    naampad = _normalize_text(row.get(cols["arango_naampad"]))
    if naampad:
        return f"NAAMPAD::{naampad.upper()}"

    site = _normalize_site_id(row.get(cols["site"]))
    place = _normalize_text(row.get(cols["visiting_place"]))
    if site or place:
        return f"FALLBACK::{site}::{place.upper()}"

    return "UNKNOWN_ASSET"


def extract_red_infractions(
    *,
    input_excel: Path,
    output_prefix: Path,
    arei_excel: Path,
    since_date: str,
) -> dict[str, int]:
    df = pd.read_excel(input_excel)

    cols = {
        "site": _resolve_column(df, ["Site ID", "SiteID"], fallback_index=0),
        "filename": _resolve_column(df, ["FILENAME", "Filename"], fallback_index=11),
        "label": _resolve_column(df, ["Label"]),
        "visit_date": _resolve_column(df, ["Visit Date", "Visit date"]),
        "arango_uuid": _resolve_column(df, ["Arango UUID"]),
        "arango_naam": _resolve_column(df, ["Arango Naam"]),
        "arango_naampad": _resolve_column(df, ["Arango Naampad"]),
        "city": _resolve_column(df, ["City"], fallback_index=10),
        "description": _resolve_column(df, ["Description"], fallback_index=2),
        "visiting_place": _resolve_column(df, ["Visiting place", "Visiting Place"], fallback_index=8),
        "site_name": _resolve_column(df, ["Site name", "Site Name"], fallback_index=1),
    }

    visit_dates = pd.to_datetime(df[cols["visit_date"]], errors="coerce", dayfirst=True)
    since_ts = pd.Timestamp(since_date)

    working = df.copy()
    working["__visit_date"] = visit_dates
    working["__label_norm"] = working[cols["label"]].map(_normalize_label)
    working["__asset_key"] = working.apply(lambda row: _build_asset_key(row, cols), axis=1)

    # Keep most recent inspection per asset.
    working = working[working["__visit_date"].notna() & (working["__visit_date"] > since_ts)]
    working = working.sort_values(["__asset_key", "__visit_date"], ascending=[True, False])
    latest = working.groupby("__asset_key", as_index=False).head(1)

    red_latest = latest[latest["__label_norm"] == "RED"].copy()

    exact_fcode, fcode_ranges, numeric_ranges = _parse_arei_mapping(arei_excel)

    detail_rows: list[dict[str, Any]] = []
    stats = {
        "assets_scanned": int(len(red_latest)),
        "pdf_missing": 0,
        "pdf_unreadable": 0,
        "assets_with_no_infractions_found": 0,
        "infractions_extracted": 0,
        "unmapped_infractions": 0,
    }

    base_dir = input_excel.parent

    for _, row in tqdm(red_latest.iterrows(), total=len(red_latest), desc="Inbreuken extractie", unit="asset"):
        site_id = _normalize_site_id(row.get(cols["site"]))
        filename = _normalize_text(row.get(cols["filename"]))
        pdf_path = _find_pdf_path(base_dir, site_id, filename)
        if pdf_path is None:
            stats["pdf_missing"] += 1
            continue

        try:
            pdf_text = _read_pdf_text(pdf_path)
        except Exception:
            stats["pdf_unreadable"] += 1
            continue

        infractions = _extract_infractions_from_text(pdf_text)
        if not infractions:
            stats["assets_with_no_infractions_found"] += 1
            continue

        province = _infer_province(row.get(cols["city"]), row.get(cols["description"]))
        for item in infractions:
            mapping = _map_infraction_code(item.code, exact_fcode, fcode_ranges, numeric_ranges)
            if mapping.mapping_code == "UNMAPPED":
                stats["unmapped_infractions"] += 1

            detail_rows.append(
                {
                    "asset_key": row["__asset_key"],
                    "arango_uuid": _normalize_text(row.get(cols["arango_uuid"])),
                    "arango_naam": _normalize_text(row.get(cols["arango_naam"])),
                    "arango_naampad": _normalize_text(row.get(cols["arango_naampad"])),
                    "site_id": site_id,
                    "site_name": _normalize_text(row.get(cols["site_name"])),
                    "description": _normalize_text(row.get(cols["description"])),
                    "city": _normalize_text(row.get(cols["city"])),
                    "province": province,
                    "visit_date": row["__visit_date"].date().isoformat(),
                    "filename": filename,
                    "pdf_path": str(pdf_path),
                    "infraction_code_raw": item.code,
                    "infraction_text_raw": item.description,
                    "infraction_family": mapping.mapping_code,
                    "infraction_family_context": mapping.mapping_context,
                    "mapping_source": mapping.mapping_source,
                }
            )

    detail_df = pd.DataFrame(detail_rows)
    stats["infractions_extracted"] = int(len(detail_df))

    detail_csv = output_prefix.with_suffix(".infractions-detail.csv")
    detail_xlsx = output_prefix.with_suffix(".infractions-detail.xlsx")
    detail_df.to_csv(detail_csv, index=False)
    detail_df.to_excel(detail_xlsx, index=False)

    if detail_df.empty:
        pivot_df = pd.DataFrame(columns=["province"])
        per_asset_df = pd.DataFrame(columns=["asset_key", "province", "visit_date", "filename", "infractions"])
    else:
        pivot_df = (
            detail_df.assign(infraction_key=detail_df["infraction_family"].where(detail_df["infraction_family"] != "UNMAPPED", detail_df["infraction_code_raw"]))
            .pivot_table(index="province", columns="infraction_key", values="asset_key", aggfunc="count", fill_value=0)
            .reset_index()
        )

        per_asset_df = (
            detail_df.assign(
                infraction_key=detail_df["infraction_family"].where(
                    detail_df["infraction_family"] != "UNMAPPED", detail_df["infraction_code_raw"]
                )
            )
            .groupby(["asset_key", "arango_uuid", "arango_naam", "arango_naampad", "province", "visit_date", "filename"], as_index=False)
            .agg(
                infractions=("infraction_key", lambda s: "; ".join(sorted(dict.fromkeys(s)))),
                infractions_raw=("infraction_code_raw", lambda s: "; ".join(sorted(dict.fromkeys(s)))),
            )
        )

    pivot_csv = output_prefix.with_suffix(".infractions-pivot-province.csv")
    pivot_xlsx = output_prefix.with_suffix(".infractions-pivot-province.xlsx")
    per_asset_csv = output_prefix.with_suffix(".infractions-per-asset.csv")
    per_asset_xlsx = output_prefix.with_suffix(".infractions-per-asset.xlsx")

    pivot_df.to_csv(pivot_csv, index=False)
    pivot_df.to_excel(pivot_xlsx, index=False)
    per_asset_df.to_csv(per_asset_csv, index=False)
    per_asset_df.to_excel(per_asset_xlsx, index=False)

    logger.info("Detail CSV: %s", detail_csv)
    logger.info("Detail XLSX: %s", detail_xlsx)
    logger.info("Per-asset CSV: %s", per_asset_csv)
    logger.info("Per-asset XLSX: %s", per_asset_xlsx)
    logger.info("Pivot CSV: %s", pivot_csv)
    logger.info("Pivot XLSX: %s", pivot_xlsx)

    return stats


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extraheer inbreuken voor meest recente RED-keuring per asset en maak provincie-draaitabel."
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_PATH, help="Pad naar enriched input Excel")
    parser.add_argument(
        "--output-prefix",
        type=Path,
        default=None,
        help="Output prefix (zonder extensie). Default: <input>.red-infractions",
    )
    parser.add_argument(
        "--arei-codes",
        type=Path,
        default=None,
        help="Pad naar AREI codes mapping Excel (default: Vincotte/*AREI*.xlsx)",
    )
    parser.add_argument(
        "--since-date",
        default=DEFAULT_SINCE_DATE,
        help="Verwerk enkel inspecties na deze datum (YYYY-MM-DD)",
    )
    return parser.parse_args()


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    args = parse_args()
    input_excel = args.input.resolve()
    if not input_excel.exists():
        logger.error("Inputbestand bestaat niet: %s", input_excel)
        return 1

    output_prefix = (
        args.output_prefix.resolve()
        if args.output_prefix
        else input_excel.with_name(f"{input_excel.stem}.red-infractions")
    )

    if args.arei_codes:
        arei_path = args.arei_codes.resolve()
    else:
        arei_candidates = sorted(Path(__file__).resolve().parent.glob("*AREI*.xlsx"))
        if not arei_candidates:
            logger.error("Geen AREI mappingbestand gevonden in Vincotte-map")
            return 2
        arei_path = arei_candidates[0]

    if not arei_path.exists():
        logger.error("AREI mappingbestand bestaat niet: %s", arei_path)
        return 2

    try:
        stats = extract_red_infractions(
            input_excel=input_excel,
            output_prefix=output_prefix,
            arei_excel=arei_path,
            since_date=args.since_date,
        )
    except Exception as exc:
        logger.error("Extractie mislukt: %s", exc)
        return 3

    logger.info("Klaar. Stats: %s", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

