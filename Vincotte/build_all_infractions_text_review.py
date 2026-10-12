from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import Any

import pandas as pd
from tqdm import tqdm

from extract_red_infractions import (
    DEFAULT_INPUT_DIR,
    DEFAULT_INPUT_FILENAME,
    _extract_infractions_from_text,
    _find_pdf_path,
    _normalize_site_id,
    _normalize_text,
    _read_pdf_text,
    _resolve_column,
)

logger = logging.getLogger(__name__)

DEFAULT_INPUT_PATH = DEFAULT_INPUT_DIR / DEFAULT_INPUT_FILENAME
DEFAULT_SINCE_DATE = "2021-01-01"
DEFAULT_MAPPED_MASTER = DEFAULT_INPUT_DIR / "final_corrected_with_review.xlsx"
DEFAULT_UNMAPPED_MASTER = DEFAULT_INPUT_DIR / "unmapped-ollama-review-template_verwerkt.xlsx"


def _clean(value: Any) -> str:
    text = _normalize_text(value)
    if text.lower() == "nan":
        return ""
    return text


def _norm_text(value: Any) -> str:
    text = _clean(value).lower()
    return re.sub(r"\s+", " ", text)


def _norm_family(value: Any) -> str:
    return _clean(value).upper().replace(" ", "")


def _pick_family(row: pd.Series, candidates: list[str]) -> str:
    for col in candidates:
        if col in row.index:
            fam = _norm_family(row.get(col, ""))
            if fam:
                return fam
    return ""


def _read_master(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Master file not found: {path}")
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path)


def _build_text_lookup(mapped_master: pd.DataFrame, unmapped_master: pd.DataFrame) -> pd.DataFrame:
    mapped = mapped_master.copy()
    mapped["text_norm"] = mapped.get("infraction_text_raw", "").map(_norm_text)
    mapped["code_norm"] = mapped.get("infraction_code_raw", "").map(_clean).str.upper()
    mapped["family_final"] = mapped.apply(
        lambda row: _pick_family(row, ["corrected_family", "verified_family", "infraction_family"]),
        axis=1,
    )
    mapped["master_source"] = "mapped_master"

    unmapped = unmapped_master.copy()
    unmapped["text_norm"] = unmapped.get("infraction_text_raw", "").map(_norm_text)
    unmapped["code_norm"] = unmapped.get("infraction_code_raw", "").map(_clean).str.upper()
    unmapped["family_final"] = unmapped.apply(
        lambda row: _pick_family(
            row,
            [
                "corrected_family",
                "verified_family",
                "infraction_family",
                "ollama_family_suggestion",
            ],
        ),
        axis=1,
    )
    unmapped["master_source"] = "unmapped_master"

    combined = pd.concat(
        [
            mapped[["text_norm", "code_norm", "family_final", "master_source"]],
            unmapped[["text_norm", "code_norm", "family_final", "master_source"]],
        ],
        ignore_index=True,
    )
    combined = combined[(combined["text_norm"] != "") & (combined["family_final"] != "")].copy()

    grouped = (
        combined.groupby("text_norm", dropna=False)
        .agg(
            matched_families=("family_final", lambda s: " | ".join(sorted(dict.fromkeys([v for v in s if v])))),
            matched_codes=("code_norm", lambda s: " | ".join(sorted(dict.fromkeys([v for v in s if v])))),
            matched_sources=("master_source", lambda s: " | ".join(sorted(dict.fromkeys([v for v in s if v])))),
            matched_rows=("family_final", "size"),
            matched_unique_families=("family_final", lambda s: len(set([v for v in s if v]))),
        )
        .reset_index()
    )
    grouped["mapped_in_master_text"] = True
    return grouped


def build_all_infractions_text_review(
    *,
    input_excel: Path,
    mapped_master_path: Path,
    unmapped_master_path: Path,
    output_path: Path,
    since_date: str,
    limit_rows: int | None = None,
) -> dict[str, int]:
    df = pd.read_excel(input_excel)
    mapped_master_df = _read_master(mapped_master_path)
    unmapped_master_df = _read_master(unmapped_master_path)

    cols = {
        "site": _resolve_column(df, ["Site ID", "SiteID"], fallback_index=0),
        "filename": _resolve_column(df, ["FILENAME", "Filename"], fallback_index=11),
        "label": _resolve_column(df, ["Label"]),
        "visit_date": _resolve_column(df, ["Visit Date", "Visit date"]),
        "arango_uuid": _resolve_column(df, ["Arango UUID"]),
        "arango_naam": _resolve_column(df, ["Arango Naam"]),
        "arango_naampad": _resolve_column(df, ["Arango Naampad"]),
        "description": _resolve_column(df, ["Description"], fallback_index=2),
    }

    visit_dates = pd.to_datetime(df[cols["visit_date"]], errors="coerce", dayfirst=True)
    working = df.copy()
    working["__visit_date"] = visit_dates
    working = working[working["__visit_date"].notna() & (working["__visit_date"] >= pd.Timestamp(since_date))].copy()

    if limit_rows is not None and limit_rows > 0:
        working = working.head(limit_rows)

    base_dir = input_excel.parent
    lookup_df = _build_text_lookup(mapped_master_df, unmapped_master_df)

    rows: list[dict[str, Any]] = []
    processed_pdf: dict[str, tuple[list[dict[str, str]], str]] = {}

    for row_idx, row in tqdm(working.iterrows(), total=len(working), desc="Extract all infractions", unit="row"):
        site_id = _normalize_site_id(row.get(cols["site"]))
        filename = _clean(row.get(cols["filename"]))
        pdf_path = _find_pdf_path(base_dir, site_id, filename)
        if pdf_path is None:
            continue

        pdf_key = str(pdf_path)
        if pdf_key in processed_pdf:
            infractions_serialized, pdf_text_hash = processed_pdf[pdf_key]
        else:
            try:
                pdf_text = _read_pdf_text(pdf_path)
            except Exception:
                continue

            infractions = _extract_infractions_from_text(pdf_text)
            infractions_serialized = [
                {
                    "infraction_code_raw": inf.code,
                    "infraction_text_raw": inf.description,
                    "infraction_text_norm": _norm_text(inf.description),
                }
                for inf in infractions
                if _norm_text(inf.description)
            ]
            pdf_text_hash = str(abs(hash(_norm_text(pdf_text))))
            processed_pdf[pdf_key] = (infractions_serialized, pdf_text_hash)

        for inf in infractions_serialized:
            rows.append(
                {
                    "row_index": int(row_idx) + 2,
                    "visit_date": row["__visit_date"].date().isoformat(),
                    "site_id": site_id,
                    "filename": filename,
                    "pdf_path": str(pdf_path),
                    "pdf_content_hash": pdf_text_hash,
                    "label": _clean(row.get(cols["label"])),
                    "arango_uuid": _clean(row.get(cols["arango_uuid"])),
                    "arango_naam": _clean(row.get(cols["arango_naam"])),
                    "arango_naampad": _clean(row.get(cols["arango_naampad"])),
                    "description": _clean(row.get(cols["description"])),
                    "infraction_code_raw": inf["infraction_code_raw"],
                    "infraction_text_raw": inf["infraction_text_raw"],
                    "infraction_text_norm": inf["infraction_text_norm"],
                }
            )

    out_df = pd.DataFrame(rows)
    if out_df.empty:
        raise RuntimeError("No infractions extracted for selected scope.")

    out_df = out_df.merge(
        lookup_df,
        left_on="infraction_text_norm",
        right_on="text_norm",
        how="left",
    )
    if "text_norm" in out_df.columns:
        out_df = out_df.drop(columns=["text_norm"])
    out_df["mapped_in_master_text"] = out_df["mapped_in_master_text"].fillna(False)
    for col in ["matched_families", "matched_codes", "matched_sources"]:
        out_df[col] = out_df[col].fillna("")
    for col in ["matched_rows", "matched_unique_families"]:
        out_df[col] = out_df[col].fillna(0).astype(int)

    out_df["match_status"] = out_df["mapped_in_master_text"].map(lambda v: "MAPPED_BY_TEXT" if v else "NOT_FOUND_BY_TEXT")

    summary = (
        out_df.groupby("match_status", dropna=False)
        .size()
        .reset_index(name="rows")
        .sort_values("rows", ascending=False)
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="all_infractions")
        summary.to_excel(writer, index=False, sheet_name="summary")

    out_df.to_csv(output_path.with_suffix(".csv"), index=False)

    return {
        "rows_input": int(len(working)),
        "pdf_unique_processed": int(len(processed_pdf)),
        "rows_infractions": int(len(out_df)),
        "rows_mapped_by_text": int((out_df["mapped_in_master_text"] == True).sum()),
        "rows_not_found_by_text": int((out_df["mapped_in_master_text"] == False).sum()),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Extract all infractions (since date) and mark whether infraction text exists in the 2 master files."
        )
    )
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_PATH, help="Path to enriched Excel")
    parser.add_argument("--mapped-master", type=Path, default=DEFAULT_MAPPED_MASTER, help="Path to final_corrected_with_review.xlsx")
    parser.add_argument("--unmapped-master", type=Path, default=DEFAULT_UNMAPPED_MASTER, help="Path to unmapped-ollama-review-template_verwerkt.xlsx")
    parser.add_argument("--since-date", default=DEFAULT_SINCE_DATE, help="Only process Visit Date >= this date (YYYY-MM-DD)")
    parser.add_argument("--limit", type=int, default=None, help="Optional row limit for quick tests")
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Output .xlsx path. Default: <input>.all-infractions-text-review.xlsx",
    )
    return parser.parse_args()


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    args = parse_args()
    input_path = args.input.resolve()
    mapped_master = args.mapped_master.resolve()
    unmapped_master = args.unmapped_master.resolve()

    if not input_path.exists():
        candidates = sorted(DEFAULT_INPUT_DIR.glob("*.enriched.xlsx"))
        if candidates:
            input_path = candidates[0].resolve()
            logger.info("Geen bestaande input op --input/default, gebruik fallback: %s", input_path)
        else:
            logger.error("Inputbestand bestaat niet: %s", input_path)
            return 1
    if not mapped_master.exists() or not unmapped_master.exists():
        logger.error("Masterbestand ontbreekt: %s / %s", mapped_master, unmapped_master)
        return 2

    output_path = (
        args.output.resolve()
        if args.output is not None
        else input_path.with_name(f"{input_path.stem}.all-infractions-text-review.xlsx")
    )

    try:
        stats = build_all_infractions_text_review(
            input_excel=input_path,
            mapped_master_path=mapped_master,
            unmapped_master_path=unmapped_master,
            output_path=output_path,
            since_date=args.since_date,
            limit_rows=args.limit,
        )
    except Exception as exc:
        logger.error("Build failed: %s", exc)
        return 3

    logger.info("Output XLSX: %s", output_path)
    logger.info("Output CSV: %s", output_path.with_suffix('.csv'))
    logger.info("Stats: %s", stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())



