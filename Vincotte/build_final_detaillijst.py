#!/usr/bin/env python3
"""Build final per-asset detail list using the combined text-family mapping.

The combined mapping file (combined-text-family-mapping.xlsx) is the single
source of truth for infraction text → family assignments.  It already contains
a resolved ``family`` column and a ``verified`` flag, so no priority logic is
needed here.

Steps:
1. Load the combined mapping (sheet: combined_text_mapping).
2. Load all extracted infractions (sheet: all_infractions).
3. Join on normalised infraction text (text_norm).
4. Keep only the most recent inspection per asset.
5. Aggregate per asset and write output Excel.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd


# Known one-off visit date corrections.
VISIT_DATE_OVERRIDES = {
    "arango_uuid": {
        "1d163dee-5ad4-44bb-ad75-63a40dd4b32a": "2009-09-26",
    },
    "filename": {
        "WRK000000000000000001351961700003NL.pdf": "2009-09-26",
    },
}


def _norm_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).strip().lower().split())


def _clean(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _extract_families(value: object) -> list[str]:
    text = _clean(value).upper()
    if not text:
        return []
    families = re.findall(r"F\d+(?:-\d+)?", text)
    seen: set[str] = set()
    out: list[str] = []
    for fam in families:
        if fam not in seen:
            seen.add(fam)
            out.append(fam)
    return out


def _join_unique(series: pd.Series, sep: str = " | ") -> str:
    seen: set[str] = set()
    out: list[str] = []
    for val in series:
        txt = _clean(val)
        if txt and txt not in seen:
            seen.add(txt)
            out.append(txt)
    return sep.join(out)


def build_detaillijst(
    combined_mapping_file: Path,
    all_infractions_file: Path,
    output_file: Path,
) -> None:
    # ------------------------------------------------------------------
    # 1. Load combined mapping — single source of truth for text→family.
    # ------------------------------------------------------------------
    mapping_df = pd.read_excel(combined_mapping_file, sheet_name="combined_text_mapping")
    mapping_df["text_norm"] = mapping_df["text_norm"].map(_norm_text)

    # Normalise family column: strip, uppercase, keep as-is (may be semicolon-separated).
    mapping_df["family"] = mapping_df["family"].fillna("").astype(str).str.strip()
    mapping_df["verified"] = pd.to_numeric(mapping_df.get("verified", 1), errors="coerce").fillna(1)

    # Keep only the columns we need for the join.
    map_cols = [c for c in [
        "text_norm",
        "family",
        "family_source",
        "verified",
        "infraction_code_raw",
        "occurrences",
        "sources_present",
    ] if c in mapping_df.columns]

    mapping_slim = mapping_df[map_cols].drop_duplicates(subset=["text_norm"], keep="first")

    # ------------------------------------------------------------------
    # 2. Load all extracted infractions.
    # ------------------------------------------------------------------
    all_df = pd.read_excel(all_infractions_file, sheet_name="all_infractions")
    all_df["infraction_text_norm"] = (
        all_df.get("infraction_text_norm", all_df["infraction_text_raw"])
        .map(_norm_text)
    )

    # ------------------------------------------------------------------
    # 3. Join mapping onto infractions.
    # ------------------------------------------------------------------
    enriched = all_df.merge(
        mapping_slim,
        left_on="infraction_text_norm",
        right_on="text_norm",
        how="left",
    )

    # Use ``family`` as the resolved final family (already prioritised in combined mapping).
    enriched["final_family"] = enriched["family"].fillna("UNKNOWN")
    enriched["final_family_source"] = enriched.get("family_source", pd.Series([""] * len(enriched), index=enriched.index)).fillna("")

    # ------------------------------------------------------------------
    # 4. Apply visit-date corrections.
    # ------------------------------------------------------------------
    for key_col, overrides in VISIT_DATE_OVERRIDES.items():
        if key_col in enriched.columns:
            key_values = enriched[key_col].astype(str)
            mask = key_values.isin(overrides)
            if mask.any():
                enriched.loc[mask, "visit_date"] = key_values.loc[mask].map(overrides)

    enriched["visit_date"] = pd.to_datetime(enriched["visit_date"], errors="coerce", dayfirst=True)

    # ------------------------------------------------------------------
    # 5. Determine asset key and filter to rows with valid asset + date.
    # ------------------------------------------------------------------
    if "arango_uuid" in enriched.columns:
        asset_key = "arango_uuid"
    elif "site_id" in enriched.columns:
        asset_key = "site_id"
    else:
        raise ValueError("No asset key found (expected arango_uuid or site_id).")

    valid_asset_mask = enriched[asset_key].notna() & (enriched[asset_key].astype(str).str.strip() != "")
    valid_date_mask = enriched["visit_date"].notna()
    working = enriched[valid_asset_mask & valid_date_mask].copy()

    latest_per_asset = working.groupby(asset_key, as_index=False)["visit_date"].max()
    latest_rows = working.merge(latest_per_asset, on=[asset_key, "visit_date"], how="inner")

    # ------------------------------------------------------------------
    # 6. Aggregate per asset.
    # ------------------------------------------------------------------
    metadata_cols = [
        col
        for col in ["site_id", "arango_naam", "arango_naampad", "description", "filename", "pdf_path"]
        if col in latest_rows.columns
    ]

    def _first_non_empty(group: pd.DataFrame, column: str) -> str:
        for value in group[column]:
            txt = _clean(value)
            if txt:
                return txt
        return ""

    def agg_asset(group: pd.DataFrame) -> pd.Series:
        categories: list[str] = []
        for val in group["final_family"].fillna(""):
            categories.extend(_extract_families(val))
        seen: set[str] = set()
        categories_unique: list[str] = []
        for fam in categories:
            if fam not in seen:
                seen.add(fam)
                categories_unique.append(fam)

        full_texts = _join_unique(group["infraction_text_raw"], sep="\n\n")
        result: dict = {
            "visit_date": pd.to_datetime(group["visit_date"], errors="coerce").max(),
            "infraction_count_latest": int(len(group)),
            "categories_latest": ";".join(categories_unique),
            "categories_count_latest": int(len(categories_unique)),
            "full_infraction_texts_latest": full_texts,
            "families_sources_used": _join_unique(group["final_family_source"], sep=";"),
        }
        for col in metadata_cols:
            result[col] = _first_non_empty(group, col)
        return pd.Series(result)

    aggregated = latest_rows.groupby(asset_key, dropna=False).apply(agg_asset).reset_index()
    aggregated = aggregated.sort_values(["visit_date", asset_key], ascending=[False, True])

    # ------------------------------------------------------------------
    # 7. Per-category count columns (pivot).
    # ------------------------------------------------------------------
    exploded = latest_rows[[asset_key, "final_family"]].copy()
    exploded["final_family"] = exploded["final_family"].fillna("").astype(str)
    exploded = exploded.assign(family=exploded["final_family"].str.split(";")).explode("family")
    exploded["family"] = exploded["family"].str.strip().str.upper()
    exploded = exploded[exploded["family"].str.match(r"^F\d+(?:-\d+)?$", na=False)]
    if not exploded.empty:
        pivot = (
            exploded.groupby([asset_key, "family"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )
        family_cols = sorted(
            [c for c in pivot.columns if c != asset_key],
            key=lambda x: [int(p) for p in re.findall(r"\d+", x)],
        )
        pivot = pivot[[asset_key] + family_cols]
        pivot.columns = [asset_key] + [f"cat_{c}" for c in family_cols]
        aggregated = aggregated.merge(pivot, on=asset_key, how="left")
        cat_cols = [c for c in aggregated.columns if c.startswith("cat_")]
        aggregated[cat_cols] = aggregated[cat_cols].fillna(0).astype(int)

    # ------------------------------------------------------------------
    # 8. Detail sheet (one row per infraction in latest inspection).
    # ------------------------------------------------------------------
    detail_cols = [
        asset_key,
        "visit_date",
        "site_id",
        "arango_naam",
        "arango_naampad",
        "description",
        "filename",
        "infraction_code_raw",
        "infraction_text_raw",
        "final_family",
        "final_family_source",
        "verified",
        "occurrences",
        "sources_present",
    ]
    detail_cols = [c for c in detail_cols if c in latest_rows.columns]
    latest_detail = latest_rows[detail_cols].sort_values(
        ["visit_date", asset_key], ascending=[False, True]
    )

    # ------------------------------------------------------------------
    # 9. Summary metrics.
    # ------------------------------------------------------------------
    texts_with_family = int((mapping_slim["family"] != "").sum())
    texts_verified = int((mapping_slim["verified"] == 1).sum())

    summary = pd.DataFrame(
        {
            "metric": [
                "total_unique_texts_in_mapping",
                "texts_with_family",
                "texts_verified",
                "latest_asset_count",
                "latest_asset_infraction_rows",
            ],
            "value": [
                len(mapping_slim),
                texts_with_family,
                texts_verified,
                int(aggregated[asset_key].nunique()),
                int(len(latest_detail)),
            ],
        }
    )

    # ------------------------------------------------------------------
    # 10. Write output.
    # ------------------------------------------------------------------
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        mapping_slim.to_excel(writer, sheet_name="combined_text_mapping", index=False)
        aggregated.to_excel(writer, sheet_name="latest_asset_summary", index=False)
        latest_detail.to_excel(writer, sheet_name="latest_asset_infractions", index=False)
        summary.to_excel(writer, sheet_name="summary", index=False)

    csv_out = output_file.with_suffix(".latest_asset_summary.csv")
    aggregated.to_csv(csv_out, index=False)

    print(f"Saved: {output_file}")
    print(f"Saved: {csv_out}")
    print(summary.to_string(index=False))


def parse_args() -> argparse.Namespace:
    base = Path("/home/davidlinux/Documenten/AWV/Keuringen")
    parser = argparse.ArgumentParser(
        description="Build final detail list using the combined text-family mapping."
    )
    parser.add_argument(
        "--combined-mapping",
        type=Path,
        default=base / "MyVinotte Min.Vl.Gemeenschap Reports list.combined-text-family-mapping.xlsx",
        help="Path to combined-text-family-mapping.xlsx (sheet: combined_text_mapping)",
    )
    parser.add_argument(
        "--all-infractions",
        type=Path,
        default=base / "MyVinotte Min.Vl.Gemeenschap Reports list.all-infractions-text-review.xlsx",
        help="Path to all-infractions-text-review.xlsx (sheet: all_infractions)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=base / "MyVinotte Min.Vl.Gemeenschap Reports list.detaillijst-latest-keuring.xlsx",
        help="Output Excel path",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    build_detaillijst(
        combined_mapping_file=args.combined_mapping,
        all_infractions_file=args.all_infractions,
        output_file=args.output,
    )

