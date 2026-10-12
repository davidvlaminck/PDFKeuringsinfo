#!/usr/bin/env python3
"""Combine the three infraction sources into one text-first mapping table.

Output keeps one row per normalized infraction text so it can be used for
future text-based matching.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any

import pandas as pd

FAMILY_RE = re.compile(r"F\d+(?:-\d+)?")


def _clean(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _norm_text(value: Any) -> str:
    return " ".join(_clean(value).lower().split())


def _extract_families(value: Any) -> list[str]:
    text = _clean(value).upper()
    if not text:
        return []
    seen: set[str] = set()
    out: list[str] = []
    for fam in FAMILY_RE.findall(text):
        if fam not in seen:
            seen.add(fam)
            out.append(fam)
    return out


def _pick_family_with_rank(row: pd.Series, candidates: list[tuple[str, int]]) -> tuple[str, int, str]:
    for col, rank in candidates:
        fams = _extract_families(row.get(col, ""))
        if fams:
            return ";".join(fams), rank, col
    return "", 99, ""


def _to_int_flag(value: Any) -> int:
    try:
        return int(float(value) == 1.0)
    except Exception:
        return 0


def _series_or_empty(df: pd.DataFrame, column: str) -> pd.Series:
    if column in df.columns:
        return df[column]
    return pd.Series(["" for _ in range(len(df))], index=df.index)


def _first_non_empty(series: pd.Series) -> str:
    for value in series:
        txt = _clean(value)
        if txt:
            return txt
    return ""


def _build_filename_lookup(all_infractions_file: Path | None) -> dict[str, str]:
    if all_infractions_file is None or not all_infractions_file.exists():
        return {}

    try:
        df = pd.read_excel(all_infractions_file, sheet_name="all_infractions")
    except Exception:
        df = pd.read_excel(all_infractions_file, sheet_name=0)

    if df.empty:
        return {}

    if "infraction_text_norm" in df.columns:
        text_norm = df["infraction_text_norm"].map(_norm_text)
    else:
        text_norm = _series_or_empty(df, "infraction_text_raw").map(_norm_text)

    filename_series = _series_or_empty(df, "filename").map(_clean)
    lookup_df = pd.DataFrame({"text_norm": text_norm, "filename": filename_series})
    lookup_df = lookup_df[(lookup_df["text_norm"] != "") & (lookup_df["filename"] != "")]
    if lookup_df.empty:
        return {}

    grouped = lookup_df.groupby("text_norm", dropna=False)["filename"].agg(_first_non_empty)
    return grouped.to_dict()


def _build_source_rows(df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    if source_name == "mapped_master":
        family_candidates = [
            ("corrected_family", 0),
            ("verified_family", 1),
            ("infraction_family", 2),
        ]
        source_rank = 0
        occurrences_col = "occurrences"
    elif source_name == "unmapped_master":
        family_candidates = [
            ("corrected_family", 0),
            ("verified_family", 1),
            ("infraction_family", 2),
            ("ollama_family_suggestion", 4),
        ]
        source_rank = 1
        occurrences_col = "occurrences"
    elif source_name == "suggestions":
        family_candidates = [
            ("corrected_family", 0),
            ("master_combined_families", 3),
            ("ollama_family_suggestion", 4),
        ]
        source_rank = 2
        occurrences_col = "occurrence_count"
    else:
        raise ValueError(f"Unsupported source_name: {source_name}")

    working = df.copy()
    working["infraction_text_raw"] = _series_or_empty(working, "infraction_text_raw").map(_clean)
    working["text_norm"] = working["infraction_text_raw"].map(_norm_text)
    working["infraction_code_raw"] = _series_or_empty(working, "infraction_code_raw").map(_clean).str.upper()
    working["sample_filename"] = _series_or_empty(working, "sample_filename").map(_clean)
    working["verified"] = _series_or_empty(working, "verified").map(_to_int_flag)
    working["occurrences"] = pd.to_numeric(_series_or_empty(working, occurrences_col), errors="coerce").fillna(0).astype(int)

    family_data = working.apply(lambda row: _pick_family_with_rank(row, family_candidates), axis=1)
    working["family"] = [item[0] for item in family_data]
    working["family_rank"] = [item[1] for item in family_data]
    working["family_from_col"] = [item[2] for item in family_data]
    working["source"] = source_name
    working["source_rank"] = source_rank

    return working[
        [
            "source",
            "source_rank",
            "text_norm",
            "infraction_text_raw",
            "infraction_code_raw",
            "sample_filename",
            "occurrences",
            "family",
            "family_rank",
            "family_from_col",
            "verified",
        ]
    ]


def build_combined_mapping(
    *,
    mapped_master: Path,
    unmapped_master: Path,
    suggestions_file: Path,
    all_infractions_file: Path | None,
    output_xlsx: Path,
) -> dict[str, int]:
    mapped_df = pd.read_excel(mapped_master, sheet_name=0)
    unmapped_df = pd.read_excel(unmapped_master, sheet_name=0)
    suggestions_df = pd.read_excel(suggestions_file, sheet_name="infractions_with_suggestions")
    filename_lookup = _build_filename_lookup(all_infractions_file)

    combined = pd.concat(
        [
            _build_source_rows(mapped_df, "mapped_master"),
            _build_source_rows(unmapped_df, "unmapped_master"),
            _build_source_rows(suggestions_df, "suggestions"),
        ],
        ignore_index=True,
    )
    combined = combined[combined["text_norm"] != ""].copy()

    rows: list[dict[str, Any]] = []
    for text_norm, group in combined.groupby("text_norm", dropna=False):
        ranked_group = group.sort_values(["family_rank", "source_rank"], ascending=[True, True])

        family_row = ranked_group[ranked_group["family"] != ""].head(1)
        if family_row.empty:
            family = ""
            family_source = ""
        else:
            family = str(family_row.iloc[0]["family"])
            family_source = f"{family_row.iloc[0]['source']}:{family_row.iloc[0]['family_from_col']}"

        preferred_text = (
            ranked_group[ranked_group["infraction_text_raw"] != ""]["infraction_text_raw"].head(1).tolist()
            or group["infraction_text_raw"].dropna().astype(str).tolist()
        )
        infraction_text_raw = preferred_text[0] if preferred_text else ""

        codes = [c for c in group["infraction_code_raw"].astype(str).tolist() if c and c != "NAN"]
        unique_codes = list(dict.fromkeys(codes))
        infraction_code_raw = ";".join(unique_codes)

        filenames = [f for f in group["sample_filename"].astype(str).tolist() if f and f != "NAN"]
        sample_filename = filenames[0] if filenames else filename_lookup.get(text_norm, "")

        suggestion_occ = group.loc[group["source"] == "suggestions", "occurrences"]
        if not suggestion_occ.empty:
            occurrences = int(suggestion_occ.max())
        else:
            occurrences = int(group["occurrences"].max())

        verified = int(group["verified"].max())

        rows.append(
            {
                "infraction_text_raw": infraction_text_raw,
                "infraction_code_raw": infraction_code_raw,
                "sample_filename": sample_filename,
                "occurrences": occurrences,
                "family": family,
                "verified": verified,
                "text_norm": text_norm,
                "family_source": family_source,
                "sources_present": ";".join(sorted(group["source"].unique().tolist())),
                "source_rows": int(len(group)),
            }
        )

    result = pd.DataFrame(rows).sort_values(["occurrences", "infraction_text_raw"], ascending=[False, True])

    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    output_csv = output_xlsx.with_suffix(".csv")

    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        result.to_excel(writer, sheet_name="combined_text_mapping", index=False)

    result.to_csv(output_csv, index=False)

    return {
        "unique_texts": int(len(result)),
        "verified_eq_1": int((result["verified"] == 1).sum()),
        "with_family": int((result["family"] != "").sum()),
    }


def parse_args() -> argparse.Namespace:
    base = Path("/home/davidlinux/Documenten/AWV/Keuringen")
    parser = argparse.ArgumentParser(description="Combine the 3 infraction sources into one text-first mapping table.")
    parser.add_argument("--mapped-master", type=Path, default=base / "final_corrected_with_review.xlsx")
    parser.add_argument("--unmapped-master", type=Path, default=base / "unmapped-ollama-review-template_verwerkt.xlsx")
    parser.add_argument(
        "--suggestions",
        type=Path,
        default=base / "MyVinotte Min.Vl.Gemeenschap Reports list.all-infractions-with-ollama-suggestions.xlsx",
    )
    parser.add_argument(
        "--all-infractions",
        type=Path,
        default=base / "MyVinotte Min.Vl.Gemeenschap Reports list.all-infractions-text-review.xlsx",
        help="Optional source to backfill missing sample_filename by infraction text.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=base / "MyVinotte Min.Vl.Gemeenschap Reports list.combined-text-family-mapping.xlsx",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    stats = build_combined_mapping(
        mapped_master=args.mapped_master,
        unmapped_master=args.unmapped_master,
        suggestions_file=args.suggestions,
        all_infractions_file=args.all_infractions,
        output_xlsx=args.output,
    )
    print(f"Saved: {args.output}")
    print(f"Saved: {args.output.with_suffix('.csv')}")
    print(stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())



