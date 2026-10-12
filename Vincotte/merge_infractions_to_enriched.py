from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd

FAMILY_RE = re.compile(r"\bF\d+(?:-\d+)?\b", re.IGNORECASE)


def _read_table(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path)
    raise ValueError(f"Unsupported file type for {path}")


def _clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    return re.sub(r"\s+", " ", text)


def _norm_code(value: object) -> str:
    return _clean_text(value).upper()


def _norm_text(value: object) -> str:
    return _clean_text(value).lower()


def _key(code: object, text: object) -> str:
    return f"{_norm_code(code)}||{_norm_text(text)}"


def _norm_family(value: object) -> str:
    token = _clean_text(value).upper().replace(" ", "")
    return token


def _to_int_str(value: object) -> str:
    token = _clean_text(value)
    if not token:
        return ""
    try:
        f = float(token)
        if f.is_integer():
            return str(int(f))
    except ValueError:
        pass
    return token


def _extract_arei_sheets(arei_path: Path) -> tuple[pd.DataFrame, set[str]]:
    xls = pd.ExcelFile(arei_path)
    combined: list[pd.DataFrame] = []
    families: set[str] = set()

    for sheet in xls.sheet_names:
        df = pd.read_excel(arei_path, sheet_name=sheet)
        df_sheet = df.copy()
        df_sheet.insert(0, "source_sheet", sheet)
        combined.append(df_sheet)

        # Detect family tokens anywhere in the sheet values.
        series = df.astype(str).stack()
        for cell in series.tolist():
            for hit in FAMILY_RE.findall(str(cell)):
                families.add(hit.upper())

    return pd.concat(combined, ignore_index=True), families


def _build_master_mapping(
    mapped_master: pd.DataFrame,
    unmapped_master: pd.DataFrame,
) -> pd.DataFrame:
    mm = mapped_master.copy()
    um = unmapped_master.copy()

    # MAPPED master precedence: corrected_family > verified_family > infraction_family
    mm["family_final"] = (
        mm.get("corrected_family", "")
        .map(_norm_family)
        .where(mm.get("corrected_family", "").map(_clean_text) != "", mm.get("verified_family", "").map(_norm_family))
        .where(
            lambda s: s != "",
            mm.get("infraction_family", "").map(_norm_family),
        )
    )
    mm_map = pd.DataFrame(
        {
            "infraction_code_raw": mm.get("infraction_code_raw", "").map(_norm_code),
            "infraction_text_raw": mm.get("infraction_text_raw", "").map(_norm_text),
            "family_final": mm["family_final"],
            "mapping_source": "mapped_master",
        }
    )

    # UNMAPPED reviewed master precedence: verified_family > ollama_family_suggestion
    verified_flag = um.get("verified", "").map(_to_int_str).isin({"1"})
    um = um.loc[verified_flag].copy()
    um["family_final"] = (
        um.get("verified_family", "")
        .map(_norm_family)
        .where(um.get("verified_family", "").map(_clean_text) != "", um.get("ollama_family_suggestion", "").map(_norm_family))
    )
    um_map = pd.DataFrame(
        {
            "infraction_code_raw": um.get("infraction_code_raw", "").map(_norm_code),
            "infraction_text_raw": um.get("infraction_text_raw", "").map(_norm_text),
            "family_final": um["family_final"],
            "mapping_source": "unmapped_master_verified",
        }
    )

    mapping_df = pd.concat([mm_map, um_map], ignore_index=True)
    mapping_df = mapping_df[
        (mapping_df["infraction_code_raw"] != "")
        & (mapping_df["infraction_text_raw"] != "")
        & (mapping_df["family_final"] != "")
    ].copy()

    mapping_df["map_key"] = mapping_df.apply(
        lambda r: _key(r["infraction_code_raw"], r["infraction_text_raw"]), axis=1
    )

    # If duplicate keys exist, prefer reviewed-unmapped mapping over mapped master.
    priority = {"unmapped_master_verified": 0, "mapped_master": 1}
    mapping_df["prio"] = mapping_df["mapping_source"].map(priority).fillna(9)
    mapping_df = mapping_df.sort_values(["map_key", "prio"]).drop_duplicates(subset=["map_key"], keep="first")

    return mapping_df[["map_key", "family_final", "mapping_source"]]


def _sanitize_family_for_col(family: str) -> str:
    token = re.sub(r"[^A-Z0-9]+", "_", family.upper()).strip("_")
    return token or "UNKNOWN"


def merge_infractions_to_enriched(
    *,
    enriched_path: Path,
    detail_path: Path,
    mapped_master_path: Path,
    unmapped_master_path: Path,
    arei_path: Path,
    output_path: Path,
) -> dict[str, int]:
    enriched_df = pd.read_excel(enriched_path)
    detail_df = _read_table(detail_path)
    mapped_master_df = pd.read_excel(mapped_master_path)
    unmapped_master_df = pd.read_excel(unmapped_master_path)

    arei_combined_df, arei_families = _extract_arei_sheets(arei_path)

    mapping_df = _build_master_mapping(mapped_master_df, unmapped_master_df)
    map_lookup = dict(zip(mapping_df["map_key"], mapping_df["family_final"]))

    d = detail_df.copy()
    required = {"arango_uuid", "filename", "infraction_code_raw", "infraction_text_raw", "infraction_family"}
    missing = [c for c in required if c not in d.columns]
    if missing:
        raise ValueError(f"Detail file missing columns: {missing}")

    d["code_norm"] = d["infraction_code_raw"].map(_norm_code)
    d["text_norm"] = d["infraction_text_raw"].map(_norm_text)
    d["map_key"] = d.apply(lambda r: _key(r["code_norm"], r["text_norm"]), axis=1)

    d["family_from_master"] = d["map_key"].map(map_lookup)
    d["family_original"] = d["infraction_family"].map(_norm_family)
    d["infraction_family_final"] = d["family_from_master"].where(
        d["family_from_master"].map(_clean_text) != "",
        d["family_original"],
    )
    d["infraction_family_final"] = d["infraction_family_final"].where(
        d["infraction_family_final"].map(_clean_text) != "",
        "UNMAPPED",
    )

    # Aggregate counts per asset (Arango UUID) and family.
    agg = (
        d.groupby(["arango_uuid", "infraction_family_final"], dropna=False)
        .size()
        .reset_index(name="count")
    )

    all_families = set(d["infraction_family_final"].dropna().map(_norm_family).tolist())
    all_families |= set(mapping_df["family_final"].dropna().map(_norm_family).tolist())
    all_families |= set(arei_families)
    all_families.discard("")

    pivot = agg.pivot(index="arango_uuid", columns="infraction_family_final", values="count").fillna(0)
    pivot = pivot.astype(int)

    # Guarantee one column per known category/family (incl. custom families).
    for fam in sorted(all_families):
        if fam not in pivot.columns:
            pivot[fam] = 0

    pivot = pivot.reindex(sorted(pivot.columns), axis=1)

    col_map: dict[str, str] = {}
    for fam in pivot.columns:
        col_map[fam] = f"infractions_{_sanitize_family_for_col(fam)}_count"

    pivot = pivot.rename(columns=col_map)
    pivot["infractions_total"] = pivot.sum(axis=1)

    # Merge onto enriched by Arango UUID.
    out = enriched_df.copy()
    out["Arango UUID"] = out.get("Arango UUID", "").map(_clean_text)
    pivot_reset = pivot.reset_index().rename(columns={"arango_uuid": "Arango UUID"})
    pivot_reset["Arango UUID"] = pivot_reset["Arango UUID"].map(_clean_text)

    out = out.merge(pivot_reset, on="Arango UUID", how="left")

    count_cols = [c for c in out.columns if c.startswith("infractions_") and c.endswith("_count")]
    count_cols += [c for c in ["infractions_total"] if c in out.columns]
    for c in count_cols:
        out[c] = out[c].fillna(0).astype(int)

    # Helpful trace columns for business QA.
    out["mapping_master_used"] = "final_corrected_with_review + unmapped-ollama-review-template_verwerkt"
    out["mapping_rows_in_detail"] = len(d)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="enriched_with_infractions")
        d[
            [
                "arango_uuid",
                "filename",
                "infraction_code_raw",
                "infraction_text_raw",
                "family_original",
                "family_from_master",
                "infraction_family_final",
            ]
        ].to_excel(writer, index=False, sheet_name="infractions_mapped_detail")
        arei_combined_df.to_excel(writer, index=False, sheet_name="AREI_codes_reference")

    out.to_csv(output_path.with_suffix(".csv"), index=False)

    return {
        "rows_enriched": len(out),
        "rows_detail": len(d),
        "rows_mapping_keys": len(mapping_df),
        "family_columns": len([c for c in out.columns if c.startswith("infractions_") and c.endswith("_count")]),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Merge MASTER infraction mappings into enriched Excel and add one column per infraction family."
    )
    parser.add_argument("--enriched", type=Path, required=True, help="Input enriched Excel (preferably with Step 2b already merged).")
    parser.add_argument("--detail", type=Path, required=True, help="Infractions detail file (.xlsx or .csv).")
    parser.add_argument("--mapped-master", type=Path, required=True, help="Path to final_corrected_with_review.xlsx")
    parser.add_argument("--unmapped-master", type=Path, required=True, help="Path to unmapped-ollama-review-template_verwerkt.xlsx")
    parser.add_argument("--arei-codes", type=Path, required=True, help="Path to AREI Codes workbook to copy into extra sheet.")
    parser.add_argument("--output", type=Path, default=None, help="Output .xlsx path. Default: <enriched>.with-infractions.xlsx")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    enriched = args.enriched.resolve()
    detail = args.detail.resolve()
    mapped_master = args.mapped_master.resolve()
    unmapped_master = args.unmapped_master.resolve()
    arei = args.arei_codes.resolve()

    for p in [enriched, detail, mapped_master, unmapped_master, arei]:
        if not p.exists():
            print(f"ERROR: File not found: {p}")
            return 1

    output = (
        args.output.resolve()
        if args.output is not None
        else enriched.with_name(f"{enriched.stem}.with-infractions.xlsx")
    )

    try:
        stats = merge_infractions_to_enriched(
            enriched_path=enriched,
            detail_path=detail,
            mapped_master_path=mapped_master,
            unmapped_master_path=unmapped_master,
            arei_path=arei,
            output_path=output,
        )
    except Exception as exc:
        print(f"ERROR: Merge failed: {exc}")
        return 2

    print(f"Merged output: {output}")
    print(f"CSV output: {output.with_suffix('.csv')}")
    print(f"Stats: {stats}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())



